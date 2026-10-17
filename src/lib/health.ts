import "server-only";

import type { GraphHealthReport, HealthProbe, ProbeStatus } from "@/types/meeting";
import { getEnvVar, getSummaryEmailRecipient, getTargetUserId } from "@/lib/env";
import { GraphError } from "@/lib/errors";
import {
  acquireAccessToken,
  decodeAccessTokenRoles,
  getRecentMeetings,
  listRecordings,
  listTranscripts,
  resetGraphAuthCache,
  resolveOnlineMeetingId,
  tokenHasChatRole,
} from "@/lib/m365";

const REQUIRED_ROLES = [
  "Calendars.Read",
  "OnlineMeetings.Read.All",
  "OnlineMeetingTranscript.Read.All",
  "OnlineMeetingRecording.Read.All",
  "User.Read.All",
  "Mail.Send",
  "Chat.Read.All",
  "ChatMessage.Read.All",
  "Tasks.ReadWrite.All",
] as const;

interface GraphErrorPayload {
  error?: { code?: string; message?: string };
}

interface OnlineMeetingChatInfo {
  chatInfo?: { threadId?: string };
}

function probe(status: ProbeStatus, detail: string): HealthProbe {
  return { status, detail };
}

function missingRequiredRoles(roles: string[]): string[] {
  const hasChat = tokenHasChatRole(roles);
  return REQUIRED_ROLES.filter((role) => {
    if (role === "Chat.Read.All" || role === "ChatMessage.Read.All") {
      return !hasChat;
    }
    return !roles.includes(role);
  });
}

async function graphGet(
  token: string,
  path: string,
): Promise<{
  ok: boolean;
  status: number;
  code: string | null;
  message: string | null;
}> {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const body = await response.text();
  let code: string | null = null;
  let message: string | null = null;
  try {
    const parsed = JSON.parse(body) as GraphErrorPayload;
    code = parsed.error?.code ?? null;
    message = parsed.error?.message ?? null;
  } catch {
    code = null;
  }
  return { ok: response.ok, status: response.status, code, message };
}

function classifyGraphProbe(
  result: {
    ok: boolean;
    status: number;
    code: string | null;
    message: string | null;
  },
  okDetail: string,
  missingRole: boolean,
): HealthProbe {
  if (result.ok) {
    return probe("ok", okDetail);
  }
  if (result.code === "GraphAccessToTranscriptsDisabled") {
    return probe(
      "denied",
      "Tenant control EnableGraphTranscriptAccess is still off for Microsoft Graph.",
    );
  }
  if (result.status === 403 || result.status === 401) {
    const graphDetail = result.message ?? result.code ?? `Graph returned ${result.status}.`;
    return probe(
      "denied",
      missingRole ? "Permission is missing from the app token." : graphDetail,
    );
  }
  if (result.status === 404) {
    return probe("unknown", "Graph returned 404 for this probe.");
  }
  return probe("failed", result.code ?? `Graph returned ${result.status}.`);
}


async function probePerMeetingArtifact(
  userId: string,
  meetingId: string,
  kind: "transcripts" | "recordings",
  hasRole: boolean,
): Promise<HealthProbe> {
  try {
    const items =
      kind === "transcripts"
        ? await listTranscripts(userId, meetingId)
        : await listRecordings(userId, meetingId);
    const noun = kind === "transcripts" ? "transcripts" : "recordings";
    return probe(
      "ok",
      items.length > 0
        ? `Graph listed ${items.length} ${noun} on a recent meeting.`
        : `Graph allowed listing ${noun} on a recent meeting (none on that meeting).`,
    );
  } catch (error) {
    if (error instanceof GraphError && error.code === "TRANSCRIPTS_DISABLED") {
      return probe(
        "denied",
        "Tenant control EnableGraphTranscriptAccess is still off for Microsoft Graph.",
      );
    }
    if (!hasRole) {
      return probe("denied", "Permission is missing from the app token.");
    }
    const message = error instanceof Error ? error.message : "Graph probe failed.";
    return probe("denied", message);
  }
}

async function probeMeetingChat(
  token: string,
  encodedUser: string,
  meetingId: string,
  missingChatRole: boolean,
): Promise<HealthProbe> {
  const meetingResponse = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodedUser}/onlineMeetings/${encodeURIComponent(meetingId)}?$select=id,chatInfo`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (meetingResponse.ok) {
    const payload = (await meetingResponse.json()) as OnlineMeetingChatInfo;
    const threadId = payload.chatInfo?.threadId;
    if (threadId) {
      return classifyGraphProbe(
        await graphGet(
          token,
          `/chats/${encodeURIComponent(threadId)}/messages?$top=1`,
        ),
        "Graph can read this meeting's chat thread.",
        missingChatRole,
      );
    }
  }

  return classifyGraphProbe(
    await graphGet(token, `/users/${encodedUser}/chats?$top=1`),
    "Graph can list this user's Teams chats.",
    missingChatRole,
  );
}

export async function getGraphHealthReport(
  options?: { refresh?: boolean },
): Promise<GraphHealthReport> {
  const targetUser = getTargetUserId();
  const summaryEmailRecipient = getSummaryEmailRecipient();
  const clientId = getEnvVar("AZURE_CLIENT_ID");

  try {
    if (options?.refresh) {
      resetGraphAuthCache();
    }

    const token = await acquireAccessToken({ skipCache: true });
    const roles = decodeAccessTokenRoles(token);
    const missingRoles = missingRequiredRoles(roles);
    const encodedUser = encodeURIComponent(targetUser);
    const hasChat = tokenHasChatRole(roles);

    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 2);
    const calendarPath =
      `/users/${encodedUser}/calendarView?startDateTime=${encodeURIComponent(start.toISOString())}` +
      `&endDateTime=${encodeURIComponent(end.toISOString())}&$top=1`;

    const [calendar, todo] = await Promise.all([
      graphGet(token, calendarPath),
      graphGet(token, `/users/${encodedUser}/todo/lists?$top=1`),
    ]);

    let meetingId: string | null = null;
    let organizerId = targetUser;
    try {
      const meetings = await getRecentMeetings(targetUser, 30);
      const sample = meetings[0];
      if (sample) {
        organizerId = sample.organizerId;
        meetingId =
          sample.onlineMeetingId ??
          (await resolveOnlineMeetingId(sample.organizerId, sample.joinUrl));
      }
    } catch {
      meetingId = null;
    }

    const noMeetingDetail =
      "No recent online meeting to probe. Token roles are shown below.";
    const encodedOrganizer = encodeURIComponent(organizerId);

    const [transcripts, recordings, chat] = meetingId
      ? await Promise.all([
          probePerMeetingArtifact(
            organizerId,
            meetingId,
            "transcripts",
            roles.includes("OnlineMeetingTranscript.Read.All"),
          ),
          probePerMeetingArtifact(
            organizerId,
            meetingId,
            "recordings",
            roles.includes("OnlineMeetingRecording.Read.All"),
          ),
          probeMeetingChat(token, encodedOrganizer, meetingId, !hasChat),
        ])
      : [
          roles.includes("OnlineMeetingTranscript.Read.All")
            ? probe("unknown", noMeetingDetail)
            : probe("denied", "Permission is missing from the app token."),
          roles.includes("OnlineMeetingRecording.Read.All")
            ? probe("unknown", noMeetingDetail)
            : probe("denied", "Permission is missing from the app token."),
          hasChat
            ? classifyGraphProbe(
                await graphGet(token, `/users/${encodedUser}/chats?$top=1`),
                "Graph can list this user's Teams chats.",
                false,
              )
            : probe(
                "denied",
                "Chat.Read.All is not on this app's token yet.",
              ),
        ];

    const mail = roles.includes("Mail.Send")
      ? probe(
          "ok",
          "Mail.Send is on the token. The app sends as the target mailbox; it does not need to read Sent Items.",
        )
      : probe("denied", "Mail.Send is missing from the app token.");

    const probes = {
      calendar: classifyGraphProbe(
        calendar,
        "Calendar view is readable.",
        !roles.includes("Calendars.Read"),
      ),
      transcripts,
      recordings,
      chat,
      mail,
      todo: classifyGraphProbe(
        todo,
        "Microsoft To Do lists are readable.",
        !roles.includes("Tasks.ReadWrite.All"),
      ),
    };

    const hints: string[] = [];
    if (missingRoles.includes("Chat.Read.All") || missingRoles.includes("ChatMessage.Read.All")) {
      hints.push(
        `This app's token (Application ID ${clientId}) still has no Chat.Read.All. Open that exact App registration in Entra — the Overview Application (client) ID must match — grant Chat.Read.All and ChatMessage.Read.All as Application permissions, click Grant admin consent, then click Refresh token. Granted on a different app will not appear here.`,
      );
    }
    if (probes.transcripts.status === "denied") {
      hints.push(
        probes.transcripts.detail.includes("EnableGraphTranscriptAccess")
          ? "Turn on Teams Graph transcript access: Set-CsTeamsMeetingConfiguration -EnableGraphTranscriptAccess $true -Identity Global."
          : "Transcript listing was denied for this meeting. Confirm an Application Access Policy covers the organizer.",
      );
    }
    if (probes.chat.status === "denied" && hasChat) {
      hints.push(
        "Chat.Read.All is on the token, but Graph still denied this chat. App-only access cannot read chats with guests or people from other tenants.",
      );
    }
    if (probes.calendar.status === "denied") {
      hints.push(
        "Confirm an Application Access Policy covers the TARGET_USER_ID mailbox.",
      );
    }

    return {
      token: "ok",
      targetUser,
      summaryEmailRecipient,
      clientId,
      roles,
      missingRoles,
      probes,
      hint: hints[0] ?? null,
      hints,
    };
  } catch (error) {
    return {
      token: "failed",
      targetUser,
      summaryEmailRecipient,
      clientId,
      roles: [],
      missingRoles: [...REQUIRED_ROLES],
      probes: {
        calendar: probe("unknown", "Skipped because token acquisition failed."),
        transcripts: probe("unknown", "Skipped because token acquisition failed."),
        recordings: probe("unknown", "Skipped because token acquisition failed."),
        chat: probe("unknown", "Skipped because token acquisition failed."),
        mail: probe("unknown", "Skipped because token acquisition failed."),
        todo: probe("unknown", "Skipped because token acquisition failed."),
      },
      hint:
        error instanceof Error
          ? error.message
          : "Could not acquire a Microsoft Graph token.",
      hints: [
        error instanceof Error
          ? error.message
          : "Could not acquire a Microsoft Graph token.",
      ],
    };
  }
}
