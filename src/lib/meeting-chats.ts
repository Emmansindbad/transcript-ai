import "server-only";

import type {
  MeetingChatOption,
  PmtAnalyticsReport,
  PmtAssigneeAgingReport,
  TenantUser,
} from "@/types/meeting";
import { agingDayLabel, analyticsAdvice, VISIBLE_AGING_TICKETS } from "@/lib/analytics-brief";
import { matchAssignee } from "@/lib/assignees";
import { getPmtBaseUrl } from "@/lib/env";
import { AppError, GraphError } from "@/lib/errors";
import { getGraphClient, listTenantUsers } from "@/lib/m365";
import { openDirectChatWithEmail } from "@/lib/teams-message";
import { getDelegatedGraphClient, getTeamsAccount } from "@/lib/teams-auth";

interface GraphChat {
  id?: string;
  topic?: string;
  chatType?: string;
  lastUpdatedDateTime?: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function chatTopic(chat: GraphChat): string {
  const topic = chat.topic?.trim() ?? "";
  if (topic) {
    return topic;
  }
  const when = chat.lastUpdatedDateTime?.slice(0, 10) ?? "";
  return when ? `Meeting chat · ${when}` : "Meeting chat";
}

async function signedInEmail(): Promise<string> {
  const account = await getTeamsAccount();
  const email = account.username?.trim() ?? "";
  if (!account.signedIn || !email.includes("@")) {
    throw new AppError(
      "Sign in to Teams to load meeting chats.",
      401,
      "TEAMS_SIGN_IN",
    );
  }
  return email;
}

async function readChats(userId: string): Promise<GraphChat[]> {
  const client = getGraphClient();
  try {
    const filtered = (await client
      .api(`/users/${userId}/chats`)
      .filter("chatType eq 'meeting'")
      .select("id,topic,chatType,lastUpdatedDateTime")
      .top(50)
      .get()) as { value?: GraphChat[] };
    return filtered.value ?? [];
  } catch {
    const listed = (await client
      .api(`/users/${userId}/chats`)
      .select("id,topic,chatType,lastUpdatedDateTime")
      .top(50)
      .get()) as { value?: GraphChat[] };
    return (listed.value ?? []).filter((chat) => chat.chatType === "meeting");
  }
}

export async function listMeetingChats(): Promise<MeetingChatOption[]> {
  const email = await signedInEmail();
  const users = await listTenantUsers();
  const user = users.find(
    (candidate) =>
      candidate.userPrincipalName.toLowerCase() === email.toLowerCase() ||
      candidate.mail?.toLowerCase() === email.toLowerCase(),
  );
  if (!user) {
    throw new AppError(
      "The signed-in Microsoft 365 account is not in the company directory.",
      404,
      "TEAMS_USER",
    );
  }

  const chats = await readChats(user.id);
  return chats
    .filter((chat): chat is GraphChat & { id: string } => Boolean(chat.id))
    .sort((left, right) => (right.lastUpdatedDateTime ?? "").localeCompare(left.lastUpdatedDateTime ?? ""))
    .slice(0, 40)
    .map((chat) => ({ id: chat.id, topic: chatTopic(chat) }));
}

function ticketLine(ticket: PmtAnalyticsReport["hiddenAging"][number]): string {
  const due = ticket.due ? `Due ${ticket.due}` : "No due date";
  const label = `#${ticket.number} ${ticket.name}`;
  return `<li><a href="${escapeHtml(ticket.url)}"><b>${escapeHtml(label)}</b></a><br>${escapeHtml(ticket.column)} · ${ticket.days} days · ${escapeHtml(ticket.owner)} · ${escapeHtml(due)}</li>`;
}

interface TeamsMention {
  id: number;
  mentionText: string;
  mentioned: Record<string, unknown>;
}

function findDirectoryUser(name: string, users: TenantUser[]): TenantUser | null {
  const exact = users.find((user) => user.displayName.toLowerCase() === name.toLowerCase());
  if (exact) {
    return exact;
  }
  const matched = matchAssignee(name, users);
  if (!matched) {
    return null;
  }
  return (
    users.find(
      (user) => user.userPrincipalName.toLowerCase() === matched.userPrincipalName.toLowerCase(),
    ) ?? null
  );
}

function noteWithMentions(
  note: string,
  chatId: string,
  users: TenantUser[],
): { html: string; mentions: TeamsMention[] } {
  const mentions: TeamsMention[] = [];
  const pattern = /@Everyone|@([^@\n]+?)(?=\s@|\s—|\n|$)/g;
  let html = "";
  let cursor = 0;
  for (const match of note.matchAll(pattern)) {
    const index = match.index ?? 0;
    html += escapeHtml(note.slice(cursor, index));
    if (match[0] === "@Everyone") {
      const id = mentions.length;
      mentions.push({
        id,
        mentionText: "Everyone",
        mentioned: {
          conversation: {
            id: chatId,
            displayName: "Everyone",
            conversationIdentityType: "chat",
          },
        },
      });
      html += `<at id="${id}">Everyone</at>`;
    } else {
      const name = match[1]?.trim() ?? "";
      const person = name ? findDirectoryUser(name, users) : null;
      if (!person) {
        html += escapeHtml(match[0]);
      } else {
        const id = mentions.length;
        mentions.push({
          id,
          mentionText: person.displayName,
          mentioned: {
            user: {
              id: person.id,
              displayName: person.displayName,
              userIdentityType: "aadUser",
            },
          },
        });
        html += `<at id="${id}">${escapeHtml(person.displayName)}</at> `;
      }
    }
    cursor = index + match[0].length;
  }
  html += escapeHtml(note.slice(cursor));
  return { html: html.split("\n").join("<br>"), mentions };
}

function previewHtml(report: PmtAnalyticsReport, noteHtml: string): string {
  const shown = report.hiddenAging.slice(0, VISIBLE_AGING_TICKETS);
  const remaining = report.hiddenAging.length - shown.length;
  const tickets =
    shown.length === 0
      ? "<p>Every aging ticket is already named on the PMT page.</p>"
      : `<p><b>Aging tickets to act on</b></p><ul>${shown.map(ticketLine).join("")}</ul>${
          remaining > 0 ? `<p>And ${remaining} more on the board.</p>` : ""
        }`;
  return [
    noteHtml ? `<p>${noteHtml}</p>` : "",
    `<p><b>${escapeHtml(report.board.name)}</b><br>${escapeHtml(report.fromLabel)} – ${escapeHtml(report.toLabel)}</p>`,
    `<p><b>Open</b> ${report.open}&nbsp;&nbsp;<b>Closed</b> ${report.closed}&nbsp;&nbsp;<b>Overdue</b> ${report.overdue}&nbsp;&nbsp;<b>Unassigned</b> ${report.unassigned}</p>`,
    `<p>${escapeHtml(analyticsAdvice(report))}</p>`,
    `<p><b>${escapeHtml(report.chartTitle)}</b></p>`,
    `<p><img alt="${escapeHtml(report.chartTitle)}" src="../hostedContents/1/$value" style="width:100%;max-width:960px;"></p>`,
    tickets,
    `<p><a href="${escapeHtml(report.board.url)}">Open analytics</a></p>`,
  ].join("");
}

function assertPmtLink(value: string): void {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new AppError("The preview has a link PMT did not provide.", 400, "PMT_LINK");
  }
  if (url.host !== new URL(getPmtBaseUrl()).host) {
    throw new AppError("The preview has a link PMT did not provide.", 400, "PMT_LINK");
  }
}

function agingPreviewHtml(report: PmtAssigneeAgingReport, noteHtml: string): string {
  const blocks = report.projects
    .filter((project) => project.tickets.length > 0)
    .map((project) => {
      const items = project.tickets
        .map((ticket) => {
          const days =
            ticket.column && ticket.column !== "No status"
              ? `${agingDayLabel(ticket.days)} in ${ticket.column}`
              : agingDayLabel(ticket.days);
          const label = `#${ticket.number} ${ticket.name}`;
          return `<li><a href="${escapeHtml(ticket.url)}">${escapeHtml(label)}</a> (${escapeHtml(days)})</li>`;
        })
        .join("");
      const more =
        project.more > 0
          ? `<p>And ${project.more} more open ${project.more === 1 ? "ticket" : "tickets"} on this project.</p>`
          : "";
      return `<p><b>${escapeHtml(project.name)} - Top 5 aging tickets</b></p><ol>${items}</ol>${more}`;
    });
  const omitted =
    report.omittedProjects > 0
      ? `<p>And ${report.omittedProjects} other ${report.omittedProjects === 1 ? "project" : "projects"}.</p>`
      : "";
  return [`<p>${noteHtml}</p>`, ...blocks, omitted].filter(Boolean).join("");
}

export async function sendAssigneeAgingToPerson(
  note: string,
  report: PmtAssigneeAgingReport,
): Promise<string> {
  const projects = report.projects
    .map((project) => ({
      ...project,
      tickets: project.tickets.filter((ticket) => ticket.days >= 7).slice(0, 5),
    }))
    .filter((project) => project.tickets.length > 0);
  if (!report.person || projects.length === 0) {
    throw new AppError("Choose at least one aging ticket to send.", 400, "PMT_AGING");
  }
  const trimmedNote = note.trim().slice(0, 4000);
  for (const project of projects) {
    for (const ticket of project.tickets) {
      assertPmtLink(ticket.url);
    }
  }

  const { chatId, displayName } = await openDirectChatWithEmail(report.person.email);
  const users = await listTenantUsers();
  const noted = noteWithMentions(trimmedNote, chatId, users);
  const { client } = await getDelegatedGraphClient();
  try {
    await client.api(`/chats/${chatId}/messages`).post({
      body: {
        contentType: "html",
        content: agingPreviewHtml({ ...report, projects }, noted.html),
      },
      ...(noted.mentions.length > 0 ? { mentions: noted.mentions } : {}),
    });
  } catch (error) {
    const statusCode = isRecord(error) && typeof error.statusCode === "number" ? error.statusCode : 502;
    const message =
      isRecord(error) && typeof error.message === "string"
        ? error.message
        : "Could not send the aging tickets to that person.";
    throw new GraphError(message, statusCode, "TEAMS_SEND_FAILED");
  }
  return displayName;
}

export async function sendAnalyticsToMeetingChat(
  chatId: string,
  note: string,
  report: PmtAnalyticsReport,
): Promise<void> {
  const trimmedNote = note.trim().slice(0, 4000);
  assertPmtLink(report.board.url);
  for (const ticket of report.hiddenAging) {
    assertPmtLink(ticket.url);
  }
  const chats = await listMeetingChats();
  if (!chats.some((chat) => chat.id === chatId)) {
    throw new AppError("Choose a meeting chat from the list.", 400, "TEAMS_CHAT");
  }

  const png = Buffer.from(report.chartPng, "base64");
  const users = await listTenantUsers();
  const noted = noteWithMentions(trimmedNote, chatId, users);
  const { client } = await getDelegatedGraphClient();
  try {
    await client.api(`/chats/${chatId}/messages`).post({
      body: {
        contentType: "html",
        content: previewHtml(report, noted.html),
      },
      ...(noted.mentions.length > 0 ? { mentions: noted.mentions } : {}),
      hostedContents: [
        {
          "@microsoft.graph.temporaryId": "1",
          contentBytes: png.toString("base64"),
          contentType: "image/png",
        },
      ],
    });
  } catch (error) {
    const statusCode = isRecord(error) && typeof error.statusCode === "number" ? error.statusCode : 502;
    const message =
      isRecord(error) && typeof error.message === "string"
        ? error.message
        : "Could not post the analytics to that meeting chat.";
    throw new GraphError(message, statusCode, "TEAMS_SEND_FAILED");
  }
}
