import { NextRequest, NextResponse } from "next/server";
import {
  fetchLatestParsedTranscript,
  fetchMeetingChat,
  listRecordings,
  parseOrganizerIdFromJoinUrl,
  pickArtifactForOccurrence,
  resolveOnlineMeetingId,
} from "@/lib/m365";
import {
  formatKsaDateTime,
  getKsaDateKey,
} from "@/lib/ksa-time";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

interface TranscriptRequestBody {
  joinUrl?: string;
  organizerId?: string;
  onlineMeetingId?: string;
  latestTranscriptId?: string;
  subject?: string;
  meetingStart?: string;
  meetingEnd?: string;
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  let body: TranscriptRequestBody = {};

  try {
    body = (await request.json()) as TranscriptRequestBody;

    if (!body.joinUrl) {
      return NextResponse.json(
        { error: "joinUrl is required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const organizerId =
      body.organizerId ?? parseOrganizerIdFromJoinUrl(body.joinUrl);

    if (!organizerId) {
      return NextResponse.json(
        {
          error:
            "Could not determine the meeting organizer from the join URL.",
          code: "MISSING_ORGANIZER",
        },
        { status: 400 },
      );
    }

    const meetingId =
      body.onlineMeetingId ??
      (await resolveOnlineMeetingId(organizerId, body.joinUrl));

    let transcript = "";
    let transcriptId: string | undefined;
    let transcriptCreatedAt = "";
    let transcriptError: string | null = null;

    try {
      const parsed = await fetchLatestParsedTranscript(
        body.joinUrl,
        organizerId,
        meetingId,
        body.latestTranscriptId,
        body.meetingStart,
        body.meetingEnd,
      );
      transcript = parsed.transcript;
      transcriptId = parsed.transcriptId;
      transcriptCreatedAt = parsed.createdDateTime;
    } catch (error) {
      transcriptError = getErrorCode(error);
      if (
        transcriptError !== "TRANSCRIPTS_DISABLED" &&
        transcriptError !== "NO_TRANSCRIPT" &&
        transcriptError !== "EMPTY_TRANSCRIPT"
      ) {
        throw error;
      }
    }

    let recordingAvailable = false;
    let recordingCreatedAt = "";
    try {
      const recordings = await listRecordings(organizerId, meetingId);
      const matched = body.meetingStart
        ? pickArtifactForOccurrence(
            recordings,
            body.meetingStart,
            body.meetingEnd,
          )
        : recordings[0];
      recordingAvailable = Boolean(matched);
      recordingCreatedAt = matched?.createdDateTime ?? "";
    } catch {
      recordingAvailable = false;
    }

    const occurrenceDateKsa = body.meetingStart
      ? getKsaDateKey(body.meetingStart)
      : "";
    const artifactCreatedAt =
      transcriptCreatedAt || recordingCreatedAt || body.meetingStart || "";
    const artifactDateKsa =
      occurrenceDateKsa && occurrenceDateKsa !== "unknown"
        ? occurrenceDateKsa
        : artifactCreatedAt
          ? getKsaDateKey(artifactCreatedAt)
          : "";
    const artifactDateLabel = artifactCreatedAt
      ? formatKsaDateTime(artifactCreatedAt)
      : "";
    const artifactSource = transcriptCreatedAt
      ? "transcript"
      : recordingCreatedAt
        ? "recording"
        : "meeting start";

    let chat = "";
    let chatMessageCount = 0;
    let chatStatus: "ok" | "empty" | "forbidden" | "unavailable" = "empty";
    let chatError: string | null = null;
    try {
      const chatResult = await fetchMeetingChat(
        organizerId,
        body.joinUrl,
        meetingId,
        artifactDateKsa && artifactDateKsa !== "unknown"
          ? artifactDateKsa
          : undefined,
      );
      chat = chatResult.chat;
      chatMessageCount = chatResult.messageCount;
      chatStatus = chatResult.status;
      chatError = chatResult.errorMessage ?? null;
    } catch (error) {
      chat = "";
      chatStatus = "unavailable";
      chatError = getErrorMessage(error);
    }

    if (transcript || chat) {
      return NextResponse.json({
        transcript: transcript || null,
        chat: chat || null,
        chatMessageCount,
        chatStatus,
        chatError,
        meetingSubject: body.subject ?? "Meeting",
        transcriptId,
        recordingAvailable,
        transcriptUnavailable: !transcript,
        artifactDateKsa,
        artifactDateLabel,
        artifactSource,
      });
    }

    if (recordingAvailable) {
      return NextResponse.json(
        {
          error:
            transcriptError === "TRANSCRIPTS_DISABLED"
              ? "Microsoft Graph access to meeting transcripts is turned off for this organization, so meetings cannot be loaded."
              : getErrorMessage(
                  new Error("No transcript or meeting chat was available."),
                ),
          code: transcriptError ?? "NO_TRANSCRIPT",
          recordingAvailable: true,
          chat: chat || null,
          chatMessageCount,
          chatStatus,
          chatError,
        },
        { status: 403 },
      );
    }

    return NextResponse.json(
      {
        error:
          "No transcript or meeting chat was available for this meeting.",
        code: transcriptError ?? "NO_TRANSCRIPT",
        recordingAvailable: false,
      },
      { status: 404 },
    );
  } catch (error) {
    const code = getErrorCode(error);
    let recordingAvailable = false;

    if (code === "TRANSCRIPTS_DISABLED") {
      try {
        const joinUrl = body.joinUrl;
        const organizerId =
          body.organizerId ??
          (joinUrl ? parseOrganizerIdFromJoinUrl(joinUrl) : null);
        if (joinUrl && organizerId) {
          const meetingId =
            body.onlineMeetingId ??
            (await resolveOnlineMeetingId(organizerId, joinUrl));
          recordingAvailable =
            (await listRecordings(organizerId, meetingId)).length > 0;
        }
      } catch {
        recordingAvailable = false;
      }
    }

    return NextResponse.json(
      {
        error: getErrorMessage(error),
        code,
        recordingAvailable,
      },
      { status: getErrorStatusCode(error) },
    );
  }
}
