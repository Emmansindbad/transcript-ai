import { NextRequest } from "next/server";
import {
  analyzeTranscript,
  analyzeTranscriptWithVideo,
  type MeetingAnalysisSource,
} from "@/lib/gemini";
import {
  fetchLatestRecording,
  listTenantUsers,
  parseOrganizerIdFromJoinUrl,
} from "@/lib/m365";
import { applyAssigneeMatches } from "@/lib/assignees";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";
import { listFiledKeys, saveAnalysis } from "@/lib/store";
import { toMeetingActionItems } from "@/lib/filing";
import { buildMeetingKey } from "@/lib/meeting-key";
import type { MeetingSummary } from "@/types/meeting";

export const maxDuration = 300;

interface AnalyzeRequestBody {
  transcript?: string;
  chat?: string;
  meetingDate?: string;
  artifactDateKsa?: string;
  artifactDateLabel?: string;
  artifactSource?: string;
  includeVideo?: boolean;
  joinUrl?: string;
  organizerId?: string;
  onlineMeetingId?: string;
  meeting?: MeetingSummary;
}

function encodeSse(event: Record<string, unknown>): Uint8Array {
  return new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`);
}

async function resolveMeetingSource(
  body: AnalyzeRequestBody,
): Promise<MeetingAnalysisSource> {
  return {
    transcript: body.transcript?.trim() ?? "",
    chat: body.chat?.trim() ?? "",
    meetingDate: body.meetingDate?.trim() || undefined,
    artifactDateLabel: body.artifactDateLabel?.trim() || undefined,
    artifactSource: body.artifactSource?.trim() || undefined,
  };
}

export async function POST(request: NextRequest): Promise<Response> {
  const stream = new ReadableStream({
    async start(controller) {
      const send = (event: Record<string, unknown>): void => {
        controller.enqueue(encodeSse(event));
      };

      try {
        const body = (await request.json()) as AnalyzeRequestBody;
        const meeting = body.meeting;
        const source = await resolveMeetingSource(body);

        if (!source.transcript && !source.chat && !body.includeVideo) {
          send({
            type: "error",
            error: "A transcript, meeting chat, or recording is required.",
            code: "INVALID_REQUEST",
            status: 400,
          });
          controller.close();
          return;
        }

        send({ type: "status", message: "Starting Gemini analysis…" });

        let videoUsed = false;
        let videoNotice: string | undefined;
        const onChunk = (text: string): void => {
          send({ type: "chunk", text });
        };

        const organizerId =
          body.organizerId ??
          meeting?.organizerId ??
          (body.joinUrl ? parseOrganizerIdFromJoinUrl(body.joinUrl) : null);
        const joinUrl = body.joinUrl ?? meeting?.joinUrl;

        let result;
        if (body.includeVideo && joinUrl && organizerId) {
          send({ type: "status", message: "Downloading meeting recording…" });
          try {
            const recording = await fetchLatestRecording(
              joinUrl,
              organizerId,
              body.onlineMeetingId ?? meeting?.onlineMeetingId,
              meeting?.start ?? body.meetingDate,
              meeting?.end,
            );
            send({
              type: "status",
              message: "Uploading recording to Gemini…",
            });
            result = await analyzeTranscriptWithVideo(source, {
              data: recording.data,
              mimeType: recording.mimeType,
            }, { onChunk });
            videoUsed = true;
          } catch (error) {
            const code = getErrorCode(error);
            if (
              code === "NO_RECORDING" ||
              code === "VIDEO_TOO_LARGE" ||
              code === "VIDEO_UPLOAD_FAILED" ||
              code === "VIDEO_PROCESSING_TIMEOUT" ||
              code === "VIDEO_PROCESSING_FAILED"
            ) {
              videoNotice = getErrorMessage(error);
              if (!source.transcript && !source.chat) {
                throw error;
              }
              send({
                type: "status",
                message: `${videoNotice} Analyzing transcript and chat instead.`,
              });
              result = await analyzeTranscript(source, { onChunk });
            } else {
              throw error;
            }
          }
        } else {
          if (body.includeVideo) {
            videoNotice =
              "Could not identify the meeting to fetch its recording; analyzed the transcript and chat only.";
          }
          result = await analyzeTranscript(source, { onChunk });
        }

        send({ type: "status", message: "Matching assignees and saving…" });

        const users = await listTenantUsers();
        const parsedItems = applyAssigneeMatches(result.actionItems, users);
        const filedKeys = meeting ? await listFiledKeys() : new Set<string>();
        const actionItems = meeting
          ? toMeetingActionItems(meeting, parsedItems, filedKeys)
          : [];

        if (meeting) {
          await saveAnalysis({
            meetingKey: buildMeetingKey(meeting),
            analysis: result.analysis,
            actionItems,
            modelUsed: result.modelUsed,
            videoUsed,
            videoNotice,
            savedAt: new Date().toISOString(),
          });
        }

        send({
          type: "done",
          analysis: result.analysis,
          actionItems,
          parsedActionItems: parsedItems,
          modelUsed: result.modelUsed,
          videoUsed,
          videoNotice,
        });
      } catch (error) {
        send({
          type: "error",
          error: getErrorMessage(error),
          code: getErrorCode(error),
          status: getErrorStatusCode(error),
        });
      } finally {
        controller.close();
      }
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
    },
  });
}
