import { NextRequest, NextResponse } from "next/server";
import {
  getAllRecentMeetings,
  getRecentMeetingsWithTranscripts,
} from "@/lib/m365";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export async function GET(request: NextRequest): Promise<NextResponse> {
  try {
    const userId = request.nextUrl.searchParams.get("userId");
    const result =
      !userId || userId === "all"
        ? await getAllRecentMeetings()
        : await getRecentMeetingsWithTranscripts(userId);

    return NextResponse.json({
      meetings: result.meetings,
      transcriptAccessDisabled: result.transcriptAccessDisabled,
      fromCache: result.fromCache ?? false,
      ...(result.transcriptAccessDisabled
        ? {
            error:
              "Microsoft Graph transcript access is still disabled for this tenant. Showing calendar meetings instead. Recordings can still be analyzed.",
            code: "TRANSCRIPTS_DISABLED",
          }
        : {}),
    });
  } catch (error) {
    return NextResponse.json(
      {
        error: getErrorMessage(error),
        code: getErrorCode(error),
      },
      { status: getErrorStatusCode(error) },
    );
  }
}
