import { NextRequest, NextResponse } from "next/server";
import { ksaDayUtcRange } from "@/lib/ksa-time";
import {
  getAllRecordedMeetingsOnDay,
  getRecentMeetingsWithTranscripts,
  listTenantUsers,
} from "@/lib/m365";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export async function GET(request: NextRequest): Promise<NextResponse> {
  try {
    const userId = request.nextUrl.searchParams.get("userId")?.trim() ?? "";
    const date = request.nextUrl.searchParams.get("date")?.trim() ?? "";
    if (date && !ksaDayUtcRange(date)) {
      return NextResponse.json(
        { error: "Choose a valid date.", code: "INVALID_DATE" },
        { status: 400 },
      );
    }

    const everyone = !userId || userId === "all";
    if (everyone && !date) {
      return NextResponse.json(
        {
          error: "Choose a date to load every recorded meeting that day.",
          code: "DATE_REQUIRED",
        },
        { status: 400 },
      );
    }

    let result;
    if (everyone) {
      result = await getAllRecordedMeetingsOnDay(date);
    } else {
      const users = await listTenantUsers();
      const person = users.find(
        (user) => user.userPrincipalName.toLowerCase() === userId.toLowerCase(),
      );
      result = await getRecentMeetingsWithTranscripts(
        userId,
        30,
        person?.displayName ?? userId,
        date || undefined,
      );
    }

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
