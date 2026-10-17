import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import { sendAnalyticsToMeetingChat } from "@/lib/meeting-chats";
import { isPmtAnalyticsReport } from "@/lib/pmt-analytics";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";

export const dynamic = "force-dynamic";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Sending analytics is only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const body: unknown = await request.json();
    if (!isRecord(body) || typeof body.chatId !== "string" || !isPmtAnalyticsReport(body.report)) {
      return NextResponse.json(
        { error: "Choose a meeting chat and load a preview first.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }
    const note = typeof body.note === "string" ? body.note : "";
    await sendAnalyticsToMeetingChat(body.chatId, note, body.report);
    return NextResponse.json({ message: "Analytics sent to the meeting chat." });
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
