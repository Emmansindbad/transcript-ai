import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import { listMeetingChats } from "@/lib/meeting-chats";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Meeting chats are only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const chats = await listMeetingChats();
    return NextResponse.json({ chats });
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
