import { NextRequest, NextResponse } from "next/server";
import { isAssigneeAgingReport } from "@/lib/assignee-aging";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";
import { sendAssigneeAgingToPerson } from "@/lib/meeting-chats";

export const dynamic = "force-dynamic";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Sending aging tickets is only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const body: unknown = await request.json();
    if (!isRecord(body) || !isAssigneeAgingReport(body.report)) {
      return NextResponse.json(
        { error: "Choose a person and keep at least one ticket.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }
    const note = typeof body.note === "string" ? body.note : "";
    const recipient = await sendAssigneeAgingToPerson(note, body.report);
    return NextResponse.json({ message: `Aging tickets sent to ${recipient}.` });
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
