import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem } from "@/types/meeting";
import { isLocalRequest } from "@/lib/local-control";
import { meetingPmtViewerEmail } from "@/lib/pmt-analytics";
import { createPmtTickets } from "@/lib/pmt";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export const dynamic = "force-dynamic";

interface PmtRequestBody {
  actionItems?: MeetingActionItem[];
  destinationId?: number;
  destinationName?: string;
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      {
        error: "Creating PMT tickets is only allowed from this PC.",
        code: "FORBIDDEN",
      },
      { status: 403 },
    );
  }

  try {
    const body = (await request.json()) as PmtRequestBody;
    if (!body.actionItems?.length) {
      return NextResponse.json(
        { error: "actionItems are required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const included = body.actionItems.filter((item) => item.include !== false);
    if (included.length === 0) {
      return NextResponse.json(
        {
          error: "Select at least one action item to send.",
          code: "NO_ACTION_ITEMS",
        },
        { status: 400 },
      );
    }

    const destinationId = body.destinationId;
    const destinationName = body.destinationName?.trim() ?? "";
    if (!Number.isInteger(destinationId) || !destinationId || destinationId <= 0 || !destinationName) {
      return NextResponse.json(
        { error: "Choose the board or sheet for these tickets.", code: "PMT_DESTINATION" },
        { status: 400 },
      );
    }

    const viewerEmail = await meetingPmtViewerEmail();
    const result = await createPmtTickets(
      included,
      { id: destinationId, name: destinationName },
      viewerEmail,
    );
    return NextResponse.json(result);
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
