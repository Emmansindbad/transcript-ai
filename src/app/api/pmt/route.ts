import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem } from "@/types/meeting";
import { isLocalRequest } from "@/lib/local-control";
import { createPmtTickets } from "@/lib/pmt";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export const dynamic = "force-dynamic";

interface PmtRequestBody {
  actionItems?: MeetingActionItem[];
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

    const result = await createPmtTickets(included);
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
