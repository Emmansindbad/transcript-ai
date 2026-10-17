import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem } from "@/types/meeting";
import { sendActionItemsEmail } from "@/lib/email";
import { markItemsFiled } from "@/lib/store";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

interface EmailRequestBody {
  actionItems?: MeetingActionItem[];
  analysis?: string;
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  try {
    const body = (await request.json()) as EmailRequestBody;

    if (!body.actionItems?.length) {
      return NextResponse.json(
        { error: "actionItems are required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const analysis = body.analysis?.trim() ?? "";
    const included = body.actionItems.filter((item) => item.include !== false);
    if (included.length === 0) {
      return NextResponse.json(
        { error: "Select at least one action item to send.", code: "NO_ACTION_ITEMS" },
        { status: 400 },
      );
    }
    const result = await sendActionItemsEmail(included, analysis);
    await markItemsFiled(included.map((item) => item.filingKey).filter(Boolean));

    return NextResponse.json({
      ...result,
      message: `Sent to ${result.recipients[0]}.`,
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
