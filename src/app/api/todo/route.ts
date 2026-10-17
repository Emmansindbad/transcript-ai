import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem } from "@/types/meeting";
import {
  exportActionItemsToTodo,
  MICROSOFT_TODO_URL,
} from "@/lib/todo";
import { markItemsFiled } from "@/lib/store";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

interface TodoRequestBody {
  actionItems?: MeetingActionItem[];
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  try {
    const body = (await request.json()) as TodoRequestBody;

    if (!body.actionItems?.length) {
      return NextResponse.json(
        { error: "actionItems are required.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }

    const included = body.actionItems.filter((item) => item.include !== false);
    if (included.length === 0) {
      return NextResponse.json(
        { error: "Select at least one action item to send to To Do.", code: "NO_ACTION_ITEMS" },
        { status: 400 },
      );
    }
    const result = await exportActionItemsToTodo(included);
    await markItemsFiled(included.map((item) => item.filingKey).filter(Boolean));

    return NextResponse.json({
      ...result,
      todoUrl: MICROSOFT_TODO_URL,
      message: `Created ${result.created} task(s) in Microsoft To Do list "${result.listName}".`,
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
