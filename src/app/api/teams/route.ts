import { NextRequest, NextResponse } from "next/server";
import type { MeetingActionItem, TeamsSendResult } from "@/types/meeting";
import { isLocalRequest } from "@/lib/local-control";
import { markItemsFiled } from "@/lib/store";
import { sendActionItemsToAssignees } from "@/lib/teams-message";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export const dynamic = "force-dynamic";

interface TeamsRequestBody {
  actionItems?: MeetingActionItem[];
  includeIris?: boolean;
}

function failureMessage(result: TeamsSendResult): string {
  const parts: string[] = [];
  if (result.unmatched.length > 0) {
    parts.push(`No Teams user matched: ${result.unmatched.join(", ")}.`);
  }
  for (const failure of result.failed) {
    parts.push(`${failure.assignee}: ${failure.error}`);
  }
  return parts.join(" ") || "No Teams messages were sent.";
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      {
        error: "Sending Teams chats is only allowed from this PC.",
        code: "FORBIDDEN",
      },
      { status: 403 },
    );
  }

  try {
    const body = (await request.json()) as TeamsRequestBody;
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

    const result = await sendActionItemsToAssignees(included, {
      includeIris: body.includeIris === true,
    });
    if (result.filedKeys.length > 0) {
      await markItemsFiled(result.filedKeys);
    }

    if (result.sent === 0) {
      return NextResponse.json(
        {
          error: failureMessage(result),
          code: "TEAMS_SEND_FAILED",
          ...result,
        },
        { status: 400 },
      );
    }

    const recipientLabel =
      result.recipients.length === 1
        ? result.recipients[0]
        : result.recipients.join(", ");
    const irisNote = result.includedIris
      ? result.recipients.length === 1
        ? " Iris is included in the chat."
        : " Iris is included in each chat."
      : "";
    return NextResponse.json({
      ...result,
      message: `Messaged ${recipientLabel} as ${result.sender}.${irisNote}`,
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
