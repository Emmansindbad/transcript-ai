import { NextRequest, NextResponse } from "next/server";
import type { CodeQcSendRequest } from "@/types/code-qc";
import { sendCodeQcTicketsToGroup } from "@/lib/code-qc";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";

export const dynamic = "force-dynamic";
export const maxDuration = 120;

function isSendRequest(value: unknown): value is CodeQcSendRequest {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const body = value as Record<string, unknown>;
  return (
    typeof body.boardName === "string" &&
    typeof body.note === "string" &&
    Array.isArray(body.groupUserIds) &&
    body.groupUserIds.every((id) => typeof id === "string") &&
    Array.isArray(body.tickets) &&
    body.tickets.every((ticket) => {
      if (typeof ticket !== "object" || ticket === null) {
        return false;
      }
      const item = ticket as Record<string, unknown>;
      return (
        typeof item.externalKey === "string" &&
        typeof item.number === "number" &&
        typeof item.url === "string" &&
        (item.title === undefined || typeof item.title === "string")
      );
    })
  );
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Sending CODE QC tickets is only allowed from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const body: unknown = await request.json();
    if (!isSendRequest(body)) {
      return NextResponse.json(
        { error: "Create the PMT tickets first, then choose a group and message.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }
    const teams = await sendCodeQcTicketsToGroup(body);
    return NextResponse.json({
      teams,
      message: teams.sent
        ? `Sent the ticket links to the ${teams.topic} group.`
        : "Could not send the Teams message.",
    });
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
