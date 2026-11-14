import { NextRequest, NextResponse } from "next/server";
import type { CodeQcCreateRequest } from "@/types/code-qc";
import { createCodeQcTickets, loadCodeQcSnapshot } from "@/lib/code-qc";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";

export const dynamic = "force-dynamic";
export const maxDuration = 300;

function isTicketInput(value: unknown): boolean {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const ticket = value as Record<string, unknown>;
  return (
    typeof ticket.title === "string" &&
    (ticket.type === "Task" || ticket.type === "Bug") &&
    (ticket.priority === "High" || ticket.priority === "Medium" || ticket.priority === "Low") &&
    Array.isArray(ticket.issueKeys) &&
    ticket.issueKeys.every((key) => typeof key === "string")
  );
}

function isCreateRequest(value: unknown): value is CodeQcCreateRequest {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const body = value as Record<string, unknown>;
  if (typeof body.boardName !== "string") {
    return false;
  }
  if (body.assignAllUserId !== undefined && typeof body.assignAllUserId !== "string") {
    return false;
  }
  const hasTickets = Array.isArray(body.tickets) && body.tickets.every(isTicketInput);
  const hasKeys =
    Array.isArray(body.issueKeys) && body.issueKeys.every((key) => typeof key === "string");
  return hasTickets || hasKeys;
}

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "CODE QC is only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const snapshot = await loadCodeQcSnapshot();
    return NextResponse.json(snapshot);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Creating CODE QC tickets is only allowed from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const body: unknown = await request.json();
    if (!isCreateRequest(body)) {
      return NextResponse.json(
        { error: "Analyze the issues first, then create the grouped tickets.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }
    const result = await createCodeQcTickets({
      boardName: body.boardName,
      ...(body.tickets ? { tickets: body.tickets } : {}),
      ...(body.issueKeys ? { issueKeys: body.issueKeys } : {}),
      ...(body.assignAllUserId ? { assignAllUserId: body.assignAllUserId } : {}),
    });
    return NextResponse.json(result);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
