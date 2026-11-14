import { NextRequest, NextResponse } from "next/server";
import { analyzeCodeQcTickets } from "@/lib/code-qc";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";

export const dynamic = "force-dynamic";
export const maxDuration = 300;

function isAnalyzeRequest(
  value: unknown,
): value is { issueKeys: string[]; boardName?: string } {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const body = value as Record<string, unknown>;
  return (
    Array.isArray(body.issueKeys) &&
    body.issueKeys.every((key) => typeof key === "string") &&
    (body.boardName === undefined || typeof body.boardName === "string")
  );
}

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "CODE QC analysis is only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const body: unknown = await request.json();
    if (!isAnalyzeRequest(body)) {
      return NextResponse.json(
        { error: "Select at least one SonarQube issue to analyze.", code: "INVALID_REQUEST" },
        { status: 400 },
      );
    }
    const result = await analyzeCodeQcTickets(body.issueKeys, body.boardName);
    return NextResponse.json(result);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
