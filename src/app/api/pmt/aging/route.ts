import { NextRequest, NextResponse } from "next/server";
import { getAssigneeAging } from "@/lib/assignee-aging";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Aging tickets are only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const report = await getAssigneeAging(request.nextUrl.searchParams.get("assignee") ?? "");
    return NextResponse.json(report);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
