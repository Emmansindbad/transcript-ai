import { NextRequest, NextResponse } from "next/server";
import { getGraphHealthReport } from "@/lib/health";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export async function GET(request: NextRequest): Promise<NextResponse> {
  try {
    const refresh = request.nextUrl.searchParams.get("refresh") === "1";
    const report = await getGraphHealthReport({ refresh });
    return NextResponse.json(report);
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
