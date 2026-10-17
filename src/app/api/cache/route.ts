import { NextRequest, NextResponse } from "next/server";
import { clearAnalyzerCaches } from "@/lib/cache-control";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";
import { isLocalRequest } from "@/lib/local-control";

export const dynamic = "force-dynamic";

export async function POST(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "Clearing the cache is only allowed from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  try {
    const result = await clearAnalyzerCaches();
    return NextResponse.json(result);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
