import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import { getPmtBoardAnalytics, listPmtBoardsForSignedInUser } from "@/lib/pmt-analytics";
import { getErrorCode, getErrorMessage, getErrorStatusCode } from "@/lib/errors";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      { error: "PMT analytics is only available from this PC.", code: "FORBIDDEN" },
      { status: 403 },
    );
  }

  const boardId = Number(request.nextUrl.searchParams.get("boardId") ?? "");
  const range = request.nextUrl.searchParams.get("range") === "month" ? "month" : "week";
  try {
    if (!request.nextUrl.searchParams.has("boardId")) {
      const boards = await listPmtBoardsForSignedInUser();
      return NextResponse.json({ boards });
    }
    const report = await getPmtBoardAnalytics(boardId, range);
    return NextResponse.json(report);
  } catch (error) {
    return NextResponse.json(
      { error: getErrorMessage(error), code: getErrorCode(error) },
      { status: getErrorStatusCode(error) },
    );
  }
}
