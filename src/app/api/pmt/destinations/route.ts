import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";
import { listMeetingDestinations } from "@/lib/pmt-analytics";

export const dynamic = "force-dynamic";

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      {
        error: "Listing PMT destinations is only allowed from this PC.",
        code: "FORBIDDEN",
      },
      { status: 403 },
    );
  }

  const lane = request.nextUrl.searchParams.get("lane");
  if (lane !== "development" && lane !== "sheets") {
    return NextResponse.json(
      { error: "Choose Meetings or Meeting for Sheets.", code: "INVALID_REQUEST" },
      { status: 400 },
    );
  }

  try {
    const destinations = await listMeetingDestinations(lane);
    return NextResponse.json({ destinations });
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
