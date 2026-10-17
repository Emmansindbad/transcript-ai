import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import { getTeamsAccount, signOutTeams } from "@/lib/teams-auth";
import {
  getErrorCode,
  getErrorMessage,
  getErrorStatusCode,
} from "@/lib/errors";

export const dynamic = "force-dynamic";

function forbidden(): NextResponse {
  return NextResponse.json(
    {
      error: "Teams sign-in is only allowed from this PC.",
      code: "FORBIDDEN",
    },
    { status: 403 },
  );
}

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return forbidden();
  }

  try {
    const account = await getTeamsAccount();
    return NextResponse.json(account);
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

export async function DELETE(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return forbidden();
  }

  try {
    await signOutTeams();
    return NextResponse.json({ signedIn: false });
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
