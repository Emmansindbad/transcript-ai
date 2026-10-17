import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import {
  createTeamsSignInRequest,
  TEAMS_STATE_COOKIE,
  TEAMS_VERIFIER_COOKIE,
  teamsRedirectUri,
} from "@/lib/teams-auth";
import { getErrorMessage } from "@/lib/errors";

export const dynamic = "force-dynamic";

const COOKIE_OPTIONS = {
  httpOnly: true,
  sameSite: "lax" as const,
  path: "/",
  maxAge: 10 * 60,
  secure: false,
};

export async function GET(request: NextRequest): Promise<NextResponse> {
  if (!isLocalRequest(request)) {
    return NextResponse.json(
      {
        error: "Teams sign-in is only allowed from this PC.",
        code: "FORBIDDEN",
      },
      { status: 403 },
    );
  }

  const host = request.headers.get("host") ?? "localhost:3000";
  try {
    const signIn = await createTeamsSignInRequest(teamsRedirectUri(host));
    const response = NextResponse.redirect(signIn.url);
    response.cookies.set(TEAMS_STATE_COOKIE, signIn.state, COOKIE_OPTIONS);
    response.cookies.set(TEAMS_VERIFIER_COOKIE, signIn.codeVerifier, COOKIE_OPTIONS);
    return response;
  } catch (error) {
    const url = new URL("/", request.url);
    url.searchParams.set("teams", "error");
    url.searchParams.set("reason", getErrorMessage(error).slice(0, 400));
    return NextResponse.redirect(url);
  }
}
