import { NextRequest, NextResponse } from "next/server";
import { isLocalRequest } from "@/lib/local-control";
import { getErrorMessage } from "@/lib/errors";
import {
  completeTeamsSignIn,
  safeEqual,
  TEAMS_STATE_COOKIE,
  TEAMS_VERIFIER_COOKIE,
  teamsRedirectUri,
} from "@/lib/teams-auth";

export const dynamic = "force-dynamic";

function redirectHome(request: NextRequest, reason?: string): NextResponse {
  const url = new URL("/", request.url);
  if (reason) {
    url.searchParams.set("teams", "error");
    url.searchParams.set("reason", reason.slice(0, 400));
  } else {
    url.searchParams.set("teams", "connected");
  }
  const response = NextResponse.redirect(url);
  response.cookies.delete(TEAMS_STATE_COOKIE);
  response.cookies.delete(TEAMS_VERIFIER_COOKIE);
  return response;
}

function explainOAuthError(error: string, description: string | null): string {
  const text = `${error} ${description ?? ""}`;
  if (text.includes("AADSTS50011")) {
    return "Microsoft rejected the sign-in redirect. In the Entra app, add a Web redirect URI of http://localhost:3000/api/auth/teams/callback, then try again.";
  }
  if (text.includes("AADSTS65001") || text.includes("AADSTS650053") || error === "consent_required") {
    return "This app still cannot send chats as you. Add delegated Chat.Create, ChatMessage.Send, and User.Read, grant admin consent, then sign in again.";
  }
  if (description?.trim()) {
    return description.trim().slice(0, 400);
  }
  return "Teams sign-in was cancelled or denied.";
}

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

  const oauthError = request.nextUrl.searchParams.get("error");
  if (oauthError) {
    return redirectHome(
      request,
      explainOAuthError(
        oauthError,
        request.nextUrl.searchParams.get("error_description"),
      ),
    );
  }

  const code = request.nextUrl.searchParams.get("code");
  const state = request.nextUrl.searchParams.get("state");
  const expectedState = request.cookies.get(TEAMS_STATE_COOKIE)?.value;
  const codeVerifier = request.cookies.get(TEAMS_VERIFIER_COOKIE)?.value;
  if (!code || !state || !expectedState || !codeVerifier || !safeEqual(state, expectedState)) {
    return redirectHome(
      request,
      "Teams sign-in expired. Click Sign in to Teams and try again.",
    );
  }

  const host = request.headers.get("host") ?? "localhost:3000";
  try {
    await completeTeamsSignIn({
      code,
      redirectUri: teamsRedirectUri(host),
      codeVerifier,
    });
    return redirectHome(request);
  } catch (error) {
    return redirectHome(request, getErrorMessage(error));
  }
}
