import "server-only";

import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  ConfidentialClientApplication,
  InteractionRequiredAuthError,
  type AccountInfo,
  type ICachePlugin,
} from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import type { TeamsAccountStatus } from "@/types/meeting";
import { getEnvVar } from "@/lib/env";
import { AppError } from "@/lib/errors";

const TEAMS_SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Chat.Create",
  "https://graph.microsoft.com/ChatMessage.Send",
];

const STORE_DIR = path.join(process.cwd(), "data");
const CACHE_PATH = path.join(STORE_DIR, "teams-user-cache.json");
const PROFILE_PATH = path.join(STORE_DIR, "teams-user.json");

export const TEAMS_REDIRECT_PATH = "/api/auth/teams/callback";
export const TEAMS_STATE_COOKIE = "teams_oauth_state";
export const TEAMS_VERIFIER_COOKIE = "teams_oauth_verifier";

interface TeamsProfile {
  homeAccountId: string;
  username: string;
  displayName: string;
}

interface DelegatedSession {
  accessToken: string;
  displayName: string;
  username: string;
}

export interface TeamsSignInRequest {
  url: string;
  state: string;
  codeVerifier: string;
}

let delegatedClient: ConfidentialClientApplication | null = null;
let operationChain: Promise<void> = Promise.resolve();

function readableTokenCache(raw: string): string | null {
  try {
    JSON.parse(raw);
    return raw;
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    const match = /position (\d+)/.exec(message);
    if (!match) {
      return null;
    }
    const sliced = raw.slice(0, Number(match[1]));
    try {
      JSON.parse(sliced);
      return sliced;
    } catch {
      return null;
    }
  }
}

const cachePlugin: ICachePlugin = {
  beforeCacheAccess: async (context): Promise<void> => {
    try {
      const raw = await readFile(CACHE_PATH, "utf8");
      const cache = readableTokenCache(raw);
      if (!cache) {
        return;
      }
      context.tokenCache.deserialize(cache);
      if (cache !== raw) {
        await writeFile(CACHE_PATH, cache, "utf8");
      }
    } catch {
      // No cache until the first successful sign-in.
    }
  },
  afterCacheAccess: async (context): Promise<void> => {
    if (!context.cacheHasChanged) {
      return;
    }
    await mkdir(STORE_DIR, { recursive: true });
    await writeFile(CACHE_PATH, context.tokenCache.serialize(), "utf8");
  },
};

function withTeamsLock<T>(fn: () => Promise<T>): Promise<T> {
  const run = operationChain.then(fn, fn);
  operationChain = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

function getDelegatedClient(): ConfidentialClientApplication {
  if (!delegatedClient) {
    delegatedClient = new ConfidentialClientApplication({
      auth: {
        clientId: getEnvVar("AZURE_CLIENT_ID"),
        authority: `https://login.microsoftonline.com/${getEnvVar("AZURE_TENANT_ID")}`,
        clientSecret: getEnvVar("AZURE_CLIENT_SECRET"),
      },
      cache: { cachePlugin },
    });
  }
  return delegatedClient;
}

function createPkce(): { verifier: string; challenge: string } {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

export function teamsRedirectUri(host: string): string {
  return `http://${host}${TEAMS_REDIRECT_PATH}`;
}

export function safeEqual(left: string, right: string): boolean {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  if (leftBuffer.length !== rightBuffer.length) {
    return false;
  }
  return timingSafeEqual(leftBuffer, rightBuffer);
}

function accountProfile(account: AccountInfo): TeamsProfile {
  return {
    homeAccountId: account.homeAccountId,
    username: account.username,
    displayName: account.name?.trim() || account.username,
  };
}

async function readProfile(): Promise<TeamsProfile | null> {
  try {
    const raw = await readFile(PROFILE_PATH, "utf8");
    const parsed = JSON.parse(raw) as Partial<TeamsProfile>;
    if (
      typeof parsed.homeAccountId !== "string" ||
      typeof parsed.username !== "string" ||
      typeof parsed.displayName !== "string"
    ) {
      return null;
    }
    return {
      homeAccountId: parsed.homeAccountId,
      username: parsed.username,
      displayName: parsed.displayName,
    };
  } catch {
    return null;
  }
}

async function writeProfile(profile: TeamsProfile): Promise<void> {
  await mkdir(STORE_DIR, { recursive: true });
  await writeFile(PROFILE_PATH, JSON.stringify(profile, null, 2), "utf8");
}

function explainAuthFailure(error: unknown): AppError {
  const message = error instanceof Error ? error.message : "Teams sign-in failed.";
  if (message.includes("AADSTS50011")) {
    return new AppError(
      "Microsoft rejected the sign-in redirect. In the Entra app, add a Web redirect URI of http://localhost:3000/api/auth/teams/callback, then try again.",
      400,
      "TEAMS_REDIRECT_URI",
    );
  }
  if (message.includes("AADSTS65001") || message.includes("AADSTS650053")) {
    return new AppError(
      "This app still cannot send chats as you. Add delegated Chat.Create, ChatMessage.Send, and User.Read, grant admin consent, then sign in again.",
      403,
      "TEAMS_CONSENT_REQUIRED",
    );
  }
  if (error instanceof InteractionRequiredAuthError) {
    return new AppError(
      "Sign in to Teams again. The saved session can no longer send messages.",
      401,
      "TEAMS_SIGNIN_REQUIRED",
    );
  }
  if (message.includes("Unexpected non-whitespace character after JSON")) {
    return new AppError(
      "The saved Teams sign-in was damaged. Sign in to Teams again.",
      401,
      "TEAMS_SIGNIN_REQUIRED",
    );
  }
  return new AppError(message, 401, "TEAMS_SIGNIN_REQUIRED");
}

async function acquireDelegatedSession(): Promise<DelegatedSession> {
  const client = getDelegatedClient();
  const profile = await readProfile();
  const account = profile
    ? await client.getTokenCache().getAccountByHomeId(profile.homeAccountId)
    : (await client.getTokenCache().getAllAccounts())[0] ?? null;

  if (!account) {
    throw new AppError(
      "Sign in to Teams before sending. Messages are posted from your account.",
      401,
      "TEAMS_SIGNIN_REQUIRED",
    );
  }

  try {
    const result = await client.acquireTokenSilent({
      account,
      scopes: TEAMS_SCOPES,
    });
    if (!result?.accessToken) {
      throw new AppError(
        "Sign in to Teams before sending. Messages are posted from your account.",
        401,
        "TEAMS_SIGNIN_REQUIRED",
      );
    }
    const nextProfile = accountProfile(account);
    await writeProfile(nextProfile);
    return {
      accessToken: result.accessToken,
      displayName: nextProfile.displayName,
      username: nextProfile.username,
    };
  } catch (error) {
    if (error instanceof AppError) {
      throw error;
    }
    throw explainAuthFailure(error);
  }
}

export async function createTeamsSignInRequest(
  redirectUri: string,
): Promise<TeamsSignInRequest> {
  return withTeamsLock(async () => {
    const { verifier, challenge } = createPkce();
    const state = randomBytes(16).toString("base64url");
    const url = await getDelegatedClient().getAuthCodeUrl({
      scopes: TEAMS_SCOPES,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
      state,
      prompt: "select_account",
    });
    return { url, state, codeVerifier: verifier };
  });
}

export async function completeTeamsSignIn(input: {
  code: string;
  redirectUri: string;
  codeVerifier: string;
}): Promise<TeamsAccountStatus> {
  return withTeamsLock(async () => {
    try {
      const result = await getDelegatedClient().acquireTokenByCode({
        code: input.code,
        redirectUri: input.redirectUri,
        scopes: TEAMS_SCOPES,
        codeVerifier: input.codeVerifier,
      });
      if (!result.accessToken || !result.account) {
        throw new AppError(
          "Microsoft did not return a user session. Sign in again.",
          401,
          "TEAMS_SIGNIN_REQUIRED",
        );
      }
      const profile = accountProfile(result.account);
      await writeProfile(profile);
      return {
        signedIn: true,
        displayName: profile.displayName,
        username: profile.username,
      };
    } catch (error) {
      if (error instanceof AppError) {
        throw error;
      }
      throw explainAuthFailure(error);
    }
  });
}

export async function getTeamsAccount(): Promise<TeamsAccountStatus> {
  return withTeamsLock(async () => {
    try {
      const session = await acquireDelegatedSession();
      return {
        signedIn: true,
        displayName: session.displayName,
        username: session.username,
      };
    } catch {
      return { signedIn: false };
    }
  });
}

/** Drops the in-memory Teams client. The saved sign-in file stays. */
export function releaseDelegatedClient(): void {
  delegatedClient = null;
}

export async function signOutTeams(): Promise<void> {
  return withTeamsLock(async () => {
    const client = getDelegatedClient();
    const profile = await readProfile();
    if (profile) {
      const account = await client
        .getTokenCache()
        .getAccountByHomeId(profile.homeAccountId);
      if (account) {
        await client.getTokenCache().removeAccount(account);
      }
    }
    await rm(CACHE_PATH, { force: true });
    await rm(PROFILE_PATH, { force: true });
    delegatedClient = null;
  });
}

export async function getDelegatedGraphClient(): Promise<{
  client: Client;
  displayName: string;
  username: string;
}> {
  return withTeamsLock(async () => {
    const session = await acquireDelegatedSession();
    const client = Client.init({
      authProvider: (done) => {
        done(null, session.accessToken);
      },
    });
    return {
      client,
      displayName: session.displayName,
      username: session.username,
    };
  });
}
