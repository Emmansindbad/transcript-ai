import "server-only";

import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import path from "node:path";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import type {
  ChatFetchStatus,
  MeetingListResult,
  MeetingSummary,
  TenantUser,
  TranscriptInfo,
} from "@/types/meeting";
import {
  formatChatSourceLinks,
  isImageAttachmentName,
  isSafeHttpUrl,
  toTeamsMessagePermalink,
} from "@/lib/chat-links";
import { getEnvVar } from "@/lib/env";
import { GraphError } from "@/lib/errors";
import {
  formatKsaDateTime,
  getKsaDateKey,
  parseGraphDate,
} from "@/lib/ksa-time";
import { meetingListCacheKey } from "@/lib/meeting-key";

export const MAX_RECORDING_BYTES = 150 * 1024 * 1024;
const MEETING_CACHE_TTL_MS = 12 * 60 * 60 * 1000;
const USER_CACHE_TTL_MS = 10 * 60 * 1000;
const USER_DISK_CACHE_TTL_MS = 12 * 60 * 60 * 1000;
const MEETING_CACHE_PATH = path.join(process.cwd(), "data", "meetings-cache.json");
const USER_CACHE_PATH = path.join(process.cwd(), "data", "users-cache.json");
const USER_CACHE_VERSION = 2;

interface MemoryCacheEntry<T> {
  expiresAt: number;
  value: T;
}

const meetingListCache = new Map<string, MemoryCacheEntry<MeetingListResult>>();
const meetingListInflight = new Map<string, Promise<MeetingListResult>>();

interface StoredMeetingList {
  savedAt: number;
  meetings: MeetingSummary[];
  transcriptAccessDisabled: boolean;
}

interface MeetingCacheFile {
  entries: Record<string, StoredMeetingList>;
}
let tenantUsersCache: MemoryCacheEntry<TenantUser[]> | null = null;
let transcriptsDisabledKnown = false;

interface GraphUser {
  id?: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  accountEnabled?: boolean;
  userType?: string;
  assignedLicenses?: Array<{ skuId?: string }>;
}

interface GraphUsersResponse {
  value?: GraphUser[];
  "@odata.nextLink"?: string;
}

interface GraphCalendarEvent {
  id?: string;
  subject?: string;
  start?: { dateTime?: string };
  end?: { dateTime?: string };
  isOnlineMeeting?: boolean;
  seriesMasterId?: string;
  onlineMeeting?: { joinUrl?: string };
}

interface GraphCalendarViewResponse {
  value?: GraphCalendarEvent[];
}

interface GraphOnlineMeeting {
  id?: string;
  subject?: string;
  joinWebUrl?: string;
  chatInfo?: {
    threadId?: string;
  };
}

interface GraphChatAttachment {
  id?: string;
  contentType?: string;
  contentUrl?: string;
  name?: string;
}

interface GraphChatMessage {
  id?: string;
  createdDateTime?: string;
  messageType?: string;
  deletedDateTime?: string;
  webUrl?: string;
  body?: {
    content?: string;
    contentType?: string;
  };
  from?: {
    user?: {
      displayName?: string;
    };
  };
  attachments?: GraphChatAttachment[];
}

interface GraphChatMessagesResponse {
  value?: GraphChatMessage[];
  "@odata.nextLink"?: string;
}

export interface MeetingChatResult {
  chat: string;
  messageCount: number;
  status: ChatFetchStatus;
  errorCode?: string;
  errorMessage?: string;
}

interface GraphOnlineMeetingsResponse {
  value?: GraphOnlineMeeting[];
}

interface GraphTranscriptsResponse {
  value?: Array<{
    id?: string;
    createdDateTime?: string;
  }>;
  "@odata.nextLink"?: string;
}

interface GraphCallRecording {
  id?: string;
  meetingId?: string;
  callId?: string;
  createdDateTime?: string;
  endDateTime?: string;
  recordingContentUrl?: string;
}

interface GraphRecordingsResponse {
  value?: GraphCallRecording[];
  "@odata.nextLink"?: string;
}

export interface RecordingInfo {
  id: string;
  createdDateTime: string;
  endDateTime?: string;
  meetingId?: string;
  recordingContentUrl?: string;
}

export interface RecordingContent {
  data: Buffer;
  mimeType: string;
}

interface GraphCallTranscript {
  id?: string;
  meetingId?: string;
  createdDateTime?: string;
}

interface GraphGetAllTranscriptsResponse {
  value?: GraphCallTranscript[];
  "@odata.nextLink"?: string;
}

interface GraphGetAllRecordingsResponse {
  value?: GraphCallRecording[];
  "@odata.nextLink"?: string;
}

const OCCURRENCE_BEFORE_MS = 30 * 60 * 1000;
const OCCURRENCE_AFTER_MS = 12 * 60 * 60 * 1000;

/** Graph @odata.nextLink includes /v1.0; client.api() also prefixes v1.0. Strip it. */
function toGraphApiPath(nextLink: string): string {
  const trimmed = nextLink.trim();
  try {
    const url = trimmed.startsWith("http")
      ? new URL(trimmed)
      : new URL(trimmed, "https://graph.microsoft.com");
    const combined = `${url.pathname}${url.search}`;
    const stripped = combined.replace(/^\/(v1\.0|beta)(?=\/|$)/i, "");
    return stripped.startsWith("/") ? stripped : `/${stripped}`;
  } catch {
    const stripped = trimmed.replace(
      /^https?:\/\/graph\.microsoft\.com/i,
      "",
    ).replace(/^\/(v1\.0|beta)(?=\/|$)/i, "");
    return stripped.startsWith("/") ? stripped : `/${stripped}`;
  }
}

export function pickArtifactForOccurrence<T extends { createdDateTime: string }>(
  artifacts: T[],
  occurrenceStart: string,
  occurrenceEnd?: string,
): T | null {
  const start = parseGraphDate(occurrenceStart);
  if (!start || artifacts.length === 0) {
    return null;
  }

  const end = parseGraphDate(occurrenceEnd ?? "") ?? start;
  const windowStart = start.getTime() - OCCURRENCE_BEFORE_MS;
  const windowEnd = Math.max(end.getTime(), start.getTime()) + OCCURRENCE_AFTER_MS;
  const occurrenceDay = getKsaDateKey(occurrenceStart);

  const inWindow = artifacts.filter((item) => {
    const created = parseGraphDate(item.createdDateTime);
    if (!created) {
      return false;
    }
    const createdAt = created.getTime();
    return createdAt >= windowStart && createdAt <= windowEnd;
  });

  const matches =
    inWindow.length > 0
      ? inWindow
      : artifacts.filter(
          (item) => getKsaDateKey(item.createdDateTime) === occurrenceDay,
        );

  if (matches.length === 0) {
    return null;
  }

  return [...matches].sort((left, right) => {
    const leftTime = parseGraphDate(left.createdDateTime)?.getTime() ?? 0;
    const rightTime = parseGraphDate(right.createdDateTime)?.getTime() ?? 0;
    return (
      Math.abs(leftTime - start.getTime()) - Math.abs(rightTime - start.getTime())
    );
  })[0] ?? null;
}

function toRecordingInfo(item: GraphCallRecording): RecordingInfo | null {
  if (!item.id) {
    return null;
  }

  return {
    id: item.id,
    createdDateTime: item.createdDateTime ?? "",
    endDateTime: item.endDateTime,
    meetingId: item.meetingId,
    recordingContentUrl: item.recordingContentUrl,
  };
}

function isRedirectStatus(status: number): boolean {
  return (
    status === 301 ||
    status === 302 ||
    status === 303 ||
    status === 307 ||
    status === 308
  );
}

function shouldSendGraphToken(url: string): boolean {
  try {
    return new URL(url).hostname.toLowerCase() === "graph.microsoft.com";
  } catch {
    return false;
  }
}

async function fetchRecordingResponse(
  url: string,
  token: string,
): Promise<Response> {
  let currentUrl = url;

  for (let hop = 0; hop < 5; hop += 1) {
    const response = await fetch(currentUrl, {
      headers: shouldSendGraphToken(currentUrl)
        ? { Authorization: `Bearer ${token}` }
        : {},
      redirect: "manual",
    });

    if (!isRedirectStatus(response.status)) {
      return response;
    }

    const location = response.headers.get("location");
    if (!location) {
      return response;
    }

    currentUrl = new URL(location, currentUrl).toString();
  }

  throw new GraphError(
    "Too many redirects while downloading the meeting recording.",
    502,
    "GRAPH_ERROR",
  );
}

let msalClient: ConfidentialClientApplication | null = null;

function getMsalClient(): ConfidentialClientApplication {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: getEnvVar("AZURE_CLIENT_ID"),
        authority: `https://login.microsoftonline.com/${getEnvVar("AZURE_TENANT_ID")}`,
        clientSecret: getEnvVar("AZURE_CLIENT_SECRET"),
      },
    });
  }
  return msalClient;
}

export async function acquireAccessToken(
  options?: { skipCache?: boolean },
): Promise<string> {
  const result = await getMsalClient().acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
    skipCache: options?.skipCache === true,
  });

  if (!result?.accessToken) {
    throw new GraphError(
      "Failed to acquire Microsoft Graph access token.",
      401,
      "AUTH_FAILED",
    );
  }

  return result.accessToken;
}

/** Clears cached Graph/MSAL clients so the next request fetches a fresh token. */
export function resetGraphAuthCache(): void {
  msalClient = null;
  graphClient = null;
  tenantUsersCache = null;
}

/** Drops saved meeting lists and the people directory from memory. */
export function clearGraphDataCaches(): void {
  meetingListCache.clear();
  meetingListInflight.clear();
  resetGraphAuthCache();
}

export function decodeAccessTokenRoles(token: string): string[] {
  try {
    const payload = JSON.parse(
      Buffer.from(token.split(".")[1] ?? "", "base64url").toString(),
    ) as { roles?: string[] };
    return payload.roles ?? [];
  } catch {
    return [];
  }
}

export function tokenHasChatRole(roles: string[]): boolean {
  return (
    roles.includes("Chat.Read.All") || roles.includes("ChatMessage.Read.All")
  );
}

class MsalAuthenticationProvider {
  async getAccessToken(): Promise<string> {
    return acquireAccessToken();
  }
}

let graphClient: Client | null = null;

export function getGraphClient(): Client {
  if (!graphClient) {
    graphClient = Client.initWithMiddleware({
      authProvider: new MsalAuthenticationProvider(),
    });
  }
  return graphClient;
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

export function parseOrganizerIdFromJoinUrl(joinUrl: string): string | null {
  try {
    const context = new URL(joinUrl).searchParams.get("context");
    if (!context) {
      return null;
    }

    const parsed = JSON.parse(decodeURIComponent(context)) as { Oid?: string };
    return parsed.Oid ?? null;
  } catch {
    return null;
  }
}

export function mapGraphError(error: unknown, fallbackMessage: string): GraphError {
  if (error instanceof GraphError) {
    return error;
  }

  const graphError = error as {
    statusCode?: number;
    code?: string;
    message?: string;
    body?: string;
  };

  const statusCode = graphError.statusCode ?? 500;
  const graphCode = graphError.code ?? "GRAPH_ERROR";
  const graphMessage = graphError.message ?? fallbackMessage;

  if (statusCode === 403) {
    const transcriptAccessText = `${graphCode} ${graphMessage} ${graphError.body ?? ""}`;
    if (
      transcriptAccessText.includes("GraphAccessToTranscriptsDisabled") ||
      transcriptAccessText.includes("Graph API access to transcripts is disabled")
    ) {
      return new GraphError(
        "Microsoft Graph access to meeting transcripts is turned off for this organization, so meetings cannot be loaded.",
        403,
        "TRANSCRIPTS_DISABLED",
      );
    }

    const isAccessDenied =
      graphCode === "ErrorAccessDenied" ||
      graphMessage.toLowerCase().includes("access is denied");

    if (isAccessDenied) {
      return new GraphError(
        `${fallbackMessage} Graph returned ErrorAccessDenied. Your Azure app likely lacks admin-consented Application permissions (Calendars.Read, OnlineMeetings.Read.All, OnlineMeetingTranscript.Read.All, OnlineMeetingRecording.Read.All, User.Read.All, Mail.Send). Run scripts/grant-graph-permissions.ps1 or add them in Azure Portal, then restart the dev server.`,
        403,
        "MISSING_APP_PERMISSIONS",
      );
    }

    return new GraphError(
      `${fallbackMessage} ${graphMessage} Ensure admin consent for OnlineMeetingTranscript.Read.All, OnlineMeetingRecording.Read.All, OnlineMeetings.Read.All, Calendars.Read, and Mail.Send application permissions, and that an Application Access Policy is granted to the target user.`,
      403,
      graphCode === "GRAPH_ERROR" ? "FORBIDDEN" : graphCode,
    );
  }

  const paginationText = `${graphCode} ${graphMessage} ${graphError.body ?? ""}`;
  if (
    paginationText.includes("segment 'v1.0'") ||
    paginationText.includes("Resource not found for the segment")
  ) {
    return new GraphError(
      "Microsoft Graph pagination used a duplicated /v1.0 path. Retry Fetch Data.",
      statusCode === 500 ? 400 : statusCode,
      "GRAPH_PAGINATION",
    );
  }

  if (statusCode === 404) {
    return new GraphError(
      "The requested meeting or transcript was not found.",
      404,
      "NOT_FOUND",
    );
  }

  if (statusCode === 429) {
    return new GraphError(
      "Microsoft Graph rate limit exceeded. Please try again shortly.",
      429,
      "RATE_LIMIT",
    );
  }

  const resolvedMessage = (graphError.message ?? "").trim() || fallbackMessage;

  if (statusCode === 401) {
    const isTodoContext = fallbackMessage.toLowerCase().includes("to do");
    return new GraphError(
      isTodoContext
        ? `${resolvedMessage} If you recently granted Tasks.ReadWrite.All, the app refreshes its token automatically. Otherwise run scripts/grant-graph-permissions.ps1 and restart npm run dev.`
        : resolvedMessage,
      401,
      isTodoContext ? "MISSING_TASKS_PERMISSION" : graphCode,
    );
  }

  return new GraphError(resolvedMessage, statusCode, graphCode);
}

async function mapWithConcurrency<T, R>(
  items: T[],
  concurrency: number,
  fn: (item: T) => Promise<R>,
): Promise<R[]> {
  const results: R[] = [];

  for (let index = 0; index < items.length; index += concurrency) {
    const batch = items.slice(index, index + concurrency);
    const batchResults = await Promise.all(batch.map(fn));
    results.push(...batchResults);
  }

  return results;
}

function dedupeMeetings(meetings: MeetingSummary[]): MeetingSummary[] {
  const seen = new Map<string, MeetingSummary>();

  for (const meeting of meetings) {
    const key = `${meeting.joinUrl}|${meeting.start}`;
    const existing = seen.get(key);
    if (!existing) {
      seen.set(key, meeting);
      continue;
    }

    seen.set(key, {
      ...existing,
      hasTranscript: Boolean(existing.hasTranscript || meeting.hasTranscript),
      hasRecording: Boolean(existing.hasRecording || meeting.hasRecording),
      latestTranscriptId:
        existing.latestTranscriptId ?? meeting.latestTranscriptId,
      onlineMeetingId: existing.onlineMeetingId ?? meeting.onlineMeetingId,
      seriesMasterId: existing.seriesMasterId ?? meeting.seriesMasterId,
    });
  }

  return [...seen.values()].sort(
    (left, right) =>
      new Date(right.start).getTime() - new Date(left.start).getTime(),
  );
}

function emptyMeetingCacheFile(): MeetingCacheFile {
  return { entries: {} };
}

function readMeetingCacheFile(): MeetingCacheFile {
  try {
    const parsed = JSON.parse(readFileSync(MEETING_CACHE_PATH, "utf8")) as Partial<MeetingCacheFile>;
    if (!parsed.entries || typeof parsed.entries !== "object") {
      return emptyMeetingCacheFile();
    }
    return { entries: parsed.entries };
  } catch {
    return emptyMeetingCacheFile();
  }
}

function writeMeetingCacheFile(file: MeetingCacheFile): void {
  mkdirSync(path.dirname(MEETING_CACHE_PATH), { recursive: true });
  const tempPath = `${MEETING_CACHE_PATH}.tmp`;
  writeFileSync(tempPath, JSON.stringify(file));
  renameSync(tempPath, MEETING_CACHE_PATH);
}

function isStoredMeetingList(value: StoredMeetingList | undefined): value is StoredMeetingList {
  return Boolean(
    value &&
      typeof value.savedAt === "number" &&
      Array.isArray(value.meetings) &&
      typeof value.transcriptAccessDisabled === "boolean",
  );
}

function rememberMeetingList(cacheKey: string, stored: StoredMeetingList): MeetingListResult {
  meetingListCache.set(cacheKey, {
    expiresAt: stored.savedAt + MEETING_CACHE_TTL_MS,
    value: {
      meetings: stored.meetings,
      transcriptAccessDisabled: stored.transcriptAccessDisabled,
    },
  });
  return {
    meetings: stored.meetings,
    transcriptAccessDisabled: stored.transcriptAccessDisabled,
    fromCache: true,
  };
}

function getCachedMeetingList(cacheKey: string): MeetingListResult | null {
  const cached = meetingListCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) {
    return { ...cached.value, fromCache: true };
  }
  if (cached) {
    meetingListCache.delete(cacheKey);
  }

  const stored = readMeetingCacheFile().entries[cacheKey];
  if (!isStoredMeetingList(stored) || stored.savedAt + MEETING_CACHE_TTL_MS <= Date.now()) {
    return null;
  }
  return rememberMeetingList(cacheKey, stored);
}

function setCachedMeetingList(
  cacheKey: string,
  result: MeetingListResult,
): MeetingListResult {
  const stored: StoredMeetingList = {
    savedAt: Date.now(),
    meetings: result.meetings,
    transcriptAccessDisabled: result.transcriptAccessDisabled,
  };
  meetingListCache.set(cacheKey, {
    expiresAt: stored.savedAt + MEETING_CACHE_TTL_MS,
    value: {
      meetings: stored.meetings,
      transcriptAccessDisabled: stored.transcriptAccessDisabled,
    },
  });

  const file = readMeetingCacheFile();
  const now = Date.now();
  for (const [key, entry] of Object.entries(file.entries)) {
    if (!isStoredMeetingList(entry) || entry.savedAt + MEETING_CACHE_TTL_MS <= now) {
      delete file.entries[key];
    }
  }
  file.entries[cacheKey] = stored;
  try {
    writeMeetingCacheFile(file);
  } catch {
    // The in-memory copy still serves this process when the disk write fails.
  }
  return result;
}

async function getOnlineMeetingDetails(
  organizerId: string,
  meetingId: string,
): Promise<{ subject: string; joinUrl: string } | null> {
  const client = getGraphClient();
  try {
    const meeting = (await client
      .api(`/users/${organizerId}/onlineMeetings/${meetingId}`)
      .select("id,subject,joinWebUrl")
      .get()) as GraphOnlineMeeting;
    if (!meeting.joinWebUrl) {
      return null;
    }
    return {
      subject: meeting.subject?.trim() || "Recorded meeting",
      joinUrl: meeting.joinWebUrl,
    };
  } catch {
    return null;
  }
}

function isTenantUser(value: unknown): value is TenantUser {
  if (!value || typeof value !== "object") {
    return false;
  }
  const user = value as Partial<TenantUser>;
  return (
    typeof user.id === "string" &&
    user.id.length > 0 &&
    typeof user.displayName === "string" &&
    user.displayName.length > 0 &&
    typeof user.userPrincipalName === "string" &&
    user.userPrincipalName.length > 0
  );
}

function readTenantUsersDiskCache(allowStale = false): TenantUser[] | null {
  try {
    const parsed = JSON.parse(readFileSync(USER_CACHE_PATH, "utf8")) as {
      version?: unknown;
      savedAt?: unknown;
      users?: unknown;
    };
    const stale =
      typeof parsed.savedAt !== "number" ||
      parsed.savedAt + USER_DISK_CACHE_TTL_MS <= Date.now();
    if (
      parsed.version !== USER_CACHE_VERSION ||
      typeof parsed.savedAt !== "number" ||
      (!allowStale && stale) ||
      !Array.isArray(parsed.users)
    ) {
      return null;
    }
    const users = parsed.users.filter(isTenantUser);
    return users.length > 0 ? users : null;
  } catch {
    return null;
  }
}

function writeTenantUsersDiskCache(users: TenantUser[]): void {
  try {
    mkdirSync(path.dirname(USER_CACHE_PATH), { recursive: true });
    const tempPath = `${USER_CACHE_PATH}.tmp`;
    writeFileSync(
      tempPath,
      JSON.stringify({ version: USER_CACHE_VERSION, savedAt: Date.now(), users }),
    );
    renameSync(tempPath, USER_CACHE_PATH);
  } catch {
    // The in-memory copy still serves this process when the disk write fails.
  }
}

function rememberTenantUsers(users: TenantUser[]): TenantUser[] {
  tenantUsersCache = {
    expiresAt: Date.now() + USER_CACHE_TTL_MS,
    value: users,
  };
  return users;
}

let tenantUsersRefresh: Promise<TenantUser[]> | null = null;

async function refreshTenantUsersFromGraph(): Promise<TenantUser[]> {
  const client = getGraphClient();
  const collected: GraphUser[] = [];

  try {
    // Filtering plus ordering this list makes Graph return Request_UnsupportedQuery.
    // Page the directory, then keep enabled members in code.
    let nextPath = "/users";
    let firstPage = true;
    while (nextPath) {
      const response = (
        firstPage
          ? await client
              .api(nextPath)
              .select("id,displayName,userPrincipalName,mail,accountEnabled,userType")
              .top(999)
              .get()
          : await client.api(nextPath).get()
      ) as GraphUsersResponse;
      firstPage = false;
      collected.push(...(response.value ?? []));
      nextPath = response["@odata.nextLink"]
        ? toGraphApiPath(response["@odata.nextLink"])
        : "";
    }
  } catch {
    const stale = readTenantUsersDiskCache(true);
    if (stale) {
      return rememberTenantUsers(stale);
    }
  }

  let users: TenantUser[] = collected
    .filter(
      (user): user is GraphUser & { id: string; userPrincipalName: string } =>
        Boolean(user.id && user.userPrincipalName) &&
        user.accountEnabled !== false &&
        (user.userType ?? "Member") === "Member",
    )
    .map((user) => ({
      id: user.id,
      displayName: user.displayName ?? user.userPrincipalName,
      userPrincipalName: user.userPrincipalName,
      ...(typeof user.mail === "string" && user.mail.includes("@")
        ? { mail: user.mail }
        : {}),
    }))
    .sort((left, right) => left.displayName.localeCompare(right.displayName));

  if (users.length === 0) {
    const targetUserId = getEnvVar("TARGET_USER_ID");
    users = [
      {
        id: targetUserId,
        displayName: targetUserId,
        userPrincipalName: targetUserId,
      },
    ];
    return rememberTenantUsers(users);
  }

  writeTenantUsersDiskCache(users);
  return rememberTenantUsers(users);
}

export async function listTenantUsers(): Promise<TenantUser[]> {
  if (tenantUsersCache && tenantUsersCache.expiresAt > Date.now()) {
    return tenantUsersCache.value;
  }

  const cachedUsers = readTenantUsersDiskCache();
  if (cachedUsers) {
    return rememberTenantUsers(cachedUsers);
  }

  const staleUsers = readTenantUsersDiskCache(true);
  if (staleUsers) {
    if (!tenantUsersRefresh) {
      tenantUsersRefresh = refreshTenantUsersFromGraph().finally(() => {
        tenantUsersRefresh = null;
      });
    }
    return rememberTenantUsers(staleUsers);
  }

  if (!tenantUsersRefresh) {
    tenantUsersRefresh = refreshTenantUsersFromGraph().finally(() => {
      tenantUsersRefresh = null;
    });
  }
  return tenantUsersRefresh;
}

export async function listLicensedTenantUsers(): Promise<TenantUser[]> {
  const client = getGraphClient();
  const collected: GraphUser[] = [];

  try {
    let nextPath = "/users";
    let firstPage = true;
    while (nextPath) {
      const response = (
        firstPage
          ? await client
              .api(nextPath)
              .select(
                "id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedLicenses",
              )
              .top(999)
              .get()
          : await client.api(nextPath).get()
      ) as GraphUsersResponse;
      firstPage = false;
      collected.push(...(response.value ?? []));
      nextPath = response["@odata.nextLink"]
        ? toGraphApiPath(response["@odata.nextLink"])
        : "";
    }
  } catch {
    return listTenantUsers();
  }

  const licensed = collected
    .filter(
      (user): user is GraphUser & { id: string; userPrincipalName: string } =>
        Boolean(user.id && user.userPrincipalName) &&
        user.accountEnabled !== false &&
        (user.userType ?? "Member") === "Member" &&
        (user.assignedLicenses ?? []).length > 0,
    )
    .map((user) => ({
      id: user.id,
      displayName: user.displayName ?? user.userPrincipalName,
      userPrincipalName: user.userPrincipalName,
      ...(typeof user.mail === "string" && user.mail.includes("@")
        ? { mail: user.mail }
        : {}),
    }))
    .sort((left, right) => left.displayName.localeCompare(right.displayName));

  return licensed.length > 0 ? licensed : listTenantUsers();
}

export async function getRecentMeetings(
  userId: string,
  daysBack = 30,
  userDisplayName = userId,
): Promise<MeetingSummary[]> {
  const client = getGraphClient();
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);

  try {
    const response = (await client
      .api(`/users/${userId}/calendarView`)
      .query({
        startDateTime: startDate.toISOString(),
        endDateTime: endDate.toISOString(),
      })
      .header("Prefer", 'outlook.timezone="UTC"')
      .select("id,subject,start,end,isOnlineMeeting,onlineMeeting,seriesMasterId")
      .orderby("start/dateTime desc")
      .top(50)
      .get()) as GraphCalendarViewResponse;

    const meetings: MeetingSummary[] = [];

    for (const event of response.value ?? []) {
      const joinUrl = event.onlineMeeting?.joinUrl;
      if (!event.isOnlineMeeting || !joinUrl || !event.id) {
        continue;
      }

      meetings.push({
        eventId: event.id,
        subject: event.subject ?? "Untitled Meeting",
        start: event.start?.dateTime ?? "",
        end: event.end?.dateTime ?? "",
        joinUrl,
        userId,
        userDisplayName,
        organizerId: parseOrganizerIdFromJoinUrl(joinUrl) ?? userId,
        seriesMasterId: event.seriesMasterId,
        hasTranscript: false,
        hasRecording: false,
      });
    }

    return meetings;
  } catch (error) {
    throw mapGraphError(error, "Failed to fetch recent meetings.");
  }
}

async function getOrganizerTranscriptIndex(
  organizerId: string,
  startDate: Date,
  endDate: Date,
): Promise<Map<string, TranscriptInfo[]>> {
  const client = getGraphClient();
  const index = new Map<string, TranscriptInfo[]>();

  const startIso = startDate.toISOString();
  const endIso = endDate.toISOString();
  let nextPath =
    `/users/${organizerId}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='${organizerId}',startDateTime=${startIso},endDateTime=${endIso})`;

  try {
    while (nextPath) {
      const response = (await client.api(nextPath).get()) as GraphGetAllTranscriptsResponse;

      for (const item of response.value ?? []) {
        if (!item.meetingId || !item.id) {
          continue;
        }

        const list = index.get(item.meetingId) ?? [];
        list.push({
          id: item.id,
          createdDateTime: item.createdDateTime ?? "",
        });
        index.set(item.meetingId, list);
      }

      if (response["@odata.nextLink"]) {
        nextPath = toGraphApiPath(response["@odata.nextLink"]);
      } else {
        nextPath = "";
      }
    }
  } catch (error) {
    const mapped = mapGraphError(error, "Failed to list meeting transcripts.");
    if (mapped.code === "TRANSCRIPTS_DISABLED") {
      transcriptsDisabledKnown = true;
      throw mapped;
    }
    return index;
  }

  return index;
}

async function getOrganizerRecordingIndex(
  organizerId: string,
  startDate: Date,
  endDate: Date,
): Promise<Map<string, RecordingInfo[]>> {
  const client = getGraphClient();
  const index = new Map<string, RecordingInfo[]>();
  const startIso = startDate.toISOString();
  const endIso = endDate.toISOString();
  let nextPath =
    `/users/${organizerId}/onlineMeetings/getAllRecordings(meetingOrganizerUserId='${organizerId}',startDateTime=${startIso},endDateTime=${endIso})`;

  try {
    while (nextPath) {
      const response = (await client
        .api(nextPath)
        .get()) as GraphGetAllRecordingsResponse;

      for (const item of response.value ?? []) {
        const recording = toRecordingInfo(item);
        if (!recording || !item.meetingId) {
          continue;
        }
        const list = index.get(item.meetingId) ?? [];
        list.push({ ...recording, meetingId: item.meetingId });
        index.set(item.meetingId, list);
      }

      nextPath = response["@odata.nextLink"]
        ? toGraphApiPath(response["@odata.nextLink"])
        : "";
    }
  } catch {
    return index;
  }

  return index;
}

function latestArtifact<T extends { createdDateTime: string }>(
  artifacts: T[],
): T | null {
  if (artifacts.length === 0) {
    return null;
  }
  return [...artifacts].sort((left, right) => {
    const leftTime = parseGraphDate(left.createdDateTime)?.getTime() ?? 0;
    const rightTime = parseGraphDate(right.createdDateTime)?.getTime() ?? 0;
    return rightTime - leftTime;
  })[0] ?? null;
}

async function enrichMeetingsWithArtifacts(
  meetings: MeetingSummary[],
  daysBack: number,
  options?: { transcriptsDisabled?: boolean },
): Promise<MeetingSummary[]> {
  if (meetings.length === 0) {
    return [];
  }

  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);

  const organizerIds = [...new Set(meetings.map((meeting) => meeting.organizerId))];
  const transcriptsByOrganizer = new Map<string, Map<string, TranscriptInfo[]>>();
  const recordingsByOrganizer = new Map<string, Map<string, RecordingInfo[]>>();
  const meetingIdByJoinUrl = new Map<string, string>();

  const [transcriptIndexes, recordingIndexes] = await Promise.all([
    options?.transcriptsDisabled
      ? Promise.resolve([] as Array<{ organizerId: string; index: Map<string, TranscriptInfo[]> }>)
      : mapWithConcurrency(organizerIds, 3, async (organizerId) => {
          try {
            const index = await getOrganizerTranscriptIndex(
              organizerId,
              startDate,
              endDate,
            );
            return { organizerId, index };
          } catch (error) {
            if (error instanceof GraphError && error.code === "TRANSCRIPTS_DISABLED") {
              throw error;
            }
            return {
              organizerId,
              index: new Map<string, TranscriptInfo[]>(),
            };
          }
        }),
    mapWithConcurrency(organizerIds, 3, async (organizerId) => {
      const index = await getOrganizerRecordingIndex(
        organizerId,
        startDate,
        endDate,
      );
      return { organizerId, index };
    }),
  ]);

  for (const { organizerId, index } of transcriptIndexes) {
    transcriptsByOrganizer.set(organizerId, index);
  }
  for (const { organizerId, index } of recordingIndexes) {
    recordingsByOrganizer.set(organizerId, index);
  }

  const usedTranscriptIds = new Set<string>();
  const usedRecordingIds = new Set<string>();

  const annotated = await mapWithConcurrency(meetings, 5, async (meeting) => {
    try {
      const cachedId = meeting.onlineMeetingId ?? meetingIdByJoinUrl.get(meeting.joinUrl);
      const onlineMeetingId =
        cachedId ??
        (await resolveOnlineMeetingId(meeting.organizerId, meeting.joinUrl));
      meetingIdByJoinUrl.set(meeting.joinUrl, onlineMeetingId);

      const transcriptIndex =
        transcriptsByOrganizer.get(meeting.organizerId)?.get(onlineMeetingId) ?? [];
      const recordingIndex =
        recordingsByOrganizer.get(meeting.organizerId)?.get(onlineMeetingId) ?? [];

      const transcripts =
        transcriptIndex.length > 0 || options?.transcriptsDisabled
          ? transcriptIndex
          : await listTranscripts(meeting.organizerId, onlineMeetingId);
      const recordings =
        recordingIndex.length > 0
          ? recordingIndex
          : await listRecordings(meeting.organizerId, onlineMeetingId);

      const transcriptMatch = pickArtifactForOccurrence(
        transcripts,
        meeting.start,
        meeting.end,
      );
      const recordingMatch = pickArtifactForOccurrence(
        recordings,
        meeting.start,
        meeting.end,
      );

      if (transcriptMatch?.id) {
        usedTranscriptIds.add(transcriptMatch.id);
      }
      if (recordingMatch?.id) {
        usedRecordingIds.add(recordingMatch.id);
      }

      return {
        ...meeting,
        onlineMeetingId,
        latestTranscriptId: transcriptMatch?.id,
        hasTranscript: Boolean(transcriptMatch),
        hasRecording: Boolean(recordingMatch),
      };
    } catch {
      return { ...meeting, hasTranscript: false, hasRecording: false };
    }
  });

  const templateByMeetingId = new Map<string, MeetingSummary>();
  for (const meeting of annotated) {
    if (meeting.onlineMeetingId && !templateByMeetingId.has(meeting.onlineMeetingId)) {
      templateByMeetingId.set(meeting.onlineMeetingId, meeting);
    }
  }

  type UnmatchedGroup = {
    organizerId: string;
    meetingId: string;
    day: string;
    recordings: RecordingInfo[];
    transcripts: TranscriptInfo[];
  };
  const unmatched = new Map<string, UnmatchedGroup>();

  const addUnmatched = (
    organizerId: string,
    meetingId: string,
    createdDateTime: string,
    kind: "recording" | "transcript",
    recording?: RecordingInfo,
    transcript?: TranscriptInfo,
  ): void => {
    const day = getKsaDateKey(createdDateTime);
    if (!day || day === "unknown") {
      return;
    }
    const key = `${organizerId}|${meetingId}|${day}`;
    const group = unmatched.get(key) ?? {
      organizerId,
      meetingId,
      day,
      recordings: [],
      transcripts: [],
    };
    if (kind === "recording" && recording) {
      group.recordings.push(recording);
    }
    if (kind === "transcript" && transcript) {
      group.transcripts.push(transcript);
    }
    unmatched.set(key, group);
  };

  for (const organizerId of organizerIds) {
    const recordings = recordingsByOrganizer.get(organizerId);
    if (recordings) {
      for (const [meetingId, items] of recordings) {
        for (const item of items) {
          if (!usedRecordingIds.has(item.id)) {
            addUnmatched(organizerId, meetingId, item.createdDateTime, "recording", item);
          }
        }
      }
    }
    const transcripts = transcriptsByOrganizer.get(organizerId);
    if (transcripts) {
      for (const [meetingId, items] of transcripts) {
        for (const item of items) {
          if (!usedTranscriptIds.has(item.id)) {
            addUnmatched(
              organizerId,
              meetingId,
              item.createdDateTime,
              "transcript",
              undefined,
              item,
            );
          }
        }
      }
    }
  }

  const mailbox = meetings[0];
  const extraMeetings: MeetingSummary[] = [];
  const detailsCache = new Map<string, { subject: string; joinUrl: string } | null>();

  for (const group of unmatched.values()) {
    const recording = latestArtifact(group.recordings);
    const transcript = latestArtifact(group.transcripts);
    if (!recording && !transcript) {
      continue;
    }

    const template = templateByMeetingId.get(group.meetingId);
    let subject = template?.subject;
    let joinUrl = template?.joinUrl;
    if (!joinUrl) {
      const cacheKey = `${group.organizerId}|${group.meetingId}`;
      let details = detailsCache.get(cacheKey);
      if (details === undefined) {
        details = await getOnlineMeetingDetails(group.organizerId, group.meetingId);
        detailsCache.set(cacheKey, details);
      }
      subject = details?.subject;
      joinUrl = details?.joinUrl;
    }
    if (!joinUrl || !subject) {
      continue;
    }

    const start =
      recording?.createdDateTime ||
      transcript?.createdDateTime ||
      "";
    const end = recording?.endDateTime || start;
    extraMeetings.push({
      eventId: `artifact:${group.meetingId}:${group.day}`,
      subject,
      start,
      end,
      joinUrl,
      userId: mailbox.userId,
      userDisplayName: mailbox.userDisplayName,
      organizerId: group.organizerId,
      seriesMasterId: template?.seriesMasterId,
      onlineMeetingId: group.meetingId,
      latestTranscriptId: transcript?.id,
      hasTranscript: Boolean(transcript),
      hasRecording: Boolean(recording),
    });
  }

  return [...annotated, ...extraMeetings];
}

export async function getRecentMeetingsWithTranscripts(
  userId: string,
  daysBack = 30,
  userDisplayName = userId,
): Promise<MeetingListResult> {
  const cacheKey = meetingListCacheKey(userId, daysBack);
  const cached = getCachedMeetingList(cacheKey);
  if (cached) {
    return cached;
  }

  const pending = meetingListInflight.get(cacheKey);
  if (pending) {
    return pending;
  }

  const load = loadRecentMeetingsWithTranscripts(
    cacheKey,
    userId,
    daysBack,
    userDisplayName,
  );
  meetingListInflight.set(cacheKey, load);
  try {
    return await load;
  } finally {
    meetingListInflight.delete(cacheKey);
  }
}

async function loadRecentMeetingsWithTranscripts(
  cacheKey: string,
  userId: string,
  daysBack: number,
  userDisplayName: string,
): Promise<MeetingListResult> {
  const meetings = await getRecentMeetings(userId, daysBack, userDisplayName);

  const withArtifacts = async (
    transcriptsDisabled: boolean,
  ): Promise<MeetingSummary[]> => {
    const annotated = await enrichMeetingsWithArtifacts(meetings, daysBack, {
      transcriptsDisabled,
    });
    return dedupeMeetings(
      annotated.filter(
        (meeting) => meeting.hasTranscript || meeting.hasRecording,
      ),
    );
  };

  if (transcriptsDisabledKnown) {
    return setCachedMeetingList(cacheKey, {
      meetings: await withArtifacts(true),
      transcriptAccessDisabled: true,
    });
  }

  try {
    return setCachedMeetingList(cacheKey, {
      meetings: await withArtifacts(false),
      transcriptAccessDisabled: false,
    });
  } catch (error) {
    if (error instanceof GraphError && error.code === "TRANSCRIPTS_DISABLED") {
      return setCachedMeetingList(cacheKey, {
        meetings: await withArtifacts(true),
        transcriptAccessDisabled: true,
      });
    }
    throw error;
  }
}

export async function getAllRecentMeetings(
  daysBack = 30,
): Promise<MeetingListResult> {
  const targetUserId = getEnvVar("TARGET_USER_ID");
  const users = await listTenantUsers();
  const target = users.find(
    (user) =>
      user.userPrincipalName.toLowerCase() === targetUserId.toLowerCase() ||
      user.id.toLowerCase() === targetUserId.toLowerCase(),
  );

  // One mailbox. Scanning every directory user leaves Fetch Data waiting on Graph.
  return getRecentMeetingsWithTranscripts(
    target?.userPrincipalName ?? targetUserId,
    daysBack,
    target?.displayName ?? targetUserId,
  );
}

export async function resolveOnlineMeetingId(
  graphUserId: string,
  joinUrl: string,
): Promise<string> {
  const client = getGraphClient();
  const filter = `JoinWebUrl eq '${escapeODataString(joinUrl)}'`;

  try {
    const response = (await client
      .api(`/users/${graphUserId}/onlineMeetings`)
      .filter(filter)
      .get()) as GraphOnlineMeetingsResponse;

    const meetingId = response.value?.[0]?.id;
    if (!meetingId) {
      throw new GraphError(
        "Could not resolve the online meeting from the join URL. The app may need an Application Access Policy for the meeting organizer.",
        404,
        "MEETING_NOT_FOUND",
      );
    }

    return meetingId;
  } catch (error) {
    throw mapGraphError(error, "Failed to resolve online meeting.");
  }
}

function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function chatImageNames(html: string, attachments: GraphChatAttachment[]): string[] {
  const names: string[] = [];
  const seen = new Set<string>();
  const add = (value: string): void => {
    const name = value.replace(/\s+/g, " ").trim();
    if (!name || seen.has(name.toLowerCase())) {
      return;
    }
    seen.add(name.toLowerCase());
    names.push(name.slice(0, 80));
  };

  for (const attachment of attachments) {
    const name = attachment.name?.trim() ?? "";
    if (name && isImageAttachmentName(name, attachment.contentType)) {
      add(name);
    }
  }

  const imageTags = html.match(/<img\b[^>]*>/gi) ?? [];
  for (const tag of imageTags) {
    const alt = tag.match(/\balt="([^"]*)"/i)?.[1]?.trim() ?? "";
    const label = tag.match(/\baria-label="([^"]*)"/i)?.[1]?.trim() ?? "";
    add(alt || label || "Screenshot");
  }

  return names;
}

function formatChatLine(entry: {
  time: string;
  speaker: string;
  text: string;
  webUrl: string;
  images: string[];
}): string {
  const stamp = entry.time ? `${formatKsaDateTime(entry.time)} ` : "";
  const links = formatChatSourceLinks(entry.webUrl, entry.images);
  const body = entry.text || (entry.images.length > 0 ? "" : "");
  return `${stamp}${entry.speaker}: ${body} ${links}`.replace(/\s+/g, " ").trim();
}

export async function fetchMeetingChat(
  organizerId: string,
  joinUrl: string,
  knownMeetingId?: string,
  ksaDateKey?: string,
): Promise<MeetingChatResult> {
  const client = getGraphClient();
  const meetingId =
    knownMeetingId ?? (await resolveOnlineMeetingId(organizerId, joinUrl));

  try {
    const meeting = (await client
      .api(`/users/${organizerId}/onlineMeetings/${meetingId}`)
      .select("id,chatInfo")
      .get()) as GraphOnlineMeeting;

    const threadId = meeting.chatInfo?.threadId;
    if (!threadId) {
      return { chat: "", messageCount: 0, status: "empty" };
    }

    const entries: Array<{
      time: string;
      speaker: string;
      text: string;
      webUrl: string;
      images: string[];
    }> = [];
    let nextPath = `/chats/${threadId}/messages`;
    let useTop = true;
    while (nextPath && entries.length < 500) {
      const request = useTop
        ? client.api(nextPath).top(50)
        : client.api(nextPath);
      let response: GraphChatMessagesResponse;
      try {
        response = (await request.get()) as GraphChatMessagesResponse;
      } catch (pageError) {
        const mapped = mapGraphError(pageError, "Failed to fetch meeting chat.");
        if (entries.length > 0 && mapped.code === "GRAPH_PAGINATION") {
          break;
        }
        throw pageError;
      }

      let pageOldestKey = "";
      for (const message of response.value ?? []) {
        if (message.deletedDateTime) {
          continue;
        }
        if (message.messageType && message.messageType !== "message") {
          continue;
        }

        const html = message.body?.content ?? "";
        const text = stripHtml(html);
        const images = chatImageNames(html, message.attachments ?? []);
        const rawUrl = message.webUrl?.trim() ?? "";
        const permalink = rawUrl ? toTeamsMessagePermalink(rawUrl) : "";
        const webUrl = isSafeHttpUrl(permalink) ? permalink : "";
        if (!text && images.length === 0) {
          continue;
        }

        const time = message.createdDateTime ?? "";
        const dayKey = time ? getKsaDateKey(time) : "";
        if (dayKey && (!pageOldestKey || dayKey < pageOldestKey)) {
          pageOldestKey = dayKey;
        }
        entries.push({
          time,
          speaker: message.from?.user?.displayName ?? "Unknown",
          text,
          webUrl,
          images,
        });
      }

      if (response["@odata.nextLink"]) {
        nextPath = toGraphApiPath(response["@odata.nextLink"]);
        useTop = false;
      } else {
        nextPath = "";
      }

      if (ksaDateKey && pageOldestKey && pageOldestKey < ksaDateKey) {
        nextPath = "";
      }
    }

    entries.sort(
      (left, right) =>
        new Date(left.time).getTime() - new Date(right.time).getTime(),
    );

    const sameDay = ksaDateKey
      ? entries.filter((entry) => getKsaDateKey(entry.time) === ksaDateKey)
      : entries;

    const chat = sameDay.map((entry) => formatChatLine(entry)).join("\n");

    return {
      chat,
      messageCount: sameDay.length,
      status: sameDay.length > 0 ? "ok" : "empty",
    };
  } catch (error) {
    const mapped = mapGraphError(error, "Failed to fetch meeting chat.");
    if (mapped.statusCode === 403) {
      const roles = decodeAccessTokenRoles(await acquireAccessToken());
      const hasChatRole = tokenHasChatRole(roles);
      return {
        chat: "",
        messageCount: 0,
        status: "forbidden",
        errorCode: "CHAT_FORBIDDEN",
        errorMessage: hasChatRole
          ? "Microsoft Graph denied this meeting chat even though Chat.Read.All is on the token. App-only access cannot read chats that include guests or people from other tenants."
          : "This app's Graph token does not include Chat.Read.All. Confirm those Application permissions on the same app as AZURE_CLIENT_ID, grant admin consent, then open Graph health and click Refresh token.",
      };
    }
    if (
      mapped.statusCode === 404 ||
      mapped.code === "TRANSCRIPTS_DISABLED" ||
      mapped.code === "GRAPH_PAGINATION"
    ) {
      return {
        chat: "",
        messageCount: 0,
        status: "unavailable",
        errorCode: mapped.code,
        errorMessage: mapped.message,
      };
    }
    throw mapped;
  }
}

export async function listTranscripts(
  userId: string,
  meetingId: string,
): Promise<TranscriptInfo[]> {
  const client = getGraphClient();
  const collected: TranscriptInfo[] = [];
  let nextPath = `/users/${userId}/onlineMeetings/${meetingId}/transcripts`;

  try {
    while (nextPath) {
      const response = (await client
        .api(nextPath)
        .get()) as GraphTranscriptsResponse;

      for (const item of response.value ?? []) {
        if (!item.id) {
          continue;
        }
        collected.push({
          id: item.id,
          createdDateTime: item.createdDateTime ?? "",
        });
      }

      nextPath = response["@odata.nextLink"]
        ? toGraphApiPath(response["@odata.nextLink"])
        : "";
    }

    return collected.sort(
      (a, b) =>
        new Date(b.createdDateTime).getTime() -
        new Date(a.createdDateTime).getTime(),
    );
  } catch (error) {
    throw mapGraphError(error, "Failed to list meeting transcripts.");
  }
}

export async function getTranscriptContent(
  userId: string,
  meetingId: string,
  transcriptId: string,
): Promise<string> {
  const token = await acquireAccessToken();
  const url =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}` +
    `/onlineMeetings/${encodeURIComponent(meetingId)}` +
    `/transcripts/${encodeURIComponent(transcriptId)}/content?$format=text/vtt`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "text/vtt",
      },
    });

    if (!response.ok) {
      let message = "Failed to fetch transcript content.";
      try {
        const errorBody = (await response.json()) as {
          error?: { code?: string; message?: string };
        };
        message = errorBody.error?.message ?? message;
        throw new GraphError(message, response.status, errorBody.error?.code ?? "GRAPH_ERROR");
      } catch (parseError) {
        if (parseError instanceof GraphError) {
          throw parseError;
        }
        throw new GraphError(message, response.status, "GRAPH_ERROR");
      }
    }

    return await response.text();
  } catch (error) {
    throw mapGraphError(error, "Failed to fetch transcript content.");
  }
}

function stripVttVoiceTags(text: string): string {
  return text
    .replace(/<v\s+([^>]+)>/gi, "$1: ")
    .replace(/<\/v>/gi, "")
    .replace(/<[^>]+>/g, "")
    .trim();
}

function formatCueTimestamp(value: string): string {
  const match = value.trim().match(/^(\d{2}:\d{2}:\d{2})/);
  return match?.[1] ?? value.trim();
}

export function parseVtt(raw: string): string {
  const lines = raw.split(/\r?\n/);
  const transcriptLines: string[] = [];
  let currentSpeaker = "";
  let currentText = "";
  let currentTime = "";

  const flushCue = (): void => {
    const cleaned = stripVttVoiceTags(currentText);
    if (!cleaned) {
      currentSpeaker = "";
      currentText = "";
      return;
    }

    const stamp = currentTime ? `[${currentTime}] ` : "";
    if (currentSpeaker) {
      transcriptLines.push(`${stamp}${currentSpeaker}: ${cleaned}`);
    } else {
      transcriptLines.push(`${stamp}${cleaned}`);
    }

    currentSpeaker = "";
    currentText = "";
  };

  for (const line of lines) {
    const trimmed = line.trim();

    if (!trimmed || trimmed === "WEBVTT" || trimmed.startsWith("NOTE")) {
      continue;
    }

    if (/^\d+$/.test(trimmed)) {
      continue;
    }

    const timingMatch = trimmed.match(
      /^(\d{2}:\d{2}:\d{2}\.\d{3})\s+-->\s+(\d{2}:\d{2}:\d{2}\.\d{3})/,
    );
    if (timingMatch) {
      flushCue();
      currentTime = formatCueTimestamp(timingMatch[1]);
      continue;
    }

    const voiceMatch = trimmed.match(/^<v\s+([^>]+)>(.*)$/i);
    if (voiceMatch) {
      flushCue();
      currentSpeaker = voiceMatch[1].trim();
      currentText = voiceMatch[2];
      flushCue();
      continue;
    }

    if (currentText) {
      currentText += ` ${trimmed}`;
    } else {
      currentText = trimmed;
    }
  }

  flushCue();

  return transcriptLines.join("\n");
}

export async function fetchLatestParsedTranscript(
  joinUrl: string,
  organizerId: string,
  knownMeetingId?: string,
  knownTranscriptId?: string,
  occurrenceStart?: string,
  occurrenceEnd?: string,
): Promise<{
  transcript: string;
  transcriptId: string;
  meetingId: string;
  createdDateTime: string;
}> {
  const meetingId =
    knownMeetingId ?? (await resolveOnlineMeetingId(organizerId, joinUrl));

  const transcripts = await listTranscripts(organizerId, meetingId);
  const occurrenceMatch = occurrenceStart
    ? pickArtifactForOccurrence(transcripts, occurrenceStart, occurrenceEnd)
    : null;
  const hinted =
    knownTranscriptId &&
    occurrenceMatch &&
    occurrenceMatch.id === knownTranscriptId
      ? occurrenceMatch
      : null;
  const selected = hinted ?? occurrenceMatch;

  if (!selected?.id) {
    throw new GraphError(
      occurrenceStart
        ? "No transcript is available for this meeting occurrence. Recurring series share a join URL; this date had no transcript of its own."
        : "No transcript is available for this meeting. Ensure transcription was enabled during the meeting.",
      404,
      "NO_TRANSCRIPT",
    );
  }

  const rawContent = await getTranscriptContent(
    organizerId,
    meetingId,
    selected.id,
  );
  const transcript = parseVtt(rawContent);

  if (!transcript.trim()) {
    throw new GraphError(
      "Transcript content was empty after parsing.",
      404,
      "EMPTY_TRANSCRIPT",
    );
  }

  return {
    transcript,
    transcriptId: selected.id,
    meetingId,
    createdDateTime: selected.createdDateTime,
  };
}

export async function listRecordings(
  userId: string,
  meetingId: string,
): Promise<RecordingInfo[]> {
  const client = getGraphClient();
  const collected: RecordingInfo[] = [];
  let nextPath = `/users/${userId}/onlineMeetings/${meetingId}/recordings`;

  try {
    while (nextPath) {
      const response = (await client
        .api(nextPath)
        .get()) as GraphRecordingsResponse;

      for (const item of response.value ?? []) {
        const recording = toRecordingInfo(item);
        if (!recording) {
          continue;
        }
        collected.push(recording);
      }

      nextPath = response["@odata.nextLink"]
        ? toGraphApiPath(response["@odata.nextLink"])
        : "";
    }

    return collected.sort(
      (a, b) =>
        new Date(b.createdDateTime).getTime() -
        new Date(a.createdDateTime).getTime(),
    );
  } catch (error) {
    throw mapGraphError(error, "Failed to list meeting recordings.");
  }
}

export async function getRecordingContent(
  userId: string,
  meetingId: string,
  recordingId: string,
  contentUrl?: string,
): Promise<RecordingContent> {
  const token = await acquireAccessToken();
  const url =
    contentUrl?.trim() ||
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}` +
      `/onlineMeetings/${encodeURIComponent(meetingId)}` +
      `/recordings/${encodeURIComponent(recordingId)}/content`;

  try {
    // Graph may return the bytes directly or 302 to OneDrive/SharePoint.
    // Keep the Graph bearer token only for graph.microsoft.com hops.
    const response = await fetchRecordingResponse(url, token);

    if (!response.ok) {
      let message = "Failed to fetch recording content.";
      try {
        const errorBody = (await response.json()) as {
          error?: { code?: string; message?: string };
        };
        message = errorBody.error?.message ?? message;
        throw new GraphError(
          message,
          response.status,
          errorBody.error?.code ?? "GRAPH_ERROR",
        );
      } catch (parseError) {
        if (parseError instanceof GraphError) {
          throw parseError;
        }
        throw new GraphError(message, response.status, "GRAPH_ERROR");
      }
    }

    const declaredLength = Number(response.headers.get("content-length") ?? 0);
    if (declaredLength > MAX_RECORDING_BYTES) {
      throw new GraphError(
        `Meeting recording is ${Math.round(declaredLength / (1024 * 1024))} MB, which exceeds the ${Math.round(MAX_RECORDING_BYTES / (1024 * 1024))} MB analyzer limit. Uncheck video or analyze transcript and chat only.`,
        413,
        "VIDEO_TOO_LARGE",
      );
    }

    const mimeType =
      response.headers.get("content-type")?.split(";")[0].trim() ||
      "video/mp4";

    if (!response.body) {
      const arrayBuffer = await response.arrayBuffer();
      if (arrayBuffer.byteLength > MAX_RECORDING_BYTES) {
        throw new GraphError(
          "Meeting recording exceeds the analyzer size limit.",
          413,
          "VIDEO_TOO_LARGE",
        );
      }
      return { data: Buffer.from(arrayBuffer), mimeType };
    }

    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;

    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      total += value.byteLength;
      if (total > MAX_RECORDING_BYTES) {
        await reader.cancel();
        throw new GraphError(
          `Meeting recording exceeded the ${Math.round(MAX_RECORDING_BYTES / (1024 * 1024))} MB analyzer limit while downloading.`,
          413,
          "VIDEO_TOO_LARGE",
        );
      }
      chunks.push(value);
    }

    return { data: Buffer.concat(chunks), mimeType };
  } catch (error) {
    throw mapGraphError(error, "Failed to fetch recording content.");
  }
}

export async function fetchLatestRecording(
  joinUrl: string,
  organizerId: string,
  knownMeetingId?: string,
  occurrenceStart?: string,
  occurrenceEnd?: string,
): Promise<RecordingContent & { recordingId: string; meetingId: string }> {
  const meetingId =
    knownMeetingId ?? (await resolveOnlineMeetingId(organizerId, joinUrl));

  const recordings = await listRecordings(organizerId, meetingId);
  const selected = occurrenceStart
    ? pickArtifactForOccurrence(recordings, occurrenceStart, occurrenceEnd)
    : recordings[0];

  if (!selected?.id) {
    throw new GraphError(
      "No recording is available for this meeting. Ensure the meeting was recorded.",
      404,
      "NO_RECORDING",
    );
  }

  const content = await getRecordingContent(
    organizerId,
    meetingId,
    selected.id,
    selected.recordingContentUrl,
  );

  return { ...content, recordingId: selected.id, meetingId };
}
