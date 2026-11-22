import "server-only";

import { randomUUID } from "node:crypto";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { CodeQcSendTicket, CodeQcTeamsResult } from "@/types/code-qc";
import type { MeetingActionItem, TeamsSendResult, TenantUser } from "@/types/meeting";
import { renderMarkdownLinksAsHtml, retainCrucialSourceLinks } from "@/lib/chat-links";
import { CODE_QC_MORE_BATCH, VISIBLE_CODE_QC_TICKETS, displayBoardName } from "@/lib/code-qc-brief";
import { matchAssignee } from "@/lib/assignees";
import { normalizeDevTag } from "@/lib/dev-tags";
import { AppError, GraphError } from "@/lib/errors";
import { getGraphClient, listTenantUsers } from "@/lib/m365";
import { assertPmtLinks } from "@/lib/pmt";
import { getDelegatedGraphClient } from "@/lib/teams-auth";

interface ResolvedAssignee {
  id: string;
  displayName: string;
  userPrincipalName: string;
}

interface GraphMe {
  id?: string;
  displayName?: string;
  userPrincipalName?: string;
}

interface GraphChat {
  id?: string;
}

interface AssigneeBatch {
  assignee: ResolvedAssignee;
  items: MeetingActionItem[];
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function graphFailure(error: unknown): { statusCode: number; message: string } {
  if (typeof error === "object" && error !== null) {
    const record = error as { statusCode?: unknown; message?: unknown; body?: unknown };
    const statusCode = typeof record.statusCode === "number" ? record.statusCode : 500;
    const message = typeof record.message === "string" ? record.message : "";
    const body = typeof record.body === "string" ? record.body : "";
    return { statusCode, message: `${message} ${body}`.trim() };
  }
  if (error instanceof Error) {
    return { statusCode: 500, message: error.message };
  }
  return { statusCode: 500, message: "" };
}

function toTeamsGraphError(error: unknown, fallback: string): GraphError {
  const failure = graphFailure(error);
  if (failure.statusCode === 403) {
    return new GraphError(
      "Microsoft refused this Teams chat. Confirm Chat.Create and ChatMessage.Send are delegated permissions with admin consent, then sign in again.",
      403,
      "TEAMS_FORBIDDEN",
    );
  }
  if (failure.statusCode === 404) {
    return new GraphError(
      `${fallback} That person was not found in Teams.`,
      404,
      "TEAMS_NOT_FOUND",
    );
  }
  if (failure.statusCode === 429) {
    return new GraphError(
      "Microsoft Graph rate limit exceeded. Wait a moment, then send again.",
      429,
      "RATE_LIMIT",
    );
  }
  const detail = failure.message.replace(/\s+/g, " ").trim();
  return new GraphError(
    detail ? `${fallback} ${detail.slice(0, 280)}` : fallback,
    failure.statusCode >= 400 && failure.statusCode < 600 ? failure.statusCode : 502,
    "TEAMS_SEND_FAILED",
  );
}

function resolveAssignee(
  item: MeetingActionItem,
  users: TenantUser[],
): ResolvedAssignee | null {
  const matched = matchAssignee(item.pointPerson, users);
  if (matched) {
    const user = users.find(
      (candidate) => candidate.userPrincipalName === matched.userPrincipalName,
    );
    if (user) {
      return user;
    }
  }

  const upn = item.assigneeUpn?.trim();
  if (!upn) {
    return null;
  }
  return (
    users.find(
      (candidate) => candidate.userPrincipalName.toLowerCase() === upn.toLowerCase(),
    ) ?? null
  );
}

const IRIS_EMAIL = "iris.ebale@sindbad.tech";

const ASSIGNMENT_NOTE =
  "If you think the ticket does not belong to you, or there is any discrepancy in the task, please inform Emmanuel or Iris so they can modify or change it. Thank you.";

const REFERENCE_NOTE =
  "If you need images or photos for reference, check the Teams meeting chat. Any that were shared are there.";

function sortTickets(items: MeetingActionItem[]): MeetingActionItem[] {
  return [...items].sort((left, right) =>
    ticketTitle(left).localeCompare(ticketTitle(right), undefined, {
      sensitivity: "base",
    }),
  );
}

function assignmentIntro(count: number, assigneeName: string): string {
  const name = assigneeName.trim();
  const greeting = name ? `${name}, ` : "";
  if (count === 1) {
    return `${greeting}a ticket was assigned to you from the meeting review.`;
  }
  return `${greeting}${count} tickets were assigned to you from the meeting review.`;
}

function pmtLinkLabel(item: MeetingActionItem): string {
  return item.pmtNumber ? `Open PMT #${item.pmtNumber}` : "Open PMT ticket";
}

function displayOrNone(value: string): string {
  const trimmed = value.trim();
  return trimmed || "None";
}

interface CardTextBlock {
  type: "TextBlock";
  text: string;
  wrap: true;
  weight?: "Bolder";
  size?: "Small" | "Medium" | "Large";
  color?: "Accent" | "Attention" | "Warning" | "Good" | "Default";
  isSubtle?: boolean;
  spacing?: "None" | "Small" | "Medium";
  separator?: boolean;
}

interface CardFactSet {
  type: "FactSet";
  facts: Array<{ title: string; value: string }>;
  spacing?: "Small" | "Medium";
}

interface CardActionSet {
  type: "ActionSet";
  spacing?: "Small" | "Medium";
  actions: Array<{ type: "Action.OpenUrl"; title: string; url: string }>;
}

interface CardContainer {
  type: "Container";
  style: "emphasis";
  spacing: "Medium";
  separator?: boolean;
  items: Array<CardTextBlock | CardFactSet | CardActionSet>;
}

function ticketTitle(item: MeetingActionItem): string {
  return item.title.trim() || item.task.trim() || "Untitled ticket";
}

function clipText(value: string, max: number): string {
  const trimmed = value.trim();
  if (trimmed.length <= max) {
    return trimmed;
  }
  return `${trimmed.slice(0, max - 1).trimEnd()}…`;
}

function priorityColor(
  priority: string,
): "Attention" | "Warning" | "Good" | "Default" {
  const value = priority.trim().toLowerCase();
  if (value === "high" || value === "critical" || value === "urgent") {
    return "Attention";
  }
  if (value === "medium") {
    return "Warning";
  }
  if (value === "low") {
    return "Good";
  }
  return "Default";
}

function ticketFacts(item: MeetingActionItem): Array<{ title: string; value: string }> {
  const facts: Array<{ title: string; value: string }> = [
    { title: "Type", value: item.type.trim() || "Task" },
    { title: "Priority", value: item.priority.trim() || "Unspecified" },
    { title: "Due", value: item.deadline.trim() || "TBD" },
    { title: "Status", value: item.status.trim() || "New" },
  ];
  if (item.meetingTitle.trim()) {
    facts.push({ title: "Meeting", value: clipText(item.meetingTitle, 180) });
  }
  if (item.meetingDate.trim()) {
    facts.push({ title: "When", value: item.meetingDate.trim() });
  }
  facts.push({ title: "Dev tag", value: normalizeDevTag(item.devTag) });
  facts.push({ title: "Related to", value: clipText(displayOrNone(item.relatedTo), 180) });
  return facts;
}

function buildAdaptiveCard(items: MeetingActionItem[], assigneeName: string): string {
  const ordered = sortTickets(items);
  const intro: CardTextBlock = {
    type: "TextBlock",
    text: assignmentIntro(ordered.length, assigneeName),
    wrap: true,
    weight: "Bolder",
    size: "Small",
    color: "Accent",
  };
  const tickets: CardContainer[] = ordered.map((item, index) => {
    const blocks: Array<CardTextBlock | CardFactSet | CardActionSet> = [
      {
        type: "TextBlock",
        text: clipText(ticketTitle(item), 200),
        wrap: true,
        weight: "Bolder",
        size: "Medium",
      },
      {
        type: "TextBlock",
        text: `${item.priority.trim() || "Unspecified"} priority`,
        wrap: true,
        weight: "Bolder",
        size: "Small",
        color: priorityColor(item.priority),
        spacing: "Small",
      },
      {
        type: "FactSet",
        facts: ticketFacts(item).filter((fact) => fact.title !== "Priority"),
        spacing: "Small",
      },
    ];
    const description = clipText(retainCrucialSourceLinks(item.description), 1200);
    if (description) {
      blocks.push({
        type: "TextBlock",
        text: description,
        wrap: true,
        spacing: "Medium",
      });
    }
    if (item.pmtUrl?.trim()) {
      blocks.push({
        type: "ActionSet",
        spacing: "Small",
        actions: [
          {
            type: "Action.OpenUrl",
            title: pmtLinkLabel(item),
            url: item.pmtUrl,
          },
        ],
      });
    }
    return {
      type: "Container",
      style: "emphasis",
      spacing: "Medium",
      separator: index > 0,
      items: blocks,
    };
  });

  const reference: CardTextBlock = {
    type: "TextBlock",
    text: REFERENCE_NOTE,
    wrap: true,
    size: "Small",
    isSubtle: true,
    spacing: "Medium",
    separator: true,
  };
  const note: CardTextBlock = {
    type: "TextBlock",
    text: ASSIGNMENT_NOTE,
    wrap: true,
    size: "Small",
    isSubtle: true,
    spacing: "Small",
  };

  return JSON.stringify({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [intro, ...tickets, reference, note],
  });
}

function buildChatHtml(items: MeetingActionItem[], assigneeName: string): string {
  const ordered = sortTickets(items);
  const intro = `<p><strong>${escapeHtml(assignmentIntro(ordered.length, assigneeName))}</strong></p>`;

  const blocks = ordered.map((item) => {
    const rows = ticketFacts(item)
      .map(
        (fact) =>
          `<li><strong>${escapeHtml(fact.title)}</strong> — ${escapeHtml(fact.value)}</li>`,
      )
      .join("");
    const description = clipText(retainCrucialSourceLinks(item.description), 2000);
    const detail = description ? `<p>${renderMarkdownLinksAsHtml(description)}</p>` : "";
    const title = escapeHtml(ticketTitle(item));
    const heading = item.pmtUrl?.trim()
      ? `<h2><a href="${escapeHtml(item.pmtUrl)}">${title}</a></h2><p><a href="${escapeHtml(item.pmtUrl)}">${escapeHtml(pmtLinkLabel(item))}</a></p>`
      : `<h2>${title}</h2>`;
    return `${heading}<ul>${rows}</ul>${detail}`;
  });

  const note = `<blockquote><p><em>${escapeHtml(REFERENCE_NOTE)}</em></p><p><em>${escapeHtml(ASSIGNMENT_NOTE)}</em></p></blockquote>`;
  return `${intro}${blocks.join("")}${note}`;
}

interface TeamsMention {
  id: number;
  mentionText: string;
  mentioned: Record<string, unknown>;
}

interface TeamsMessagePayload {
  body: {
    contentType: "html";
    content: string;
  };
  attachments?: Array<{
    id: string;
    contentType: "application/vnd.microsoft.card.adaptive";
    content: string;
  }>;
  mentions?: TeamsMention[];
}

function buildCardPayload(
  items: MeetingActionItem[],
  assigneeName: string,
): TeamsMessagePayload {
  const attachmentId = randomUUID();
  return {
    body: {
      contentType: "html",
      content: `<attachment id="${attachmentId}"></attachment>`,
    },
    attachments: [
      {
        id: attachmentId,
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildAdaptiveCard(items, assigneeName),
      },
    ],
  };
}

interface DirectoryUserRecord {
  id: string;
  displayName?: string | null;
  userPrincipalName: string;
}

function isDirectoryUser(value: unknown): value is DirectoryUserRecord {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const record = value as {
    id?: unknown;
    displayName?: unknown;
    userPrincipalName?: unknown;
  };
  return (
    typeof record.id === "string" &&
    record.id.length > 0 &&
    typeof record.userPrincipalName === "string" &&
    record.userPrincipalName.length > 0 &&
    (record.displayName === undefined ||
      record.displayName === null ||
      typeof record.displayName === "string")
  );
}

function toDirectoryUser(value: DirectoryUserRecord): ResolvedAssignee {
  const name = typeof value.displayName === "string" ? value.displayName.trim() : "";
  return {
    id: value.id,
    displayName: name || "Iris",
    userPrincipalName: value.userPrincipalName,
  };
}

async function lookupIrisInDirectory(): Promise<ResolvedAssignee | null> {
  const client = getGraphClient();
  try {
    const user: unknown = await client
      .api(`/users/${encodeURIComponent(IRIS_EMAIL)}`)
      .select("id,displayName,userPrincipalName")
      .get();
    if (isDirectoryUser(user)) {
      return toDirectoryUser(user);
    }
  } catch {
    // The address may be a mail alias rather than the user principal name.
  }

  try {
    const response: unknown = await client
      .api("/users")
      .filter(`mail eq '${IRIS_EMAIL}'`)
      .select("id,displayName,userPrincipalName")
      .top(5)
      .get();
    const value =
      typeof response === "object" && response !== null && "value" in response
        ? response.value
        : null;
    if (!Array.isArray(value)) {
      return null;
    }
    const match = value.find(isDirectoryUser);
    return match ? toDirectoryUser(match) : null;
  } catch {
    return null;
  }
}

async function resolveIris(users: TenantUser[]): Promise<ResolvedAssignee> {
  const fromList = users.find(
    (user) => user.userPrincipalName.toLowerCase() === IRIS_EMAIL,
  );
  if (fromList) {
    return fromList;
  }

  const fromDirectory = await lookupIrisInDirectory();
  if (fromDirectory) {
    return fromDirectory;
  }

  throw new GraphError(
    `Could not find Iris (${IRIS_EMAIL}) in the directory, so she was not added to the chat.`,
    404,
    "TEAMS_NOT_FOUND",
  );
}

function chatMemberIds(
  senderId: string,
  assigneeId: string,
  irisId: string | null,
): string[] {
  const ids = [senderId, assigneeId];
  if (irisId) {
    ids.push(irisId);
  }
  return [...new Set(ids)];
}

const GROUP_CHAT_STORE_PATH = path.join(process.cwd(), "data", "teams-group-chats.json");
const rememberedGroupChats = new Map<string, string>();

interface GroupChatStore {
  chats: Record<string, string>;
}

interface ListedGroupChat {
  id: string;
  members: string[];
  topic: string;
}

function memberKey(memberIds: string[]): string {
  return [...new Set(memberIds)].sort().join("|");
}

function isExactMemberSet(actualIds: string[], expectedIds: string[]): boolean {
  if (actualIds.length !== expectedIds.length) {
    return false;
  }
  const expected = new Set(expectedIds);
  return actualIds.every((id) => expected.has(id));
}

function chatUserIds(members: unknown): string[] {
  if (!Array.isArray(members)) {
    return [];
  }
  const ids: string[] = [];
  for (const member of members) {
    if (typeof member !== "object" || member === null) {
      continue;
    }
    const userId = (member as { userId?: unknown }).userId;
    if (typeof userId === "string" && userId.length > 0) {
      ids.push(userId);
    }
  }
  return ids;
}

function readPageValues(value: unknown): unknown[] {
  if (typeof value !== "object" || value === null || !("value" in value)) {
    return [];
  }
  const list = (value as { value?: unknown }).value;
  return Array.isArray(list) ? list : [];
}

function readNextLink(value: unknown): string {
  if (typeof value !== "object" || value === null || !("@odata.nextLink" in value)) {
    return "";
  }
  const link = (value as { "@odata.nextLink"?: unknown })["@odata.nextLink"];
  return typeof link === "string" ? link : "";
}

function toGraphApiPath(nextLink: string): string {
  try {
    const url = nextLink.startsWith("http")
      ? new URL(nextLink)
      : new URL(nextLink, "https://graph.microsoft.com");
    const combined = `${url.pathname}${url.search}`.replace(/^\/(v1\.0|beta)(?=\/|$)/i, "");
    return combined.startsWith("/") ? combined : `/${combined}`;
  } catch {
    return "";
  }
}

function parseListedChat(value: unknown): ListedGroupChat | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const record = value as { id?: unknown; chatType?: unknown; members?: unknown; topic?: unknown };
  if (typeof record.id !== "string" || record.id.length === 0) {
    return null;
  }
  if (record.chatType !== undefined && record.chatType !== "group") {
    return null;
  }
  return {
    id: record.id,
    members: chatUserIds(record.members),
    topic: typeof record.topic === "string" ? record.topic : "",
  };
}

function isGroupChatStore(value: unknown): value is GroupChatStore {
  if (typeof value !== "object" || value === null || !("chats" in value)) {
    return false;
  }
  const chats = (value as { chats?: unknown }).chats;
  if (typeof chats !== "object" || chats === null || Array.isArray(chats)) {
    return false;
  }
  return Object.values(chats).every((entry) => typeof entry === "string" && entry.length > 0);
}

async function readGroupChatStore(): Promise<GroupChatStore> {
  try {
    const parsed: unknown = JSON.parse(await readFile(GROUP_CHAT_STORE_PATH, "utf8"));
    return isGroupChatStore(parsed) ? parsed : { chats: {} };
  } catch {
    return { chats: {} };
  }
}

async function writeGroupChatStore(store: GroupChatStore): Promise<void> {
  await mkdir(path.dirname(GROUP_CHAT_STORE_PATH), { recursive: true });
  await writeFile(GROUP_CHAT_STORE_PATH, JSON.stringify(store), "utf8");
}

export async function clearRememberedGroupChats(): Promise<void> {
  rememberedGroupChats.clear();
  await rm(GROUP_CHAT_STORE_PATH, { force: true });
}

async function readRememberedChatId(key: string): Promise<string | null> {
  const cached = rememberedGroupChats.get(key);
  if (cached) {
    return cached;
  }
  const store = await readGroupChatStore();
  const chatId = store.chats[key];
  if (!chatId) {
    return null;
  }
  rememberedGroupChats.set(key, chatId);
  return chatId;
}

async function rememberGroupChat(key: string, chatId: string): Promise<void> {
  rememberedGroupChats.set(key, chatId);
  const store = await readGroupChatStore();
  if (store.chats[key] === chatId) {
    return;
  }
  store.chats[key] = chatId;
  await writeGroupChatStore(store);
}

async function forgetGroupChat(key: string): Promise<void> {
  rememberedGroupChats.delete(key);
  const store = await readGroupChatStore();
  if (!store.chats[key]) {
    return;
  }
  delete store.chats[key];
  await writeGroupChatStore(store);
}

async function readChatMemberIds(client: Client, chatId: string): Promise<string[] | null> {
  try {
    const collected: string[] = [];
    let nextPath = `/chats/${chatId}/members`;
    for (let page = 0; page < 5 && nextPath; page += 1) {
      const response: unknown = await client.api(nextPath).get();
      collected.push(...chatUserIds(readPageValues(response)));
      const link = readNextLink(response);
      nextPath = link ? toGraphApiPath(link) : "";
    }
    return collected;
  } catch (error) {
    const failure = graphFailure(error);
    if (failure.statusCode === 429) {
      throw error;
    }
    if (failure.statusCode === 404) {
      return [];
    }
    return null;
  }
}

async function fetchGroupChatPage(
  client: Client,
  requestPath: string,
  firstPage: boolean,
  expandMembers: boolean,
): Promise<unknown> {
  if (!firstPage) {
    return client.api(requestPath).get();
  }
  const request = client
    .api(requestPath)
    .filter("chatType eq 'group'")
    .select("id,topic,chatType,lastUpdatedDateTime")
    .top(50);
  return (expandMembers ? request.expand("members") : request).get();
}

async function scanGroupChats(
  client: Client,
  requestPath: string,
  expectedIds: string[],
): Promise<string | null> {
  let nextPath = requestPath;
  let firstPage = true;
  let expandMembers = true;
  let scanned = 0;

  while (nextPath && scanned < 400) {
    let response: unknown;
    try {
      response = await fetchGroupChatPage(client, nextPath, firstPage, expandMembers);
    } catch (error) {
      if (expandMembers && firstPage && graphFailure(error).statusCode === 400) {
        expandMembers = false;
        continue;
      }
      throw error;
    }

    firstPage = false;
    for (const entry of readPageValues(response)) {
      scanned += 1;
      const chat = parseListedChat(entry);
      if (!chat) {
        continue;
      }
      const memberIds =
        chat.members.length > 0 ? chat.members : await readChatMemberIds(client, chat.id);
      if (memberIds && isExactMemberSet(memberIds, expectedIds)) {
        return chat.id;
      }
      if (scanned >= 400) {
        break;
      }
    }

    const link = readNextLink(response);
    nextPath = link ? toGraphApiPath(link) : "";
  }

  return null;
}

async function findReusableGroupChat(
  delegated: Client,
  memberIds: string[],
): Promise<string | null> {
  const key = memberKey(memberIds);
  const remembered = await readRememberedChatId(key);
  if (remembered) {
    let members = await readChatMemberIds(getGraphClient(), remembered);
    if (members === null) {
      members = await readChatMemberIds(delegated, remembered);
    }
    if (members === null || isExactMemberSet(members, memberIds)) {
      return remembered;
    }
    await forgetGroupChat(key);
  }

  const senderId = memberIds[0];
  if (!senderId) {
    return null;
  }

  const attempts: Array<{ client: Client; path: string }> = [
    { client: getGraphClient(), path: `/users/${senderId}/chats` },
    { client: delegated, path: "/me/chats" },
  ];

  let searched = false;
  for (const attempt of attempts) {
    if (searched) {
      break;
    }
    try {
      const found = await scanGroupChats(attempt.client, attempt.path, memberIds);
      searched = true;
      if (found) {
        await rememberGroupChat(key, found);
        return found;
      }
    } catch (error) {
      const failure = graphFailure(error);
      if (failure.statusCode === 429) {
        throw toTeamsGraphError(error, "Could not look up the existing Teams group.");
      }
    }
  }

  return null;
}

async function openChat(
  client: Client,
  memberIds: string[],
  topic: string | null,
): Promise<string> {
  const uniqueIds = [...new Set(memberIds)];
  if (uniqueIds.length < 2) {
    throw new GraphError(
      "Teams needs another person in this chat.",
      400,
      "TEAMS_SEND_FAILED",
    );
  }

  const chatType = uniqueIds.length === 2 ? "oneOnOne" : "group";
  if (chatType === "group") {
    const existingId = await findReusableGroupChat(client, uniqueIds);
    if (existingId) {
      return existingId;
    }
  }

  try {
    const chat = (await client.api("/chats").post({
      chatType,
      ...(chatType === "group" && topic ? { topic } : {}),
      members: uniqueIds.map((id) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${id}')`,
      })),
    })) as GraphChat;
    if (!chat.id) {
      throw new GraphError(
        "Microsoft Graph created the chat without an id.",
        502,
        "TEAMS_SEND_FAILED",
      );
    }
    if (chatType === "group") {
      await rememberGroupChat(memberKey(uniqueIds), chat.id);
    }
    return chat.id;
  } catch (error) {
    if (error instanceof GraphError) {
      throw error;
    }
    throw toTeamsGraphError(
      error,
      chatType === "group"
        ? "Could not open a Teams group chat."
        : "Could not open a Teams chat with this person.",
    );
  }
}

export async function openDirectChatWithEmail(
  email: string,
): Promise<{ chatId: string; displayName: string }> {
  const trimmed = email.trim().toLowerCase();
  if (!trimmed.includes("@")) {
    throw new AppError("Choose a person to send the aging tickets to.", 400, "PMT_AGING");
  }

  const { client } = await getDelegatedGraphClient();
  let me: GraphMe;
  try {
    me = (await client.api("/me").select("id,displayName,userPrincipalName").get()) as GraphMe;
  } catch (error) {
    throw toTeamsGraphError(error, "Could not read the signed-in Teams account.");
  }
  if (!me.id) {
    throw new GraphError(
      "Microsoft Graph did not return the signed-in user.",
      502,
      "TEAMS_SEND_FAILED",
    );
  }

  const users = await listTenantUsers();
  const person = users.find(
    (user) =>
      user.userPrincipalName.toLowerCase() === trimmed || user.mail?.toLowerCase() === trimmed,
  );
  if (!person) {
    throw new AppError(
      "That person has no Microsoft 365 account the analyzer can message.",
      404,
      "TEAMS_USER",
    );
  }
  if (
    person.id === me.id ||
    person.userPrincipalName.toLowerCase() === (me.userPrincipalName ?? "").toLowerCase()
  ) {
    throw new AppError("Teams cannot open a one-to-one chat with yourself.", 400, "TEAMS_SELF");
  }

  const chatId = await openChat(client, [me.id, person.id], null);
  return { chatId, displayName: person.displayName };
}

async function postMessagePayload(
  client: Client,
  chatId: string,
  payload: TeamsMessagePayload,
): Promise<void> {
  await client.api(`/chats/${chatId}/messages`).post(payload);
}

async function postChatMessage(
  client: Client,
  chatId: string,
  items: MeetingActionItem[],
  assigneeName: string,
): Promise<void> {
  try {
    await postMessagePayload(client, chatId, buildCardPayload(items, assigneeName));
  } catch (error) {
    const failure = graphFailure(error);
    if (failure.statusCode !== 400) {
      throw toTeamsGraphError(error, "Could not post the Teams message.");
    }
    try {
      await postMessagePayload(client, chatId, {
        body: { contentType: "html", content: buildChatHtml(items, assigneeName) },
      });
    } catch (htmlError) {
      throw toTeamsGraphError(htmlError, "Could not post the Teams message.");
    }
  }
}

export async function sendActionItemsToAssignees(
  actionItems: MeetingActionItem[],
  options: { includeIris?: boolean } = {},
): Promise<TeamsSendResult> {
  if (actionItems.length === 0) {
    throw new AppError(
      "Select at least one action item to send.",
      400,
      "NO_ACTION_ITEMS",
    );
  }
  assertPmtLinks(actionItems);

  const { client, displayName } = await getDelegatedGraphClient();
  let me: GraphMe;
  try {
    me = (await client.api("/me").select("id,displayName,userPrincipalName").get()) as GraphMe;
  } catch (error) {
    throw toTeamsGraphError(error, "Could not read the signed-in Teams account.");
  }
  if (!me.id) {
    throw new GraphError(
      "Microsoft Graph did not return the signed-in user.",
      502,
      "TEAMS_SEND_FAILED",
    );
  }

  let users: TenantUser[] = [];
  try {
    users = await listTenantUsers();
  } catch {
    users = [];
  }

  const batches = new Map<string, AssigneeBatch>();
  const unmatched: string[] = [];
  const failed: TeamsSendResult["failed"] = [];
  const selfBlocked = new Set<string>();

  for (const item of actionItems) {
    const assignee = resolveAssignee(item, users);
    if (!assignee) {
      unmatched.push(item.pointPerson.trim() || "Unassigned");
      continue;
    }
    if (
      assignee.id === me.id ||
      assignee.userPrincipalName.toLowerCase() ===
        (me.userPrincipalName ?? "").toLowerCase()
    ) {
      selfBlocked.add(assignee.displayName);
      continue;
    }
    const existing = batches.get(assignee.id);
    if (existing) {
      existing.items.push(item);
    } else {
      batches.set(assignee.id, { assignee, items: [item] });
    }
  }

  for (const name of selfBlocked) {
    failed.push({
      assignee: name,
      error: "Teams cannot open a one-to-one chat with yourself.",
    });
  }

  const recipients: string[] = [];
  const filedKeys: string[] = [];

  const orderedBatches = [...batches.values()].sort((left, right) =>
    left.assignee.displayName.localeCompare(right.assignee.displayName, undefined, {
      sensitivity: "base",
    }),
  );

  const iris =
    options.includeIris && orderedBatches.length > 0
      ? await resolveIris(users)
      : null;

  for (const batch of orderedBatches) {
    try {
      const memberIds = chatMemberIds(me.id, batch.assignee.id, iris?.id ?? null);
      const topic =
        memberIds.length > 2 ? `Tickets for ${batch.assignee.displayName}` : null;
      const chatId = await openChat(client, memberIds, topic);
      await postChatMessage(client, chatId, batch.items, batch.assignee.displayName);
      recipients.push(batch.assignee.displayName);
      for (const item of batch.items) {
        if (item.filingKey) {
          filedKeys.push(item.filingKey);
        }
      }
    } catch (error) {
      failed.push({
        assignee: batch.assignee.displayName,
        error: error instanceof Error ? error.message : "Could not send the Teams message.",
      });
    }
  }

  return {
    sent: recipients.length,
    recipients,
    unmatched: [...new Set(unmatched)],
    failed,
    filedKeys,
    sender: me.displayName?.trim() || displayName,
    includedIris: iris !== null && recipients.length > 0,
  };
}

function topicStoreKey(topic: string): string {
  return `topic:${displayBoardName(topic).toLowerCase()}`;
}

function ticketLinkLabel(ticket: CodeQcSendTicket): string {
  const title = ticket.title?.replace(/\s+/g, " ").trim() ?? "";
  const label = title ? `#${ticket.number} ${title}` : `#${ticket.number}`;
  return label.replace(/[\[\]]/g, "");
}

function ticketLinkMarkdown(ticket: CodeQcSendTicket, index: number): string {
  return `${index}. [${ticketLinkLabel(ticket)}](${ticket.url})`;
}

function ticketLinkHtml(ticket: CodeQcSendTicket, index: number): string {
  return `${index}. <a href="${escapeHtml(ticket.url)}">${escapeHtml(ticketLinkLabel(ticket))}</a>`;
}

function findDirectoryUser(name: string, users: TenantUser[]): TenantUser | null {
  const exact = users.find((user) => user.displayName.toLowerCase() === name.toLowerCase());
  if (exact) {
    return exact;
  }
  const matched = matchAssignee(name, users);
  if (!matched) {
    return null;
  }
  return (
    users.find(
      (user) => user.userPrincipalName.toLowerCase() === matched.userPrincipalName.toLowerCase(),
    ) ?? null
  );
}

function noteHtmlWithMentions(
  note: string,
  chatId: string,
  users: TenantUser[],
  startId = 0,
): { html: string; mentions: TeamsMention[] } {
  const mentions: TeamsMention[] = [];
  const pattern = /@Everyone|@([^@\n]+?)(?=\s@|\s—|\n|$)/g;
  let html = "";
  let cursor = 0;
  for (const match of note.matchAll(pattern)) {
    const index = match.index ?? 0;
    html += renderMarkdownLinksAsHtml(note.slice(cursor, index));
    if (match[0] === "@Everyone") {
      const id = startId + mentions.length;
      mentions.push({
        id,
        mentionText: "Everyone",
        mentioned: {
          conversation: {
            id: chatId,
            displayName: "Everyone",
            conversationIdentityType: "chat",
          },
        },
      });
      html += `<at id="${id}">Everyone</at>`;
    } else {
      const name = match[1]?.trim() ?? "";
      const person = name ? findDirectoryUser(name, users) : null;
      if (!person) {
        html += renderMarkdownLinksAsHtml(match[0]);
      } else {
        const id = startId + mentions.length;
        mentions.push({
          id,
          mentionText: person.displayName,
          mentioned: {
            user: {
              id: person.id,
              displayName: person.displayName,
              userIdentityType: "aadUser",
            },
          },
        });
        html += `<at id="${id}">${escapeHtml(person.displayName)}</at>`;
      }
    }
    cursor = index + match[0].length;
  }
  html += renderMarkdownLinksAsHtml(note.slice(cursor));
  return { html, mentions };
}

function codeQcMoreActions(
  hidden: CodeQcSendTicket[],
): Array<{ type: "Action.ShowCard"; title: string; card: { type: "AdaptiveCard"; body: Array<Record<string, unknown>> } }> {
  const actions: Array<{
    type: "Action.ShowCard";
    title: string;
    card: { type: "AdaptiveCard"; body: Array<Record<string, unknown>> };
  }> = [];
  for (let index = 0; index < hidden.length; index += CODE_QC_MORE_BATCH) {
    const batch = hidden.slice(index, index + CODE_QC_MORE_BATCH);
    const from = VISIBLE_CODE_QC_TICKETS + index + 1;
    const to = VISIBLE_CODE_QC_TICKETS + index + batch.length;
    actions.push({
      type: "Action.ShowCard",
      title: hidden.length <= CODE_QC_MORE_BATCH ? "View more tickets" : `More tickets (${from}–${to})`,
      card: {
        type: "AdaptiveCard",
        body: [
          {
            type: "TextBlock",
            text: batch.map((ticket, offset) => ticketLinkMarkdown(ticket, from + offset)).join("\n"),
            wrap: true,
          },
        ],
      },
    });
  }
  return actions;
}

function buildCodeQcCard(note: string, tickets: CodeQcSendTicket[], topic: string): string {
  const shown = tickets.slice(0, VISIBLE_CODE_QC_TICKETS);
  const hidden = tickets.slice(VISIBLE_CODE_QC_TICKETS);
  const body: Array<Record<string, unknown>> = [
    {
      type: "TextBlock",
      text: topic,
      wrap: true,
      weight: "Bolder",
      size: "Medium",
    },
    {
      type: "TextBlock",
      text: note.trim() || "Please fix these PMT tickets.",
      wrap: true,
      spacing: "Small",
    },
  ];
  if (shown.length > 0) {
    body.push({
      type: "TextBlock",
      text: shown.map((ticket, index) => ticketLinkMarkdown(ticket, index + 1)).join("\n"),
      wrap: true,
      spacing: "Medium",
    });
  }
  if (hidden.length > 0) {
    body.push({
      type: "TextBlock",
      text: `${hidden.length} more ${hidden.length === 1 ? "ticket is" : "tickets are"} hidden. Click More to open the rest.`,
      wrap: true,
      size: "Small",
      isSubtle: true,
      spacing: "Small",
    });
  }
  return JSON.stringify({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body,
    ...(hidden.length > 0 ? { actions: codeQcMoreActions(hidden) } : {}),
  });
}

function buildCodeQcHtml(note: string, tickets: CodeQcSendTicket[], topic: string): string {
  const shown = tickets.slice(0, VISIBLE_CODE_QC_TICKETS);
  const hidden = tickets.slice(VISIBLE_CODE_QC_TICKETS);
  const more =
    hidden.length > 0
      ? `<p>${hidden.length} more ${hidden.length === 1 ? "ticket is" : "tickets are"} hidden. Click More in Teams, or open PMT to see the rest.</p><p>${hidden
          .slice(0, CODE_QC_MORE_BATCH)
          .map((ticket, index) => ticketLinkHtml(ticket, VISIBLE_CODE_QC_TICKETS + index + 1))
          .join("<br/>")}</p>`
      : "";
  const list =
    shown.length > 0
      ? `<p>${shown.map((ticket, index) => ticketLinkHtml(ticket, index + 1)).join("<br/>")}</p>`
      : "";
  return [
    `<p><strong>${escapeHtml(topic)}</strong></p>`,
    `<p>${renderMarkdownLinksAsHtml(note.trim() || "Please fix these PMT tickets.")}</p>`,
    list,
    more,
  ].join("");
}

async function scanGroupChatsByTopic(
  client: Client,
  requestPath: string,
  topic: string,
): Promise<string | null> {
  const needle = displayBoardName(topic).toLowerCase();
  let nextPath = requestPath;
  let firstPage = true;
  let scanned = 0;
  while (nextPath && scanned < 400) {
    const response: unknown = await fetchGroupChatPage(client, nextPath, firstPage, false);
    firstPage = false;
    for (const entry of readPageValues(response)) {
      scanned += 1;
      const chat = parseListedChat(entry);
      if (chat && displayBoardName(chat.topic).toLowerCase() === needle) {
        return chat.id;
      }
      if (scanned >= 400) {
        break;
      }
    }
    const link = readNextLink(response);
    nextPath = link ? toGraphApiPath(link) : "";
  }
  return null;
}

async function findGroupChatByTopic(
  delegated: Client,
  senderId: string,
  topic: string,
): Promise<string | null> {
  const key = topicStoreKey(topic);
  const remembered = await readRememberedChatId(key);
  if (remembered) {
    let members = await readChatMemberIds(getGraphClient(), remembered);
    if (members === null) {
      members = await readChatMemberIds(delegated, remembered);
    }
    if (members === null || members.length > 0) {
      return remembered;
    }
    await forgetGroupChat(key);
  }

  const attempts: Array<{ client: Client; path: string }> = [
    { client: getGraphClient(), path: `/users/${senderId}/chats` },
    { client: delegated, path: "/me/chats" },
  ];
  for (const attempt of attempts) {
    try {
      const found = await scanGroupChatsByTopic(attempt.client, attempt.path, topic);
      if (found) {
        await rememberGroupChat(key, found);
        return found;
      }
    } catch (error) {
      if (graphFailure(error).statusCode === 429) {
        throw toTeamsGraphError(error, "Could not look up the existing Teams group.");
      }
    }
  }
  return null;
}

async function addChatMember(client: Client, chatId: string, userId: string): Promise<void> {
  await client.api(`/chats/${chatId}/members`).post({
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: ["owner"],
    "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
  });
}

async function openGroupChatWithTopic(
  client: Client,
  memberIds: string[],
  topic: string,
): Promise<string> {
  const uniqueIds = [...new Set(memberIds)];
  if (uniqueIds.length < 3) {
    throw new AppError(
      "Tick at least two licensed users for the Teams group. You are added automatically, and Teams needs three people for a named group.",
      400,
      "TEAMS_GROUP",
    );
  }

  const senderId = uniqueIds[0];
  const existing = senderId ? await findGroupChatByTopic(client, senderId, topic) : null;
  if (existing) {
    const current = (await readChatMemberIds(client, existing)) ?? [];
    const have = new Set(current);
    for (const id of uniqueIds) {
      if (have.has(id)) {
        continue;
      }
      try {
        await addChatMember(client, existing, id);
      } catch {
        // The person may already be in the group, or Graph refused the add.
      }
    }
    return existing;
  }

  try {
    const chat = (await client.api("/chats").post({
      chatType: "group",
      topic,
      members: uniqueIds.map((id) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${id}')`,
      })),
    })) as GraphChat;
    if (!chat.id) {
      throw new GraphError(
        "Microsoft Graph created the chat without an id.",
        502,
        "TEAMS_SEND_FAILED",
      );
    }
    await rememberGroupChat(topicStoreKey(topic), chat.id);
    return chat.id;
  } catch (error) {
    if (error instanceof GraphError || error instanceof AppError) {
      throw error;
    }
    throw toTeamsGraphError(error, "Could not open a Teams group chat.");
  }
}

export async function sendCodeQcGroupMessage(
  topic: string,
  memberIds: string[],
  note: string,
  tickets: CodeQcSendTicket[],
): Promise<CodeQcTeamsResult> {
  const title = displayBoardName(topic);
  if (!title) {
    throw new AppError("Choose a PMT board so the Teams group has a title.", 400, "PMT_BOARD_NAME");
  }
  const linked = tickets.filter((ticket) => ticket.url.trim() && ticket.number > 0);
  if (linked.length === 0) {
    throw new AppError("PMT did not return ticket links to send.", 502, "PMT_CREATE_FAILED");
  }

  const { client } = await getDelegatedGraphClient();
  let me: GraphMe;
  try {
    me = (await client.api("/me").select("id,displayName,userPrincipalName").get()) as GraphMe;
  } catch (error) {
    throw toTeamsGraphError(error, "Could not read the signed-in Teams account.");
  }
  if (!me.id) {
    throw new GraphError(
      "Microsoft Graph did not return the signed-in user.",
      502,
      "TEAMS_SEND_FAILED",
    );
  }

  const chatId = await openGroupChatWithTopic(client, [me.id, ...memberIds], title);
  const trimmedNote = note.trim().slice(0, 4000);
  const ticketLine = /^\d+\.\s+/;
  const lines = trimmedNote.split(/\r?\n/);
  const firstTicket = lines.findIndex((line) => ticketLine.test(line.trim()));
  const lastTicket =
    firstTicket < 0
      ? -1
      : lines.reduce((last, line, index) => (ticketLine.test(line.trim()) ? index : last), firstTicket);
  const intro = (firstTicket < 0 ? trimmedNote : lines.slice(0, firstTicket).join("\n")).trim();
  const outro = firstTicket < 0 ? "" : lines.slice(lastTicket + 1).join("\n").trim();
  const users = await listTenantUsers();
  const introNoted = noteHtmlWithMentions(intro, chatId, users);
  const outroNoted = noteHtmlWithMentions(outro, chatId, users, introNoted.mentions.length);
  const list = `<ol>${linked
    .map(
      (ticket) =>
        `<li><a href="${escapeHtml(ticket.url)}">${escapeHtml(ticketLinkLabel(ticket))}</a></li>`,
    )
    .join("")}</ol>`;
  const html = [
    `<p><strong>${escapeHtml(title)}</strong></p>`,
    introNoted.html ? `<p>${introNoted.html}</p>` : "",
    list,
    outroNoted.html ? `<p>${outroNoted.html}</p>` : "",
  ]
    .filter(Boolean)
    .join("");
  const mentions = [...introNoted.mentions, ...outroNoted.mentions];
  const hiddenCount = Math.max(0, linked.length - VISIBLE_CODE_QC_TICKETS);
  try {
    await postMessagePayload(client, chatId, {
      body: {
        contentType: "html",
        content: html,
      },
      ...(mentions.length > 0 ? { mentions } : {}),
    });
  } catch (error) {
    const failure = graphFailure(error);
    if (failure.statusCode !== 400) {
      throw toTeamsGraphError(error, "Could not post the Teams message.");
    }
    try {
      await postMessagePayload(client, chatId, {
        body: {
          contentType: "html",
          content: `<attachment id="code-qc"></attachment>`,
        },
        attachments: [
          {
            id: "code-qc",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildCodeQcCard([intro, outro].filter(Boolean).join("\n\n"), linked, title),
          },
        ],
      });
    } catch {
      try {
        await postMessagePayload(client, chatId, {
          body: {
            contentType: "html",
            content: html,
          },
        });
      } catch (htmlError) {
        throw toTeamsGraphError(htmlError, "Could not post the Teams message.");
      }
    }
  }

  return {
    sent: true,
    topic: title,
    previewCount: Math.min(VISIBLE_CODE_QC_TICKETS, linked.length),
    hiddenCount,
  };
}
