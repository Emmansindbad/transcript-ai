import "server-only";

import type {
  MeetingActionItem,
  PmtBoardOption,
  PmtCreateResult,
  PmtTicketResult,
  TenantUser,
  TicketLane,
} from "@/types/meeting";
import { resolveDirectoryAssignee } from "@/lib/assignees";
import { retainCrucialSourceLinks } from "@/lib/chat-links";
import { ticketLaneOf } from "@/lib/ticket-lane";
import { normalizeDevTag } from "@/lib/dev-tags";
import { getPmtApiToken, getPmtBaseUrl } from "@/lib/env";
import { AppError } from "@/lib/errors";
import { listTenantUsers } from "@/lib/m365";

export interface PmtOpenTicket {
  id: number;
  number: number;
  name: string;
  url: string;
  externalKey?: string;
  description?: string;
}

export interface PmtTicketRequest {
  externalKey: string;
  name: string;
  itemType: "task" | "bug";
  priority: "low" | "medium" | "high";
  description?: string;
  assigneeEmail?: string;
  dueAt?: string;
  devTag?: string;
  relatedTo?: string;
}

function clip(value: string, max: number): string {
  const trimmed = value.trim();
  return trimmed.length <= max ? trimmed : trimmed.slice(0, max);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isOpenTicket(value: unknown): value is PmtOpenTicket {
  if (!isRecord(value)) {
    return false;
  }
  return (
    typeof value.id === "number" &&
    typeof value.number === "number" &&
    typeof value.name === "string" &&
    typeof value.url === "string" &&
    (value.externalKey === undefined || value.externalKey === null || typeof value.externalKey === "string") &&
    (value.description === undefined || value.description === null || typeof value.description === "string")
  );
}

function isPmtTicket(value: unknown): value is PmtTicketResult {
  if (!isRecord(value)) {
    return false;
  }
  return (
    typeof value.externalKey === "string" &&
    typeof value.id === "number" &&
    typeof value.number === "number" &&
    typeof value.url === "string" &&
    typeof value.assigneeMatched === "boolean" &&
    typeof value.created === "boolean" &&
    (value.warning === undefined || typeof value.warning === "string") &&
    (value.relatedLabel === undefined || typeof value.relatedLabel === "string") &&
    (value.filled === undefined ||
      (Array.isArray(value.filled) && value.filled.every((entry) => typeof entry === "string"))) &&
    (value.assigneeEmail === undefined || typeof value.assigneeEmail === "string") &&
    (value.assigneeDisplayName === undefined || typeof value.assigneeDisplayName === "string")
  );
}

function dueDate(value: string): string | undefined {
  const trimmed = value.trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(trimmed) ? trimmed : undefined;
}

function assigneeEmail(item: MeetingActionItem, users: TenantUser[]): string | undefined {
  return resolveDirectoryAssignee(item.pointPerson, item.assigneeUpn, users)?.email;
}

function pmtPriority(priority: string): "low" | "medium" | "high" {
  if (priority === "High") {
    return "high";
  }
  if (priority === "Low") {
    return "low";
  }
  return "medium";
}

function relatedLabel(value: string): string {
  const trimmed = value.trim();
  if (!trimmed || /^(none|n\/a|na|tbd|-)$/i.test(trimmed)) {
    return "None";
  }
  return clip(trimmed, 1000);
}

export function assertPmtLinks(actionItems: MeetingActionItem[]): void {
  const missing = actionItems.filter((item) => !item.pmtUrl?.trim());
  if (missing.length === 0) {
    return;
  }
  const titles = missing.map((item) => item.title.trim() || item.task.trim() || "Untitled ticket");
  throw new AppError(
    `Create these on PMT before sending, so the message can link to each ticket: ${titles.join(", ")}.`,
    400,
    "PMT_LINK_REQUIRED",
  );
}

function toPmtItem(
  item: MeetingActionItem,
  users: TenantUser[],
  lane: TicketLane,
): PmtTicketRequest {
  const name = clip(item.title || item.task || "Untitled ticket", 255);
  const request: PmtTicketRequest = {
    externalKey: item.filingKey,
    name,
    itemType: item.type === "Bug" ? "bug" : "task",
    priority: pmtPriority(item.priority),
  };
  const description = clip(retainCrucialSourceLinks(item.description), 10000);
  const email = assigneeEmail(item, users);
  const dueAt = dueDate(item.deadline);
  const devTag = normalizeDevTag(item.devTag);
  const relatedTo = relatedLabel(item.relatedTo);
  if (description) {
    request.description = description;
  }
  if (email) {
    request.assigneeEmail = email;
  }
  if (dueAt) {
    request.dueAt = dueAt;
  }
  if (lane === "development" && devTag) {
    request.devTag = devTag;
  }
  if (lane === "development" || relatedTo !== "None") {
    request.relatedTo = relatedTo;
  }
  return request;
}

function toSheetItem(item: PmtTicketRequest): {
  externalKey: string;
  name: string;
  priority: "low" | "medium" | "high";
  description?: string;
  assigneeEmail?: string;
  dueAt?: string;
  relatedTo?: string;
} {
  return {
    externalKey: item.externalKey,
    name: item.name,
    priority: item.priority,
    ...(item.description ? { description: item.description } : {}),
    ...(item.assigneeEmail ? { assigneeEmail: item.assigneeEmail } : {}),
    ...(item.dueAt ? { dueAt: item.dueAt } : {}),
    ...(item.relatedTo ? { relatedTo: item.relatedTo } : {}),
  };
}

function pmtErrorMessage(status: number, body: unknown): string {
  if (!isRecord(body)) {
    return `PMT rejected the tickets (${status}).`;
  }
  const message = typeof body.message === "string" ? body.message : "";
  const boardName = typeof body.boardName === "string" ? body.boardName : "";
  if (status === 404 && boardName) {
    return `No PMT board is named "${boardName}".`;
  }
  if (status === 409 && Array.isArray(body.boards)) {
    const names = body.boards
      .filter(isRecord)
      .map((board) => (typeof board.name === "string" ? board.name : ""))
      .filter((name) => name.length > 0);
    return names.length > 0
      ? `More than one PMT board is named "${boardName}": ${names.join(", ")}.`
      : message || "More than one PMT board matches that meeting title.";
  }
  if (isRecord(body.errors)) {
    const details = Object.values(body.errors)
      .flatMap((entry) => (Array.isArray(entry) ? entry : []))
      .filter((entry): entry is string => typeof entry === "string");
    if (details.length > 0) {
      return details.join(" ");
    }
  }
  return message || `PMT rejected the tickets (${status}).`;
}

function resultMessage(
  boardName: string,
  items: PmtTicketResult[],
  noun: "ticket" | "row" = "ticket",
): string {
  const created = items.filter((item) => item.created).length;
  const filled = items.filter((item) => !item.created && (item.filled?.length ?? 0) > 0).length;
  const unchanged = items.length - created - filled;
  const parts: string[] = [];
  if (created > 0) {
    const label = created === 1 ? noun : `${noun}s`;
    parts.push(`Created ${created} ${label} on ${boardName}`);
  }
  if (filled > 0) {
    parts.push(
      filled === 1
        ? "Filled missing details on 1 existing ticket"
        : `Filled missing details on ${filled} existing tickets`,
    );
  }
  if (unchanged > 0) {
    parts.push(unchanged === 1 ? "1 was already on PMT" : `${unchanged} were already on PMT`);
  }
  return `${parts.join(". ")}.`;
}

async function postPmtTickets(
  boardName: string,
  meetingDate: string | null,
  items: PmtTicketRequest[],
  access?: { boardId: number; email: string },
): Promise<{ board: { id: number; name: string }; items: PmtTicketResult[] }> {
  let token: string;
  try {
    token = getPmtApiToken();
  } catch {
    throw new AppError(
      "PMT is not configured on this PC. Add PMT_API_TOKEN and restart the analyzer.",
      500,
      "PMT_NOT_CONFIGURED",
    );
  }

  const response = await fetch(`${getPmtBaseUrl()}/api/integrations/tickets`, {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      boardName,
      ...(access ? { boardId: access.boardId, email: access.email } : {}),
      meetingDate,
      items,
    }),
  });

  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    throw new AppError(
      pmtErrorMessage(response.status, body),
      response.status === 404 || response.status === 409 || response.status === 422
        ? response.status
        : 502,
      response.status === 401 ? "PMT_UNAUTHORIZED" : "PMT_CREATE_FAILED",
    );
  }

  const destination = isRecord(body)
    ? isRecord(body.sheet)
      ? body.sheet
      : isRecord(body.board)
        ? body.board
        : null
    : null;
  if (!isRecord(body) || !destination || !Array.isArray(body.items)) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const boardId = destination.id;
  const name = destination.name;
  if (typeof boardId !== "number" || typeof name !== "string") {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const tickets = body.items.filter(isPmtTicket);
  if (tickets.length === 0) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  return { board: { id: boardId, name }, items: tickets };
}

export async function listOpenPmtBoardTickets(
  boardName: string,
): Promise<{ board: { id: number; name: string }; tickets: PmtOpenTicket[] }> {
  const name = boardName.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
  if (!name) {
    throw new AppError("Choose a PMT board.", 400, "PMT_BOARD_NAME");
  }

  let token: string;
  try {
    token = getPmtApiToken();
  } catch {
    throw new AppError(
      "PMT is not configured on this PC. Add PMT_API_TOKEN and restart the analyzer.",
      500,
      "PMT_NOT_CONFIGURED",
    );
  }

  const response = await fetch(
    `${getPmtBaseUrl()}/api/integrations/tickets?boardName=${encodeURIComponent(name)}`,
    {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
      cache: "no-store",
    },
  );
  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    throw new AppError(
      pmtErrorMessage(response.status, body),
      response.status === 404 || response.status === 409 ? response.status : 502,
      response.status === 401 ? "PMT_UNAUTHORIZED" : "PMT_LIST_FAILED",
    );
  }
  if (!isRecord(body) || !isRecord(body.board) || !Array.isArray(body.tickets)) {
    throw new AppError("PMT returned an unexpected ticket list.", 502, "PMT_LIST_FAILED");
  }
  const boardId = body.board.id;
  const listedName = body.board.name;
  if (typeof boardId !== "number" || typeof listedName !== "string") {
    throw new AppError("PMT returned an unexpected ticket list.", 502, "PMT_LIST_FAILED");
  }

  return {
    board: { id: boardId, name: listedName },
    tickets: body.tickets.filter(isOpenTicket),
  };
}

export async function createPmtTicketsOnBoard(
  boardName: string,
  items: PmtTicketRequest[],
): Promise<PmtCreateResult> {
  const name = boardName.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
  if (!name) {
    throw new AppError("Choose a PMT board for these tickets.", 400, "PMT_BOARD_NAME");
  }
  if (items.length === 0) {
    throw new AppError("Select at least one issue to send.", 400, "NO_ACTION_ITEMS");
  }

  const BATCH_SIZE = 20;
  const collected: PmtTicketResult[] = [];
  let board = { id: 0, name };
  for (let index = 0; index < items.length; index += BATCH_SIZE) {
    const batch = items.slice(index, index + BATCH_SIZE);
    const result = await postPmtTickets(boardName.trim(), null, batch);
    board = result.board;
    collected.push(...result.items);
  }

  return {
    board,
    items: collected,
    message: resultMessage(board.name, collected),
  };
}

async function postSheetRows(
  sheetId: number,
  viewerEmail: string,
  meetingDate: string | null,
  items: PmtTicketRequest[],
): Promise<{ board: { id: number; name: string }; items: PmtTicketResult[] }> {
  let token: string;
  try {
    token = getPmtApiToken();
  } catch {
    throw new AppError(
      "PMT is not configured on this PC. Add PMT_API_TOKEN and restart the analyzer.",
      500,
      "PMT_NOT_CONFIGURED",
    );
  }

  const response = await fetch(`${getPmtBaseUrl()}/api/integrations/sheets`, {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      email: viewerEmail,
      sheetId,
      meetingDate,
      items: items.map(toSheetItem),
    }),
  });

  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const message = pmtErrorMessage(response.status, body);
    const routeMissing = !isRecord(body) || /could not be found/i.test(message);
    if (response.status === 404 && routeMissing) {
      throw new AppError(
        "PMT cannot create sheet rows yet. Deploy the sheets integration, then try again.",
        502,
        "PMT_SHEETS_UNAVAILABLE",
      );
    }
    throw new AppError(
      message,
      response.status === 403 || response.status === 404 || response.status === 422
        ? response.status
        : 502,
      response.status === 401 ? "PMT_UNAUTHORIZED" : "PMT_CREATE_FAILED",
    );
  }

  const destination = isRecord(body) && isRecord(body.sheet) ? body.sheet : null;
  if (!isRecord(body) || !destination || !Array.isArray(body.items)) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const id = destination.id;
  const name = destination.name;
  if (typeof id !== "number" || typeof name !== "string") {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const tickets = body.items.filter(isPmtTicket);
  if (tickets.length === 0) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  return { board: { id, name }, items: tickets };
}

export async function createPmtTickets(
  actionItems: MeetingActionItem[],
  destination: PmtBoardOption,
  viewerEmail: string,
): Promise<PmtCreateResult> {
  if (actionItems.length === 0) {
    throw new AppError("Select at least one action item to send.", 400, "NO_ACTION_ITEMS");
  }

  const lanes = new Set(actionItems.map((item) => ticketLaneOf(item)));
  if (lanes.size > 1) {
    throw new AppError(
      "Selected tickets mix development and Sheets. Create them separately.",
      400,
      "PMT_MIXED_LANES",
    );
  }
  const lane = ticketLaneOf(actionItems[0] ?? {});
  const destinationName = destination.name.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
  if (!Number.isInteger(destination.id) || destination.id <= 0 || !destinationName) {
    throw new AppError(
      lane === "sheets" ? "Choose a sheet." : "Choose a board.",
      400,
      "PMT_DESTINATION",
    );
  }
  const mailbox = viewerEmail.trim();
  if (!mailbox.includes("@")) {
    throw new AppError(
      "Sign in to Teams so PMT can check which boards and sheets you can use.",
      401,
      "TEAMS_SIGN_IN",
    );
  }
  const meetingTitle = actionItems[0]?.meetingTitle.trim() ?? "";
  if (
    lane === "development" &&
    actionItems.some((item) => item.meetingTitle.trim() !== meetingTitle)
  ) {
    throw new AppError(
      "Selected tickets come from more than one meeting. Create them one meeting at a time.",
      400,
      "PMT_MIXED_MEETINGS",
    );
  }

  const missingKey = actionItems.find((item) => !item.filingKey.trim());
  if (missingKey) {
    throw new AppError(
      `"${missingKey.title || "Untitled ticket"}" has no filing key, so it cannot be created on PMT.`,
      400,
      "PMT_FILING_KEY",
    );
  }

  let directory: TenantUser[] = [];
  try {
    directory = await listTenantUsers();
  } catch {
    directory = [];
  }

  const seen = new Set<string>();
  const items = actionItems
    .filter((item) => {
      if (seen.has(item.filingKey)) {
        return false;
      }
      seen.add(item.filingKey);
      return true;
    })
    .map((item) => {
      const resolved = resolveDirectoryAssignee(item.pointPerson, item.assigneeUpn, directory);
      if (!resolved) {
        return item;
      }
      return {
        ...item,
        pointPerson: item.pointPerson.trim() || resolved.displayName,
        assigneeDisplayName: resolved.displayName,
        assigneeUpn: resolved.email,
      };
    });

  const requests = items.map((item) => toPmtItem(item, directory, lane));
  const meetingDate = actionItems[0]?.meetingDate ?? null;
  const posted =
    lane === "sheets"
      ? await postSheetRows(destination.id, mailbox, meetingDate, requests)
      : await postPmtTickets(destinationName, meetingDate, requests, {
          boardId: destination.id,
          email: mailbox,
        });
  const tickets = posted.items;
  const boardId = posted.board.id;
  const name = posted.board.name;

  const withWarnings = tickets.map((ticket) => {
    const source = items.find((item) => item.filingKey === ticket.externalKey);
    const resolved = source
      ? resolveDirectoryAssignee(source.pointPerson, source.assigneeUpn, directory)
      : null;
    const withAssignee = resolved
      ? {
          ...ticket,
          assigneeEmail: resolved.email,
          assigneeDisplayName: resolved.displayName,
        }
      : ticket;
    if (withAssignee.warning || resolved) {
      return withAssignee;
    }
    const person = source?.pointPerson.trim() || "the assignee";
    return {
      ...withAssignee,
      assigneeMatched: false,
      warning:
        lane === "sheets"
          ? `No Microsoft 365 user matched ${person}, so the sheet row was left unassigned.`
          : `No Microsoft 365 user matched ${person}, so the PMT ticket was created unassigned.`,
    };
  });

  return {
    board: { id: boardId, name },
    items: withWarnings,
    message: resultMessage(name, withWarnings, lane === "sheets" ? "row" : "ticket"),
  };
}
