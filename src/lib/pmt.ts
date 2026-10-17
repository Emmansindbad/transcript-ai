import "server-only";

import type { MeetingActionItem, PmtCreateResult, PmtTicketResult, TenantUser } from "@/types/meeting";
import { resolveDirectoryAssignee } from "@/lib/assignees";
import { normalizeDevTag } from "@/lib/dev-tags";
import { getPmtApiToken, getPmtBaseUrl } from "@/lib/env";
import { AppError } from "@/lib/errors";
import { listTenantUsers } from "@/lib/m365";

interface PmtTicketRequest {
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

function toPmtItem(item: MeetingActionItem, users: TenantUser[]): PmtTicketRequest {
  const name = clip(item.title || item.task || "Untitled ticket", 255);
  const request: PmtTicketRequest = {
    externalKey: item.filingKey,
    name,
    itemType: item.type === "Bug" ? "bug" : "task",
    priority: pmtPriority(item.priority),
  };
  const description = clip(item.description, 10000);
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
  if (devTag) {
    request.devTag = devTag;
  }
  request.relatedTo = relatedTo;
  return request;
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

function resultMessage(boardName: string, items: PmtTicketResult[]): string {
  const created = items.filter((item) => item.created).length;
  const existing = items.length - created;
  const parts: string[] = [];
  if (created > 0) {
    parts.push(
      created === 1
        ? `Created 1 ticket on ${boardName}`
        : `Created ${created} tickets on ${boardName}`,
    );
  }
  if (existing > 0) {
    parts.push(existing === 1 ? "1 was already on PMT" : `${existing} were already on PMT`);
  }
  return `${parts.join(". ")}.`;
}

export async function createPmtTickets(
  actionItems: MeetingActionItem[],
): Promise<PmtCreateResult> {
  if (actionItems.length === 0) {
    throw new AppError("Select at least one action item to send.", 400, "NO_ACTION_ITEMS");
  }

  const boardName = actionItems[0]?.meetingTitle.trim() ?? "";
  if (!boardName) {
    throw new AppError(
      "This meeting has no title, so PMT cannot find a board.",
      400,
      "PMT_BOARD_NAME",
    );
  }
  if (actionItems.some((item) => item.meetingTitle.trim() !== boardName)) {
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
      meetingDate: actionItems[0]?.meetingDate ?? null,
      items: items.map((item) => toPmtItem(item, directory)),
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

  if (!isRecord(body) || !isRecord(body.board) || !Array.isArray(body.items)) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const boardId = body.board.id;
  const name = body.board.name;
  if (typeof boardId !== "number" || typeof name !== "string") {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }
  const tickets = body.items.filter(isPmtTicket);
  if (tickets.length !== body.items.length) {
    throw new AppError("PMT returned an unexpected response.", 502, "PMT_CREATE_FAILED");
  }

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
      warning: `No Microsoft 365 user matched ${person}, so the PMT ticket was created unassigned.`,
    };
  });

  return {
    board: { id: boardId, name },
    items: withWarnings,
    message: resultMessage(name, withWarnings),
  };
}
