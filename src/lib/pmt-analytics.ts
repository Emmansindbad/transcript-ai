import "server-only";

import type {
  PmtAgingBucket,
  PmtAnalyticsReport,
  PmtBoardOption,
  PmtHiddenAgingTicket,
  PmtStatusCount,
} from "@/types/meeting";
import type { AnalyticsRange } from "@/lib/analytics-brief";
import { renderAgingChartBase64 } from "@/lib/analytics-chart";
import { getPmtApiToken, getPmtBaseUrl, getTargetUserId } from "@/lib/env";
import { AppError } from "@/lib/errors";
import { getTeamsAccount } from "@/lib/teams-auth";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isBoard(value: unknown): value is PmtBoardOption {
  return isRecord(value) && typeof value.id === "number" && typeof value.name === "string";
}

function isBucket(value: unknown): value is PmtAgingBucket {
  return isRecord(value) && typeof value.label === "string" && typeof value.count === "number";
}

function isStatus(value: unknown): value is PmtStatusCount {
  return (
    isRecord(value) &&
    typeof value.name === "string" &&
    typeof value.count === "number" &&
    typeof value.color === "string"
  );
}

function isHiddenTicket(value: unknown): value is PmtHiddenAgingTicket {
  return (
    isRecord(value) &&
    typeof value.number === "number" &&
    typeof value.name === "string" &&
    typeof value.column === "string" &&
    typeof value.days === "number" &&
    typeof value.owner === "string" &&
    (value.ownerEmail === undefined || value.ownerEmail === null || typeof value.ownerEmail === "string") &&
    (value.due === null || typeof value.due === "string") &&
    typeof value.url === "string"
  );
}

async function signedInEmail(): Promise<string> {
  const account = await getTeamsAccount();
  const email = account.username?.trim() ?? "";
  if (!account.signedIn || !email.includes("@")) {
    throw new AppError(
      "Sign in to Teams so the analyzer can load the PMT boards for that account.",
      401,
      "TEAMS_SIGN_IN",
    );
  }
  return email;
}

async function pmtGet(path: string, routeMissingMessage?: string): Promise<unknown> {
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

  const response = await fetch(`${getPmtBaseUrl()}${path}`, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
    },
    cache: "no-store",
  });
  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const message = isRecord(body) && typeof body.message === "string" ? body.message : "";
    if (!message || /could not be found/i.test(message)) {
      throw new AppError(
        routeMissingMessage ??
          "PMT analytics is not on the server yet. Deploy the analytics read API, then try again.",
        response.status === 404 ? 404 : 502,
        routeMissingMessage ? "PMT_SHEETS_UNAVAILABLE" : "PMT_ANALYTICS_UNAVAILABLE",
      );
    }
    throw new AppError(message, response.status, "PMT_ANALYTICS_FAILED");
  }
  return body;
}

export function isPmtAnalyticsReport(value: unknown): value is PmtAnalyticsReport {
  return (
    isRecord(value) &&
    isRecord(value.board) &&
    typeof value.board.id === "number" &&
    typeof value.board.name === "string" &&
    typeof value.board.url === "string" &&
    (value.range === "week" || value.range === "month") &&
    typeof value.fromLabel === "string" &&
    typeof value.toLabel === "string" &&
    typeof value.open === "number" &&
    typeof value.closed === "number" &&
    typeof value.overdue === "number" &&
    typeof value.unassigned === "number" &&
    Array.isArray(value.aging) &&
    value.aging.every(isBucket) &&
    Array.isArray(value.statuses) &&
    value.statuses.every(isStatus) &&
    typeof value.chartTitle === "string" &&
    Array.isArray(value.hiddenAging) &&
    value.hiddenAging.every(isHiddenTicket) &&
    typeof value.chartPng === "string"
  );
}

async function listPmtBoardsForEmail(email: string): Promise<PmtBoardOption[]> {
  const mailbox = email.trim();
  if (!mailbox.includes("@")) {
    throw new AppError(
      "A mailbox email is required to list PMT boards.",
      400,
      "PMT_BOARD_MAILBOX",
    );
  }
  const body = await pmtGet(`/api/integrations/boards?email=${encodeURIComponent(mailbox)}`);
  if (!isRecord(body) || !Array.isArray(body.boards) || !body.boards.every(isBoard)) {
    throw new AppError("PMT returned an unexpected board list.", 502, "PMT_ANALYTICS_FAILED");
  }
  return body.boards;
}

async function listPmtSheetsForEmail(email: string): Promise<PmtBoardOption[]> {
  const mailbox = email.trim();
  if (!mailbox.includes("@")) {
    throw new AppError(
      "A mailbox email is required to list PMT sheets.",
      400,
      "PMT_BOARD_MAILBOX",
    );
  }
  const body = await pmtGet(
    `/api/integrations/sheets?email=${encodeURIComponent(mailbox)}`,
    "PMT cannot list sheets yet. Deploy the sheets integration, then try again.",
  );
  if (!isRecord(body) || !Array.isArray(body.sheets) || !body.sheets.every(isBoard)) {
    throw new AppError("PMT returned an unexpected sheet list.", 502, "PMT_ANALYTICS_FAILED");
  }
  return body.sheets;
}

export async function meetingPmtViewerEmail(): Promise<string> {
  try {
    return await signedInEmail();
  } catch (error) {
    const target = getTargetUserId();
    if (target.includes("@")) {
      return target;
    }
    throw error;
  }
}

export async function listMeetingDestinations(
  lane: "development" | "sheets",
): Promise<PmtBoardOption[]> {
  const email = await meetingPmtViewerEmail();
  if (lane === "sheets") {
    return listPmtSheetsForEmail(email);
  }
  return listPmtBoardsForEmail(email);
}

export async function listPmtBoardsForSignedInUser(): Promise<PmtBoardOption[]> {
  return listPmtBoardsForEmail(await signedInEmail());
}

export async function listPmtBoardsForTargetUser(): Promise<PmtBoardOption[]> {
  return listPmtBoardsForEmail(getTargetUserId());
}

export async function getPmtBoardAnalytics(
  boardId: number,
  range: AnalyticsRange,
): Promise<PmtAnalyticsReport> {
  if (!Number.isInteger(boardId) || boardId <= 0) {
    throw new AppError("Choose a PMT board.", 400, "PMT_BOARD");
  }
  const email = await signedInEmail();
  const body = await pmtGet(
    `/api/integrations/boards/${boardId}/analytics?email=${encodeURIComponent(email)}&range=${range}`,
  );
  if (
    !isRecord(body) ||
    !isRecord(body.board) ||
    typeof body.board.id !== "number" ||
    typeof body.board.name !== "string" ||
    typeof body.board.url !== "string" ||
    typeof body.fromLabel !== "string" ||
    typeof body.toLabel !== "string" ||
    typeof body.open !== "number" ||
    typeof body.closed !== "number" ||
    typeof body.overdue !== "number" ||
    typeof body.unassigned !== "number" ||
    !Array.isArray(body.aging) ||
    !body.aging.every(isBucket) ||
    (body.statuses !== undefined && (!Array.isArray(body.statuses) || !body.statuses.every(isStatus))) ||
    !Array.isArray(body.hiddenAging) ||
    !body.hiddenAging.every(isHiddenTicket)
  ) {
    throw new AppError("PMT returned an unexpected analytics report.", 502, "PMT_ANALYTICS_FAILED");
  }

  const aging = body.aging;
  const statuses = Array.isArray(body.statuses) && body.statuses.every(isStatus) ? body.statuses : [];
  const fromDate = Date.parse(body.fromLabel);
  const toDate = Date.parse(body.toLabel);
  const spanDays =
    Number.isNaN(fromDate) || Number.isNaN(toDate) ? null : Math.round((toDate - fromDate) / 86400000);
  const reportedRange =
    body.range === "month" || body.range === "week"
      ? body.range
      : spanDays !== null && spanDays > 10
        ? "month"
        : range;
  return {
    board: { id: body.board.id, name: body.board.name, url: body.board.url },
    range: reportedRange,
    fromLabel: body.fromLabel,
    toLabel: body.toLabel,
    open: body.open,
    closed: body.closed,
    overdue: body.overdue,
    unassigned: body.unassigned,
    aging,
    statuses,
    chartTitle: statuses.length > 0 ? "Where work sits" : "Time in the current column",
    hiddenAging: body.hiddenAging,
    chartPng: renderAgingChartBase64(
      statuses.length > 0
        ? statuses.map((status) => ({ label: status.name, count: status.count, color: status.color }))
        : aging,
    ),
  };
}
