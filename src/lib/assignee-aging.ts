import "server-only";

import type {
  PmtAgingPerson,
  PmtAssigneeAgingProject,
  PmtAssigneeAgingReport,
  PmtAssigneeAgingTicket,
} from "@/types/meeting";
import { getPmtApiToken, getPmtBaseUrl } from "@/lib/env";
import { AppError } from "@/lib/errors";
import { getTeamsAccount } from "@/lib/teams-auth";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isPerson(value: unknown): value is PmtAgingPerson {
  return isRecord(value) && typeof value.name === "string" && typeof value.email === "string";
}

function isTicket(value: unknown): value is PmtAssigneeAgingTicket {
  return (
    isRecord(value) &&
    typeof value.number === "number" &&
    typeof value.name === "string" &&
    typeof value.column === "string" &&
    typeof value.days === "number" &&
    typeof value.url === "string"
  );
}

function isProject(value: unknown): value is PmtAssigneeAgingProject {
  return (
    isRecord(value) &&
    typeof value.id === "number" &&
    typeof value.name === "string" &&
    typeof value.more === "number" &&
    Array.isArray(value.tickets) &&
    value.tickets.every(isTicket)
  );
}

export function isAssigneeAgingReport(value: unknown): value is PmtAssigneeAgingReport {
  return (
    isRecord(value) &&
    Array.isArray(value.people) &&
    value.people.every(isPerson) &&
    (value.person === null || isPerson(value.person)) &&
    Array.isArray(value.projects) &&
    value.projects.every(isProject) &&
    typeof value.omittedProjects === "number"
  );
}

async function signedInEmail(): Promise<string> {
  const account = await getTeamsAccount();
  const email = account.username?.trim() ?? "";
  if (!account.signedIn || !email.includes("@")) {
    throw new AppError(
      "Sign in to Teams so the analyzer can load aging tickets for that account.",
      401,
      "TEAMS_SIGN_IN",
    );
  }
  return email;
}

async function pmtGet(path: string): Promise<unknown> {
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
        "Aging tickets by person are not on the PMT server yet. Deploy the aging API, then try again.",
        response.status === 404 ? 404 : 502,
        "PMT_AGING_UNAVAILABLE",
      );
    }
    throw new AppError(message, response.status, "PMT_AGING_FAILED");
  }
  return body;
}

export async function getAssigneeAging(assigneeEmail: string): Promise<PmtAssigneeAgingReport> {
  const viewer = await signedInEmail();
  const assignee = assigneeEmail.trim();
  const query = new URLSearchParams({ email: viewer });
  if (assignee.includes("@")) {
    query.set("assignee", assignee);
  }
  const body = await pmtGet(`/api/integrations/aging?${query.toString()}`);
  if (!isAssigneeAgingReport(body)) {
    throw new AppError("PMT returned an unexpected aging list.", 502, "PMT_AGING_FAILED");
  }
  return body;
}
