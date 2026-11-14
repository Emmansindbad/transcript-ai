import "server-only";

import type { CodeQcIssue, CodeQcProject } from "@/types/code-qc";
import { getSonarBaseUrl, getSonarPassword, getSonarUsername } from "@/lib/env";
import { AppError } from "@/lib/errors";

const PAGE_SIZE = 500;

interface SonarIssue {
  key?: string;
  message?: string;
  type?: string;
  severity?: string;
  status?: string;
  project?: string;
  component?: string;
  line?: number;
  rule?: string;
}

interface SonarIssuesResponse {
  total?: number;
  issues?: SonarIssue[];
  components?: Array<{ key?: string; name?: string; qualifier?: string }>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function sonarAuthHeader(): string {
  const token = Buffer.from(
    `${getSonarUsername()}:${getSonarPassword()}`,
    "utf8",
  ).toString("base64");
  return `Basic ${token}`;
}

async function sonarGet(path: string): Promise<unknown> {
  const response = await fetch(`${getSonarBaseUrl()}${path}`, {
    headers: {
      Accept: "application/json",
      Authorization: sonarAuthHeader(),
    },
    cache: "no-store",
  });
  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const message =
      isRecord(body) && typeof body.errors === "object"
        ? "SonarQube rejected the request."
        : `SonarQube returned ${response.status}.`;
    throw new AppError(
      response.status === 401
        ? "SonarQube login failed. Check SONARQUBE_USERNAME and SONARQUBE_PASSWORD."
        : message,
      response.status === 401 ? 401 : 502,
      response.status === 401 ? "SONAR_UNAUTHORIZED" : "SONAR_FETCH_FAILED",
    );
  }
  return body;
}

function issueUrl(project: string, key: string): string {
  const base = getSonarBaseUrl();
  const params = new URLSearchParams({
    id: project,
    open: key,
  });
  return `${base}/project/issues?${params.toString()}`;
}

function fileName(component: string): string {
  const parts = component.split(":");
  return parts[parts.length - 1] || component;
}

const SONAR_CACHE_TTL_MS = 2 * 60 * 1000;
let sonarCache:
  | { expiresAt: number; value: { issues: CodeQcIssue[]; projects: CodeQcProject[] } }
  | null = null;

export async function listOpenSonarIssues(): Promise<{
  issues: CodeQcIssue[];
  projects: CodeQcProject[];
}> {
  if (sonarCache && sonarCache.expiresAt > Date.now()) {
    return sonarCache.value;
  }

  const collected: CodeQcIssue[] = [];
  const projectNames = new Map<string, string>();
  let page = 1;
  let total = Number.POSITIVE_INFINITY;

  while (collected.length < total) {
    const body = (await sonarGet(
      `/api/issues/search?issueStatuses=OPEN,CONFIRMED&s=IMPACT_RANK&ps=${PAGE_SIZE}&p=${page}`,
    )) as SonarIssuesResponse;
    total = typeof body.total === "number" ? body.total : collected.length;
    for (const component of body.components ?? []) {
      if (component.key && component.name && component.qualifier === "TRK") {
        projectNames.set(component.key, component.name);
      }
    }
    for (const item of body.issues ?? []) {
      if (!item.key || !item.project) {
        continue;
      }
      collected.push({
        key: item.key,
        message: item.message?.trim() || "SonarQube issue",
        type: item.type ?? "CODE_SMELL",
        severity: item.severity ?? "MAJOR",
        status: item.status ?? "OPEN",
        project: item.project,
        projectName: projectNames.get(item.project) ?? item.project,
        component: fileName(item.component ?? item.project),
        ...(typeof item.line === "number" ? { line: item.line } : {}),
        ...(item.rule ? { rule: item.rule } : {}),
        url: issueUrl(item.project, item.key),
      });
    }
    if ((body.issues ?? []).length === 0) {
      break;
    }
    page += 1;
  }

  const counts = new Map<string, CodeQcProject>();
  for (const issue of collected) {
    const existing = counts.get(issue.project);
    if (existing) {
      existing.issueCount += 1;
      continue;
    }
    counts.set(issue.project, {
      key: issue.project,
      name: issue.projectName,
      issueCount: 1,
    });
  }

  const value = {
    issues: collected,
    projects: [...counts.values()].sort((left, right) =>
      left.name.localeCompare(right.name),
    ),
  };
  sonarCache = { expiresAt: Date.now() + SONAR_CACHE_TTL_MS, value };
  return value;
}
