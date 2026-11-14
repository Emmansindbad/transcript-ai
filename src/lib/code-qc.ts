import "server-only";

import { createHash } from "node:crypto";
import type {
  CodeQcAnalyzeResult,
  CodeQcCreateRequest,
  CodeQcCreateResult,
  CodeQcCreateTicketInput,
  CodeQcExistingTicket,
  CodeQcIssue,
  CodeQcProposedTicket,
  CodeQcSendRequest,
  CodeQcSnapshot,
  CodeQcTeamsResult,
} from "@/types/code-qc";
import type { TenantUser } from "@/types/meeting";
import { markdownLink } from "@/lib/chat-links";
import {
  CODE_QC_DEVOPS_CC,
  CODE_QC_MAX_GROUPED_TICKETS,
  displayBoardName,
} from "@/lib/code-qc-brief";
import { getPmtBaseUrl } from "@/lib/env";
import { AppError } from "@/lib/errors";
import { analyzeCodeQcIssues, type CodeQcGeminiTicket } from "@/lib/gemini";
import { listLicensedTenantUsers } from "@/lib/m365";
import {
  createPmtTicketsOnBoard,
  listOpenPmtBoardTickets,
  type PmtOpenTicket,
  type PmtTicketRequest,
} from "@/lib/pmt";
import { listPmtBoardsForTargetUser } from "@/lib/pmt-analytics";
import { listOpenSonarIssues } from "@/lib/sonarqube";
import { markItemsFiled } from "@/lib/store";
import { getTeamsAccount } from "@/lib/teams-auth";
import { sendCodeQcGroupMessage } from "@/lib/teams-message";

const DEFAULT_BOARD_PATTERN = /QC for Frontend Code/i;

function clip(value: string, max: number): string {
  const trimmed = value.trim();
  return trimmed.length <= max ? trimmed : trimmed.slice(0, max);
}

function userEmail(user: TenantUser): string {
  return user.mail?.includes("@") ? user.mail : user.userPrincipalName;
}

function itemType(type: string): "task" | "bug" {
  return type === "BUG" || type === "VULNERABILITY" ? "bug" : "task";
}

function priority(severity: string): "low" | "medium" | "high" {
  if (severity === "BLOCKER" || severity === "CRITICAL") {
    return "high";
  }
  if (severity === "MAJOR") {
    return "medium";
  }
  return "low";
}

function ticketName(issue: CodeQcIssue): string {
  const place = `${issue.component}${typeof issue.line === "number" ? `:${issue.line}` : ""}`;
  const base = issue.message.trim() || "SonarQube issue";
  const named = `${base} — ${place}`;
  if (named.length <= 255) {
    return named;
  }
  return clip(`${clip(base, 180)} — ${issue.key}`, 255);
}

function ticketDescription(issue: CodeQcIssue): string {
  const lines = [
    issue.message,
    "",
    `Type: ${issue.type}`,
    `Severity: ${issue.severity}`,
    `Status: ${issue.status}`,
    `Project: ${issue.projectName} (${issue.project})`,
    `File: ${issue.component}${issue.line ? `:${issue.line}` : ""}`,
  ];
  if (issue.rule) {
    lines.push(`Rule: ${issue.rule}`);
  }
  lines.push(`SonarQube: ${issue.url}`);
  return clip(lines.join("\n"), 10000);
}

function preferredBoardName(boards: Array<{ name: string }>): string | undefined {
  return boards.find((board) => DEFAULT_BOARD_PATTERN.test(displayBoardName(board.name)))
    ?.name;
}

function normalizeTitle(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

export function sonarKeysFromPmtTicket(ticket: PmtOpenTicket): string[] {
  const keys = new Set<string>();
  const external = ticket.externalKey?.trim() ?? "";
  const single = external.match(/^sonar:(.+)$/);
  if (single?.[1]) {
    keys.add(single[1]);
  }
  const hay = `${external}\n${ticket.name}\n${ticket.description ?? ""}`;
  for (const match of hay.matchAll(/[?&](?:open|issues)=([^&\s)#]+)/g)) {
    const key = decodeURIComponent(match[1] ?? "").trim();
    if (key) {
      keys.add(key);
    }
  }
  return [...keys];
}

function toExistingTicket(ticket: PmtOpenTicket): CodeQcExistingTicket {
  return {
    externalKey: ticket.externalKey?.trim() || `pmt:${ticket.number}`,
    number: ticket.number,
    url: ticket.url,
    title: ticket.name.trim() || `#${ticket.number}`,
  };
}

function matchExistingTickets(
  issueKeys: string[],
  title: string,
  externalKey: string,
  openTickets: PmtOpenTicket[],
): PmtOpenTicket[] {
  const wanted = new Set(issueKeys);
  const exact = openTickets.find((ticket) => (ticket.externalKey?.trim() ?? "") === externalKey);
  if (exact) {
    return [exact];
  }
  const covering = openTickets.filter((ticket) =>
    sonarKeysFromPmtTicket(ticket).some((key) => wanted.has(key)),
  );
  const covered = new Set(
    covering.flatMap((ticket) => sonarKeysFromPmtTicket(ticket).filter((key) => wanted.has(key))),
  );
  if (wanted.size > 0 && [...wanted].every((key) => covered.has(key))) {
    const seen = new Set<number>();
    return covering.filter((ticket) => {
      if (seen.has(ticket.number)) {
        return false;
      }
      seen.add(ticket.number);
      return true;
    });
  }
  const needle = normalizeTitle(title);
  if (!needle) {
    return [];
  }
  return openTickets.filter((ticket) => normalizeTitle(ticket.name) === needle).slice(0, 1);
}

async function loadOpenBoardTickets(
  boardName: string,
): Promise<{ board: { id: number; name: string } | null; tickets: PmtOpenTicket[] }> {
  const name = displayBoardName(boardName);
  if (!name) {
    return { board: null, tickets: [] };
  }
  try {
    const listed = await listOpenPmtBoardTickets(name);
    return { board: listed.board, tickets: listed.tickets };
  } catch {
    return { board: null, tickets: [] };
  }
}

function groupExternalKey(keys: string[]): string {
  const unique = [...new Set(keys)].sort();
  if (unique.length === 1) {
    return `sonar:${unique[0]}`;
  }
  const payload = unique.join("|");
  return `sonar-group:${createHash("sha256").update(payload).digest("hex").slice(0, 32)}`;
}

function issueLinkTitle(issue: CodeQcIssue): string {
  const place = `${issue.component}${typeof issue.line === "number" ? `:${issue.line}` : ""}`;
  const base = issue.message.replace(/\s+/g, " ").trim() || "SonarQube issue";
  const named = `${base} — ${place}`;
  return named.length <= 120 ? named : `${clip(base, 80)} — ${issue.key}`;
}

function groupedDescription(issues: CodeQcIssue[], reason: string): string {
  const files = [...new Set(issues.map((issue) => issue.component))];
  const types = [...new Set(issues.map((issue) => issue.type.replace("_", " ")))];
  const project = issues[0]?.projectName || issues[0]?.project || "this repo";
  const worst = issues.reduce(
    (current, issue) =>
      ["BLOCKER", "CRITICAL", "MAJOR", "MINOR", "INFO"].indexOf(issue.severity) <
      ["BLOCKER", "CRITICAL", "MAJOR", "MINOR", "INFO"].indexOf(current)
        ? issue.severity
        : current,
    issues[0]?.severity ?? "MAJOR",
  );
  const header =
    issues.length === 1
      ? `Fix this SonarQube finding in ${project}.`
      : `Fix ${issues.length} related SonarQube findings in ${project}.`;
  const where =
    files.length === 1 ? `File: ${files[0]}` : `Files: ${files.slice(0, 8).join(", ")}${files.length > 8 ? ` (+${files.length - 8} more)` : ""}`;
  const links = issues
    .map((issue) => `- ${markdownLink(issueLinkTitle(issue), issue.url)}`)
    .filter(Boolean);
  const lines = [
    header,
    reason.trim() ? reason.trim() : "Grouped because a developer can clear these in one pass.",
    "",
    where,
    `Highest severity: ${worst}. Types: ${types.join(", ")}.`,
    "",
    "Issues:",
    ...links,
  ];
  return clip(lines.join("\n"), 10000);
}

function proposedType(issues: CodeQcIssue[], declared?: string): "Task" | "Bug" {
  if (declared?.toLowerCase() === "bug") {
    return "Bug";
  }
  if (issues.some((issue) => issue.type === "BUG" || issue.type === "VULNERABILITY")) {
    return "Bug";
  }
  return "Task";
}

function proposedPriority(issues: CodeQcIssue[], declared?: string): "High" | "Medium" | "Low" {
  if (declared === "High" || declared === "Medium" || declared === "Low") {
    return declared;
  }
  const mapped = priority(issues[0]?.severity ?? "MAJOR");
  if (issues.some((issue) => issue.severity === "BLOCKER" || issue.severity === "CRITICAL")) {
    return "High";
  }
  if (mapped === "high") {
    return "High";
  }
  if (mapped === "low") {
    return "Low";
  }
  return "Medium";
}

function toProposedTicket(
  issues: CodeQcIssue[],
  title: string,
  reason: string,
  declaredType?: string,
  declaredPriority?: string,
): CodeQcProposedTicket {
  const keys = [...new Set(issues.map((issue) => issue.key))];
  return {
    id: groupExternalKey(keys),
    title: clip(title.trim() || ticketName(firstIssue(issues)), 255),
    type: proposedType(issues, declaredType),
    priority: proposedPriority(issues, declaredPriority),
    issueKeys: keys,
    description: groupedDescription(issues, reason),
    relatedTo: clip(issues[0]?.projectName || issues[0]?.project || "None", 1000),
  };
}

function fallbackGroups(issues: CodeQcIssue[]): CodeQcProposedTicket[] {
  const byFile = new Map<string, CodeQcIssue[]>();
  for (const issue of issues) {
    const bucket = byFile.get(issue.component) ?? [];
    bucket.push(issue);
    byFile.set(issue.component, bucket);
  }
  const tickets: CodeQcProposedTicket[] = [];
  for (const [file, group] of byFile) {
    const title =
      group.length === 1
        ? ticketName(firstIssue(group))
        : `Fix ${group.length} Sonar issues in ${file}`;
    tickets.push(toProposedTicket(group, title, `Same file: ${file}.`));
  }
  return tickets.sort((left, right) => right.issueKeys.length - left.issueKeys.length);
}

function mergeGeminiTickets(
  issues: CodeQcIssue[],
  grouped: CodeQcGeminiTicket[],
): CodeQcProposedTicket[] {
  const byKey = new Map(issues.map((issue) => [issue.key, issue]));
  const used = new Set<string>();
  const tickets: CodeQcProposedTicket[] = [];

  for (const group of grouped) {
    if (tickets.length >= CODE_QC_MAX_GROUPED_TICKETS) {
      break;
    }
    const selected = group.issueKeys
      .map((key) => byKey.get(key))
      .filter((issue): issue is CodeQcIssue => Boolean(issue))
      .filter((issue) => !used.has(issue.key));
    if (selected.length === 0) {
      continue;
    }
    for (const issue of selected) {
      used.add(issue.key);
    }
    tickets.push(
      toProposedTicket(selected, group.title, group.reason, group.type, group.priority),
    );
  }

  const leftover = issues.filter((issue) => !used.has(issue.key));
  if (leftover.length > 0 && tickets.length < CODE_QC_MAX_GROUPED_TICKETS) {
    const extras = fallbackGroups(leftover).slice(0, CODE_QC_MAX_GROUPED_TICKETS - tickets.length);
    tickets.push(...extras);
  }

  return tickets;
}

function firstIssue(issues: CodeQcIssue[]): CodeQcIssue {
  const issue = issues[0];
  if (!issue) {
    throw new AppError("Select at least one SonarQube issue.", 400, "NO_ISSUES");
  }
  return issue;
}

function attachExistingTickets(
  tickets: CodeQcProposedTicket[],
  openTickets: PmtOpenTicket[],
): CodeQcProposedTicket[] {
  const attached: CodeQcProposedTicket[] = [];
  for (const ticket of tickets) {
    const matches = matchExistingTickets(ticket.issueKeys, ticket.title, ticket.id, openTickets);
    if (matches.length === 0) {
      attached.push(ticket);
      continue;
    }
    const covered = new Set(matches.flatMap((item) => sonarKeysFromPmtTicket(item)));
    for (const match of matches) {
      const keys = ticket.issueKeys.filter((key) => sonarKeysFromPmtTicket(match).includes(key));
      attached.push({
        ...ticket,
        id: match.externalKey?.trim() || `pmt:${match.number}`,
        title: match.name.trim() || ticket.title,
        issueKeys: keys.length > 0 ? keys : ticket.issueKeys,
        existing: toExistingTicket(match),
      });
    }
    const leftover = ticket.issueKeys.filter((key) => !covered.has(key));
    if (leftover.length > 0) {
      attached.push({
        ...ticket,
        id: groupExternalKey(leftover),
        issueKeys: leftover,
      });
    }
  }
  return attached;
}

export async function loadCodeQcSnapshot(): Promise<CodeQcSnapshot> {
  const [sonar, boards, users] = await Promise.all([
    listOpenSonarIssues(),
    listPmtBoardsForTargetUser(),
    listLicensedTenantUsers(),
  ]);
  const defaultBoard = preferredBoardName(boards);
  const listed = defaultBoard ? await loadOpenBoardTickets(defaultBoard) : { tickets: [] };
  const filed = new Set(listed.tickets.flatMap(sonarKeysFromPmtTicket));
  return {
    issues: sonar.issues.map((issue) => ({
      ...issue,
      alreadyFiled: filed.has(issue.key),
    })),
    projects: sonar.projects,
    boards,
    users,
    ...(defaultBoard ? { defaultBoardName: defaultBoard } : {}),
  };
}

export async function analyzeCodeQcTickets(
  issueKeys: string[],
  boardName?: string,
): Promise<CodeQcAnalyzeResult> {
  const keys = [...new Set(issueKeys.map((key) => key.trim()).filter(Boolean))];
  if (keys.length === 0) {
    throw new AppError("Select at least one SonarQube issue to analyze.", 400, "NO_ISSUES");
  }

  const [sonar, listed] = await Promise.all([
    listOpenSonarIssues(),
    loadOpenBoardTickets(boardName ?? ""),
  ]);
  const openTickets = listed.tickets;
  const filed = new Set(openTickets.flatMap(sonarKeysFromPmtTicket));
  const byKey = new Map(sonar.issues.map((issue) => [issue.key, issue]));
  const selected = keys.map((key) => {
    const issue = byKey.get(key);
    if (!issue) {
      throw new AppError(`SonarQube issue ${key} is no longer open.`, 400, "ISSUE_NOT_FOUND");
    }
    return issue;
  });
  const open = selected.filter((issue) => !filed.has(issue.key));
  const source = open.length > 0 ? open : selected;

  let grouped: CodeQcGeminiTicket[] = [];
  let modelUsed = "";
  try {
    const analyzed = await analyzeCodeQcIssues(source);
    grouped = analyzed.tickets;
    modelUsed = analyzed.modelUsed;
  } catch {
    grouped = [];
  }

  const raw = grouped.length > 0 ? mergeGeminiTickets(source, grouped) : fallbackGroups(source);
  const used = new Set(raw.flatMap((ticket) => ticket.issueKeys));
  const leftover = selected.filter((issue) => !used.has(issue.key));
  const extras =
    leftover.length > 0
      ? leftover.map((issue) => toProposedTicket([issue], ticketName(issue), "Already on PMT."))
      : [];
  const tickets = attachExistingTickets([...raw, ...extras], openTickets);
  if (tickets.length === 0) {
    throw new AppError("Gemini did not produce any CODE QC tickets.", 502, "NO_GROUPED_TICKETS");
  }

  const issueCount = tickets.reduce((sum, ticket) => sum + ticket.issueKeys.length, 0);
  const already = tickets.filter((ticket) => ticket.existing).length;
  const fresh = tickets.length - already;
  const summary =
    fresh === 0
      ? `These ${tickets.length} ticket${tickets.length === 1 ? "" : "s"} are already on PMT.`
      : already > 0
        ? `Gemini grouped ${issueCount} issues into ${fresh} new ticket${fresh === 1 ? "" : "s"}; ${already} already on PMT.`
        : tickets.length === 1
          ? `Gemini grouped ${issueCount} issue${issueCount === 1 ? "" : "s"} into 1 PMT ticket.`
          : `Gemini grouped ${issueCount} issues into ${tickets.length} PMT tickets.`;
  return {
    tickets,
    modelUsed,
    message: summary,
  };
}

async function resolveSelectedIssues(keys: string[]): Promise<CodeQcIssue[]> {
  const unique = [...new Set(keys.map((key) => key.trim()).filter(Boolean))];
  if (unique.length === 0) {
    throw new AppError("Select at least one SonarQube issue.", 400, "NO_ISSUES");
  }
  const sonar = await listOpenSonarIssues();
  const byKey = new Map(sonar.issues.map((issue) => [issue.key, issue]));
  return unique.map((key) => {
    const issue = byKey.get(key);
    if (!issue) {
      throw new AppError(`SonarQube issue ${key} is no longer open.`, 400, "ISSUE_NOT_FOUND");
    }
    return issue;
  });
}

function pmtPriorityLabel(value: "High" | "Medium" | "Low"): "high" | "medium" | "low" {
  if (value === "High") {
    return "high";
  }
  if (value === "Low") {
    return "low";
  }
  return "medium";
}

function groupedPmtItems(
  tickets: CodeQcCreateTicketInput[],
  issuesByKey: Map<string, CodeQcIssue>,
  assigneeEmail?: string,
): PmtTicketRequest[] {
  return tickets.map((ticket) => {
    const issues = ticket.issueKeys
      .map((key) => issuesByKey.get(key))
      .filter((issue): issue is CodeQcIssue => Boolean(issue));
    if (issues.length === 0) {
      throw new AppError("A proposed ticket had no open SonarQube issues.", 400, "ISSUE_NOT_FOUND");
    }
    const proposed = toProposedTicket(
      issues,
      ticket.title,
      "Grouped by Gemini from the repo findings.",
      ticket.type,
      ticket.priority,
    );
    const request: PmtTicketRequest = {
      externalKey: proposed.id,
      name: proposed.title,
      itemType: proposed.type === "Bug" ? "bug" : "task",
      priority: pmtPriorityLabel(proposed.priority),
      description: proposed.description,
      relatedTo: proposed.relatedTo,
    };
    if (assigneeEmail) {
      request.assigneeEmail = assigneeEmail;
    }
    return request;
  });
}

export async function createCodeQcTickets(
  input: CodeQcCreateRequest,
): Promise<CodeQcCreateResult> {
  const boardName = input.boardName.trim();
  if (!boardName) {
    throw new AppError("Choose a PMT board for these tickets.", 400, "PMT_BOARD_NAME");
  }

  const users = await listLicensedTenantUsers();
  const byId = new Map(users.map((user) => [user.id, user]));
  let assigneeEmail: string | undefined;
  if (input.assignAllUserId) {
    const assignee = byId.get(input.assignAllUserId);
    if (!assignee) {
      throw new AppError(
        "The Assign all person is not in the licensed Microsoft 365 directory.",
        400,
        "ASSIGNEE_NOT_FOUND",
      );
    }
    assigneeEmail = userEmail(assignee);
  }

  const groupedInput = (input.tickets ?? []).filter((ticket) => ticket.issueKeys.length > 0);
  const items: PmtTicketRequest[] = [];
  const memberKeys: string[] = [];
  const reused: CodeQcCreateResult["items"] = [];
  const listed = await loadOpenBoardTickets(boardName);
  const openTickets = listed.tickets;

  if (groupedInput.length > 0) {
    const allKeys = groupedInput.flatMap((ticket) => ticket.issueKeys);
    const selected = await resolveSelectedIssues(allKeys);
    const issuesByKey = new Map(selected.map((issue) => [issue.key, issue]));
    const fresh: CodeQcCreateTicketInput[] = [];
    for (const ticket of groupedInput) {
      const matches = matchExistingTickets(
        ticket.issueKeys,
        ticket.title,
        groupExternalKey(ticket.issueKeys),
        openTickets,
      );
      if (matches.length > 0) {
        const covered = new Set(matches.flatMap((item) => sonarKeysFromPmtTicket(item)));
        for (const match of matches) {
          const existing = toExistingTicket(match);
          if (reused.some((item) => item.number === existing.number)) {
            continue;
          }
          reused.push({
            externalKey: existing.externalKey,
            id: match.id,
            number: existing.number,
            url: existing.url,
            title: existing.title,
            created: false,
          });
        }
        const leftover = ticket.issueKeys.filter((key) => !covered.has(key));
        if (leftover.length > 0) {
          fresh.push({ ...ticket, issueKeys: leftover });
        }
        continue;
      }
      fresh.push(ticket);
    }
    if (fresh.length > 0) {
      items.push(...groupedPmtItems(fresh, issuesByKey, assigneeEmail));
      memberKeys.push(...fresh.flatMap((ticket) => ticket.issueKeys.map((key) => `sonar:${key}`)));
    }
  } else {
    const keys = [...new Set((input.issueKeys ?? []).map((key) => key.trim()).filter(Boolean))];
    const selected = await resolveSelectedIssues(keys);
    for (const issue of selected) {
      const matches = matchExistingTickets(
        [issue.key],
        ticketName(issue),
        `sonar:${issue.key}`,
        openTickets,
      );
      if (matches[0]) {
        const existing = toExistingTicket(matches[0]);
        if (!reused.some((item) => item.number === existing.number)) {
          reused.push({
            externalKey: existing.externalKey,
            id: matches[0].id,
            number: existing.number,
            url: existing.url,
            title: existing.title,
            created: false,
          });
        }
        continue;
      }
      const request: PmtTicketRequest = {
        externalKey: `sonar:${issue.key}`,
        name: ticketName(issue),
        itemType: itemType(issue.type),
        priority: priority(issue.severity),
        description: ticketDescription(issue),
        relatedTo: clip(issue.projectName || issue.project, 1000),
      };
      if (assigneeEmail) {
        request.assigneeEmail = assigneeEmail;
      }
      items.push(request);
      memberKeys.push(`sonar:${issue.key}`);
    }
  }

  const created =
    items.length > 0
      ? await createPmtTicketsOnBoard(boardName, items)
      : {
          board: listed.board ?? { id: 0, name: boardName },
          items: [],
          message: reused.length > 0 ? `${reused.length} already on PMT.` : "",
        };
  if (created.items.length > 0) {
    await markItemsFiled([
      ...created.items.map((item) => item.externalKey),
      ...memberKeys,
    ]);
  }

  const titles = new Map(items.map((item) => [item.externalKey, item.name]));
  const createdItems = created.items.map((item) => ({
    ...item,
    title: titles.get(item.externalKey)?.trim() || `#${item.number}`,
  }));
  const seen = new Set(reused.map((item) => item.number));
  const merged = [...reused, ...createdItems.filter((item) => !seen.has(item.number))];
  if (merged.length === 0) {
    throw new AppError("Those tickets are already on PMT, or none could be created.", 400, "NO_TICKETS");
  }

  const createdCount = createdItems.filter((item) => item.created).length;
  const alreadyCount = merged.length - createdCount;
  const parts: string[] = [];
  if (createdCount > 0) {
    parts.push(
      createdCount === 1
        ? `Created 1 ticket on ${displayBoardName(created.board.name || boardName)}`
        : `Created ${createdCount} tickets on ${displayBoardName(created.board.name || boardName)}`,
    );
  }
  if (alreadyCount > 0) {
    parts.push(alreadyCount === 1 ? "1 was already on PMT" : `${alreadyCount} were already on PMT`);
  }

  return {
    board: created.board,
    items: merged,
    message: `${parts.join(". ")}.`,
  };
}

function isPmtTicketLink(url: string): boolean {
  try {
    return new URL(url).host === new URL(getPmtBaseUrl()).host;
  } catch {
    return false;
  }
}

export async function sendCodeQcTicketsToGroup(
  input: CodeQcSendRequest,
): Promise<CodeQcTeamsResult> {
  const boardName = displayBoardName(input.boardName);
  if (!boardName) {
    throw new AppError("Choose a PMT board so the Teams group has a title.", 400, "PMT_BOARD_NAME");
  }
  const account = await getTeamsAccount();
  if (!account.signedIn) {
    throw new AppError(
      "Sign in to Teams so the analyzer can create the group and send the ticket links.",
      401,
      "TEAMS_SIGN_IN",
    );
  }

  const tickets = input.tickets.filter(
    (ticket) =>
      ticket.externalKey.trim() &&
      Number.isInteger(ticket.number) &&
      ticket.number > 0 &&
      isPmtTicketLink(ticket.url),
  );
  if (tickets.length === 0) {
    throw new AppError("Create the PMT tickets first, then send them to Teams.", 400, "NO_TICKETS");
  }

  const users = await listLicensedTenantUsers();
  const byId = new Map(users.map((user) => [user.id, user]));
  const groupById = new Map<string, TenantUser>();
  for (const id of input.groupUserIds) {
    const user = byId.get(id);
    if (user) {
      groupById.set(user.id, user);
    }
  }
  for (const cc of CODE_QC_DEVOPS_CC) {
    const email = cc.email.toLowerCase();
    const name = cc.displayName.toLowerCase();
    const user = users.find((candidate) => {
      const mail = userEmail(candidate).toLowerCase();
      return mail === email || candidate.displayName.toLowerCase() === name;
    });
    if (user) {
      groupById.set(user.id, user);
    }
  }
  const group = [...groupById.values()];
  if (group.length < 1) {
    throw new AppError(
      "Tick at least one licensed user for the Teams group. You are added automatically.",
      400,
      "TEAMS_GROUP",
    );
  }

  return sendCodeQcGroupMessage(
    boardName,
    group.map((user) => user.id),
    input.note,
    tickets,
  );
}
