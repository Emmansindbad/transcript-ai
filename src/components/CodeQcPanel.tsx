"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type {
  CodeQcAnalyzeResult,
  CodeQcCreateResult,
  CodeQcIssue,
  CodeQcProposedTicket,
  CodeQcSnapshot,
} from "@/types/code-qc";
import type { TeamsAccountStatus } from "@/types/meeting";
import { noteSegments } from "@/lib/analytics-brief";
import {
  CODE_QC_CREATE_CHUNK,
  CODE_QC_DEVOPS_CC,
  CODE_QC_MORE_BATCH,
  defaultCodeQcNote,
  displayBoardName,
  VISIBLE_CODE_QC_TICKETS,
  type CodeQcNoteTicket,
} from "@/lib/code-qc-brief";
import {
  fieldClass,
  labelClass,
  primaryButtonClass,
  quietButtonClass,
  secondaryButtonClass,
} from "@/components/ui-classes";

interface CodeQcPanelProps {
  teamsAccount: TeamsAccountStatus | null;
  onTeamsSignIn: () => void;
}

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

function userEmail(user: CodeQcSnapshot["users"][number]): string {
  return user.mail?.includes("@") ? user.mail : user.userPrincipalName;
}

function severityClass(severity: string): string {
  if (severity === "BLOCKER" || severity === "CRITICAL") {
    return "text-red-700";
  }
  if (severity === "MAJOR") {
    return "text-amber-700";
  }
  return "text-zinc-500";
}

function isCcUser(user: CodeQcSnapshot["users"][number]): boolean {
  const email = userEmail(user).toLowerCase();
  const name = user.displayName.toLowerCase();
  return CODE_QC_DEVOPS_CC.some(
    (person) => person.email.toLowerCase() === email || person.displayName.toLowerCase() === name,
  );
}

function ccUserIds(users: CodeQcSnapshot["users"]): string[] {
  return users.filter(isCcUser).map((user) => user.id);
}

function MarkdownLinks({ text }: { text: string }): React.ReactNode {
  const parts = text.split(/(\[[^\]]+\]\(https?:\/\/[^\s)]+\))/g);
  return (
    <>
      {parts.map((part, index) => {
        const match = part.match(/^\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)$/);
        if (!match) {
          return (
            <span key={`${part}-${index}`}>
              {part}
            </span>
          );
        }
        return (
          <a
            key={`${match[2]}-${index}`}
            href={match[2]}
            target="_blank"
            rel="noreferrer"
            className="text-indigo-700 hover:underline"
          >
            {match[1]}
          </a>
        );
      })}
    </>
  );
}

function NotePreview({ text }: { text: string }): React.ReactNode {
  return (
    <div className="mt-2 whitespace-pre-wrap text-sm text-zinc-700">
      {noteSegments(text).map((segment, index) =>
        segment.mention ? (
          <span key={`${segment.text}-${index}`} className="font-medium text-indigo-700">
            {segment.text}
          </span>
        ) : (
          <MarkdownLinks key={`${segment.text}-${index}`} text={segment.text} />
        ),
      )}
    </div>
  );
}

export function CodeQcPanel({
  teamsAccount,
  onTeamsSignIn,
}: CodeQcPanelProps): React.ReactNode {
  const signedIn = Boolean(teamsAccount?.signedIn);
  const [snapshot, setSnapshot] = useState<CodeQcSnapshot | null>(null);
  const [boardName, setBoardName] = useState("");
  const [projectKey, setProjectKey] = useState("all");
  const [assignAllUserId, setAssignAllUserId] = useState("");
  const [groupIds, setGroupIds] = useState<Set<string>>(new Set());
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [query, setQuery] = useState("");
  const [note, setNote] = useState("");
  const [noteTouched, setNoteTouched] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [result, setResult] = useState<CodeQcCreateResult | null>(null);
  const [proposals, setProposals] = useState<CodeQcProposedTicket[]>([]);
  const [selectedProposalIds, setSelectedProposalIds] = useState<Set<string>>(new Set());
  const [isLoading, setIsLoading] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isCreating, setIsCreating] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [showMoreLinks, setShowMoreLinks] = useState(false);
  const [showAllIssues, setShowAllIssues] = useState(false);
  const shareRef = useRef<HTMLDivElement | null>(null);
  const createAbort = useRef<AbortController | null>(null);

  const applySnapshot = useCallback((data: CodeQcSnapshot): void => {
    const nextBoard = data.defaultBoardName ?? data.boards[0]?.name ?? "";
    const openKeys = data.issues.filter((issue) => !issue.alreadyFiled).map((issue) => issue.key);
    const defaultKeys = openKeys.length > 0 ? openKeys : data.issues.map((issue) => issue.key);
    setSnapshot(data);
    setBoardName(nextBoard);
    setGroupIds(new Set());
    setSelectedKeys(new Set(defaultKeys));
    setAssignAllUserId("");
    setProjectKey("all");
    setNoteTouched(false);
    setNote("");
    setProposals([]);
    setSelectedProposalIds(new Set());
    setShowMoreLinks(false);
    setShowAllIssues(false);
  }, []);

  const load = useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    setStatus(null);
    setResult(null);
    setProposals([]);
    setSelectedProposalIds(new Set());
    try {
      const response = await fetch("/api/code-qc");
      const data = (await response.json()) as CodeQcSnapshot & ApiErrorResponse;
      if (!response.ok) {
        setSnapshot(null);
        setError(data.error ?? "Could not load CODE QC.");
        return;
      }
      applySnapshot(data);
    } catch {
      setSnapshot(null);
      setError("Could not load CODE QC.");
    } finally {
      setIsLoading(false);
    }
  }, [applySnapshot]);

  useEffect(() => {
    void load();
  }, [load]);

  const shareTickets = useMemo((): CodeQcNoteTicket[] => {
    if (result && result.items.length > 0) {
      return result.items.map((item) => ({
        title: item.title || `Ticket #${item.number}`,
        number: item.number,
        url: item.url,
      }));
    }
    return proposals
      .filter((ticket) => selectedProposalIds.has(ticket.id) || ticket.existing)
      .map((ticket) =>
        ticket.existing
          ? {
              title: ticket.existing.title || ticket.title,
              number: ticket.existing.number,
              url: ticket.existing.url,
            }
          : { title: ticket.title },
      );
  }, [proposals, result, selectedProposalIds]);

  const assignee = snapshot?.users.find((user) => user.id === assignAllUserId);

  useEffect(() => {
    if (!snapshot || noteTouched) {
      return;
    }
    if (shareTickets.length === 0) {
      return;
    }
    setNote(
      defaultCodeQcNote({
        boardName,
        assigneeName: assignee?.displayName,
        tickets: shareTickets,
      }),
    );
  }, [assignee?.displayName, boardName, noteTouched, shareTickets, snapshot]);

  useEffect(() => {
    if (!snapshot) {
      return;
    }
    setGroupIds((current) => {
      const next = new Set(current);
      let changed = false;
      for (const id of ccUserIds(snapshot.users)) {
        if (!next.has(id)) {
          next.add(id);
          changed = true;
        }
      }
      if (assignAllUserId && !next.has(assignAllUserId)) {
        next.add(assignAllUserId);
        changed = true;
      }
      return changed ? next : current;
    });
  }, [assignAllUserId, snapshot]);

  const visibleIssues = useMemo(() => {
    const issues = snapshot?.issues ?? [];
    const needle = query.trim().toLowerCase();
    return issues.filter((issue) => {
      if (projectKey !== "all" && issue.project !== projectKey) {
        return false;
      }
      if (!needle) {
        return true;
      }
      return (
        issue.message.toLowerCase().includes(needle) ||
        issue.component.toLowerCase().includes(needle) ||
        issue.key.toLowerCase().includes(needle) ||
        issue.projectName.toLowerCase().includes(needle)
      );
    });
  }, [projectKey, query, snapshot]);

  const toggleProposal = (id: string): void => {
    setSelectedProposalIds((current) => {
      const next = new Set(current);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const toggleIssue = (key: string): void => {
    setSelectedKeys((current) => {
      const next = new Set(current);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  const toggleGroup = (id: string): void => {
    setGroupIds((current) => {
      const next = new Set(current);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const selectVisible = (checked: boolean): void => {
    setSelectedKeys((current) => {
      const next = new Set(current);
      for (const issue of visibleIssues) {
        if (checked) {
          next.add(issue.key);
        } else {
          next.delete(issue.key);
        }
      }
      return next;
    });
  };

  const analyzeTickets = async (): Promise<void> => {
    const keys = [...selectedKeys];
    if (keys.length === 0) {
      return;
    }
    setIsAnalyzing(true);
    setError(null);
    setResult(null);
    setStatus(`Gemini is grouping ${keys.length} issue${keys.length === 1 ? "" : "s"}…`);
    try {
      const response = await fetch("/api/code-qc/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ issueKeys: keys, boardName }),
      });
      const data = (await response.json()) as CodeQcAnalyzeResult & ApiErrorResponse;
      if (!response.ok) {
        setProposals([]);
        setSelectedProposalIds(new Set());
        setError(data.error ?? "Could not analyze CODE QC issues.");
        setStatus(null);
        return;
      }
      setProposals(data.tickets);
      const freshIds = data.tickets.filter((ticket) => !ticket.existing).map((ticket) => ticket.id);
      setSelectedProposalIds(new Set(freshIds.length > 0 ? freshIds : data.tickets.map((ticket) => ticket.id)));
      const existingItems = data.tickets.flatMap((ticket) =>
        ticket.existing
          ? [
              {
                externalKey: ticket.existing.externalKey,
                id: ticket.existing.number,
                number: ticket.existing.number,
                url: ticket.existing.url,
                title: ticket.existing.title,
                created: false,
              },
            ]
          : [],
      );
      setResult(
        existingItems.length > 0
          ? { board: { id: 0, name: boardName }, items: existingItems, message: data.message }
          : null,
      );
      setStatus(data.message);
      setNoteTouched(false);
    } catch {
      setError("Could not analyze CODE QC issues.");
      setStatus(null);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const createTickets = async (): Promise<void> => {
    if (!snapshot) {
      return;
    }
    const selected = proposals.filter((ticket) => selectedProposalIds.has(ticket.id));
    const tickets = selected.filter((ticket) => !ticket.existing);
    const alreadyOnPmt: CodeQcCreateResult["items"] = selected.flatMap((ticket) =>
      ticket.existing
        ? [
            {
              externalKey: ticket.existing.externalKey,
              id: ticket.existing.number,
              number: ticket.existing.number,
              url: ticket.existing.url,
              title: ticket.existing.title,
              created: false,
            },
          ]
        : [],
    );
    if (tickets.length === 0) {
      if (alreadyOnPmt.length === 0) {
        return;
      }
      setResult({
        board: result?.board ?? { id: 0, name: boardName },
        items: alreadyOnPmt,
        message: `${alreadyOnPmt.length} already on PMT.`,
      });
      setStatus(
        alreadyOnPmt.length === 1
          ? "That ticket is already on PMT. Share the numbered link to Teams."
          : `${alreadyOnPmt.length} tickets are already on PMT. Share the numbered links to Teams.`,
      );
      setNoteTouched(false);
      return;
    }
    createAbort.current?.abort();
    const controller = new AbortController();
    createAbort.current = controller;
    setIsCreating(true);
    setError(null);
    setStatus(
      `Creating tickets 1–${Math.min(CODE_QC_CREATE_CHUNK, tickets.length)} of ${tickets.length}…`,
    );
    setShowMoreLinks(false);

    const collected: CodeQcCreateResult["items"] = alreadyOnPmt.length > 0
      ? [...alreadyOnPmt]
      : result?.items
        ? [...result.items]
        : [];
    const seen = new Set(collected.map((item) => item.externalKey));
    let board = result?.board ?? { id: 0, name: boardName };
    let lastError: string | null = null;
    let stopped = false;

    try {
      for (let index = 0; index < tickets.length; index += CODE_QC_CREATE_CHUNK) {
        if (controller.signal.aborted) {
          stopped = true;
          break;
        }
        const chunk = tickets.slice(index, index + CODE_QC_CREATE_CHUNK);
        setStatus(
          `Creating tickets ${index + 1}–${Math.min(index + chunk.length, tickets.length)} of ${tickets.length}…`,
        );
        const response = await fetch("/api/code-qc", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          signal: controller.signal,
          body: JSON.stringify({
            boardName,
            tickets: chunk.map((ticket) => ({
              title: ticket.title,
              type: ticket.type,
              priority: ticket.priority,
              issueKeys: ticket.issueKeys,
            })),
            ...(assignAllUserId ? { assignAllUserId } : {}),
          }),
        });
        const data = (await response.json()) as CodeQcCreateResult & ApiErrorResponse;
        if (!response.ok) {
          lastError = data.error ?? "Could not create PMT tickets.";
          break;
        }
        board = data.board;
        chunk.forEach((ticket, ticketIndex) => {
          const item = data.items[ticketIndex];
          if (!item || seen.has(item.externalKey)) {
            return;
          }
          seen.add(item.externalKey);
          collected.push({
            ...item,
            title: item.title || ticket.title,
          });
        });
        for (const item of data.items) {
          if (seen.has(item.externalKey)) {
            continue;
          }
          seen.add(item.externalKey);
          collected.push({
            ...item,
            title: item.title || `#${item.number}`,
          });
        }
        const filedKeys = new Set(chunk.flatMap((ticket) => ticket.issueKeys));
        setSnapshot((current) =>
          current
            ? {
                ...current,
                issues: current.issues.map((issue) =>
                  filedKeys.has(issue.key) ? { ...issue, alreadyFiled: true } : issue,
                ),
              }
            : current,
        );
        setSelectedKeys((current) => {
          const next = new Set(current);
          for (const key of filedKeys) {
            next.delete(key);
          }
          return next;
        });
        setResult({
          board: data.board,
          items: [...collected],
          message: data.message,
        });
      }

      if (collected.length === 0) {
        setError(lastError ?? "Could not create PMT tickets.");
        setStatus(null);
        return;
      }
      if (lastError) {
        setError(
          `${lastError} ${collected.length} of ${tickets.length} grouped tickets are on PMT. Create again for the rest.`,
        );
      }
      setStatus(
        stopped
          ? `Stopped. ${collected.length} ticket${collected.length === 1 ? "" : "s"} already on ${displayBoardName(board.name)}.`
          : lastError
            ? `Stopped after ${collected.length} ticket${collected.length === 1 ? "" : "s"} on ${displayBoardName(board.name)}.`
            : `Created ${collected.length} grouped ticket${collected.length === 1 ? "" : "s"} on ${displayBoardName(board.name)}.`,
      );
      setProposals((current) => current.filter((ticket) => !selectedProposalIds.has(ticket.id)));
      setSelectedProposalIds(new Set());
      setNoteTouched(false);
      shareRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      const aborted =
        (error instanceof DOMException && error.name === "AbortError") ||
        (error instanceof Error && error.name === "AbortError");
      if (aborted) {
        setStatus(
          collected.length > 0
            ? `Stopped. ${collected.length} ticket${collected.length === 1 ? "" : "s"} already on PMT.`
            : "Stopped. No new tickets were created.",
        );
      } else {
        setError(
          collected.length > 0
            ? `Creating tickets stopped after ${collected.length} were on PMT. Create again for the rest.`
            : "Could not create PMT tickets.",
        );
      }
    } finally {
      if (createAbort.current === controller) {
        createAbort.current = null;
      }
      setIsCreating(false);
    }
  };

  const stopCreating = (): void => {
    createAbort.current?.abort();
  };

  const sendToTeams = async (): Promise<void> => {
    if (!result) {
      return;
    }
    setIsSending(true);
    setError(null);
    setStatus(null);
    try {
      const response = await fetch("/api/code-qc/send", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          boardName: result.board.name || boardName,
          groupUserIds: [...groupIds],
          note,
          tickets: result.items.map((item) => ({
            externalKey: item.externalKey,
            number: item.number,
            url: item.url,
            title: item.title,
          })),
        }),
      });
      const data = (await response.json()) as {
        message?: string;
        teams?: CodeQcCreateResult["teams"];
      } & ApiErrorResponse;
      if (!response.ok) {
        setError(data.error ?? "Could not send to Teams.");
        return;
      }
      setResult({ ...result, teams: data.teams });
      setStatus(data.message ?? "Sent to Teams.");
    } catch {
      setError("Could not send to Teams.");
    } finally {
      setIsSending(false);
    }
  };

  const openVisible = visibleIssues;
  const selectedVisible = openVisible.filter((issue) => selectedKeys.has(issue.key)).length;
  const allVisibleSelected =
    openVisible.length > 0 && selectedVisible === openVisible.length;
  const freshSelected = proposals.filter(
    (ticket) => selectedProposalIds.has(ticket.id) && !ticket.existing,
  ).length;
  const canSendTeams =
    shareTickets.some((ticket) => Boolean(ticket.url && ticket.number)) &&
    groupIds.size >= 1 &&
    signedIn;
  const listedIssues = showAllIssues
    ? visibleIssues
    : visibleIssues.slice(0, CODE_QC_MORE_BATCH);
  const hiddenIssues = Math.max(0, visibleIssues.length - listedIssues.length);
  const resultLinks = result?.items ?? [];
  const hiddenLinks = Math.max(0, resultLinks.length - VISIBLE_CODE_QC_TICKETS);
  const shownLinks = showMoreLinks
    ? resultLinks
    : resultLinks.slice(0, VISIBLE_CODE_QC_TICKETS);

  return (
    <section className="space-y-6">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <h2 className="text-lg font-medium tracking-tight text-zinc-900">CODE QC</h2>
          <p className="mt-1 max-w-2xl text-sm text-zinc-500">
            Select Sonar issues, Analyze so Gemini groups related findings, then create those
            tickets on the QC board. After Gemini proposes tickets, assign someone, tick the group,
            and share a message with numbered ticket links.
          </p>
        </div>
        <button type="button" className={secondaryButtonClass} onClick={() => void load()} disabled={isLoading}>
          {isLoading ? "Loading…" : "Refresh"}
        </button>
      </div>

      {error ? <p className="text-sm text-red-700">{error}</p> : null}
      {status ? <p className="text-sm text-zinc-700">{status}</p> : null}

      {isLoading && !snapshot ? (
        <p className="text-sm text-zinc-500">Loading SonarQube issues, PMT boards, and licensed users…</p>
      ) : null}

      {snapshot ? (
        <>
          <div className="grid gap-4 md:grid-cols-2">
            <label className="block">
              <span className={labelClass}>PMT board</span>
              <select
                className={fieldClass}
                value={boardName}
                onChange={(event) => setBoardName(event.target.value)}
              >
                {snapshot.boards.length === 0 ? (
                  <option value="">No boards found</option>
                ) : null}
                {snapshot.boards.map((board) => (
                  <option key={board.id} value={board.name}>
                    {displayBoardName(board.name)}
                  </option>
                ))}
              </select>
            </label>
            <label className="block">
              <span className={labelClass}>Repo / project</span>
              <select
                className={fieldClass}
                value={projectKey}
                onChange={(event) => setProjectKey(event.target.value)}
              >
                <option value="all">All projects ({snapshot.issues.length})</option>
                {snapshot.projects.map((project) => (
                  <option key={project.key} value={project.key}>
                    {project.name} ({project.issueCount})
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div>
            <div className="mb-2 flex flex-wrap items-end justify-between gap-3">
              <div>
                <h3 className="text-sm font-medium text-zinc-900">Open issues</h3>
                <p className="mt-1 text-xs text-zinc-500">
                  {selectedKeys.size} selected for analysis · {selectedVisible} visible open ·{" "}
                  {snapshot.issues.filter((issue) => issue.alreadyFiled).length} already on PMT
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <input
                  type="search"
                  className={`${fieldClass} w-56`}
                  placeholder="Filter issues"
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                />
                <button
                  type="button"
                  className={quietButtonClass}
                  onClick={() => selectVisible(!allVisibleSelected)}
                >
                  {allVisibleSelected ? "Clear visible" : "Select visible"}
                </button>
              </div>
            </div>
            <div className="overflow-x-auto rounded-md border border-zinc-200">
              <table className="min-w-full text-left text-sm">
                <thead className="bg-zinc-50 text-xs font-medium uppercase tracking-wide text-zinc-500">
                  <tr>
                    <th className="w-10 px-3 py-2">
                      <input
                        type="checkbox"
                        checked={allVisibleSelected}
                        onChange={(event) => selectVisible(event.target.checked)}
                        aria-label="Select visible issues"
                      />
                    </th>
                    <th className="px-3 py-2">Type</th>
                    <th className="px-3 py-2">Severity</th>
                    <th className="px-3 py-2">Issue</th>
                    <th className="px-3 py-2">File</th>
                  </tr>
                </thead>
                <tbody>
                  {listedIssues.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="px-3 py-6 text-center text-zinc-500">
                        No open issues match this filter.
                      </td>
                    </tr>
                  ) : (
                    listedIssues.map((issue: CodeQcIssue) => (
                      <tr
                        key={issue.key}
                        className={`border-t border-zinc-100 ${issue.alreadyFiled ? "bg-zinc-50" : ""}`}
                      >
                        <td className="px-3 py-2 align-top">
                          <input
                            type="checkbox"
                            checked={selectedKeys.has(issue.key)}
                            onChange={() => toggleIssue(issue.key)}
                            aria-label={`Select ${issue.message}`}
                          />
                        </td>
                        <td className="px-3 py-2 align-top text-xs text-zinc-600">
                          {issue.type.replace("_", " ")}
                          {issue.alreadyFiled ? (
                            <span className="mt-1 block text-[11px] text-amber-700">Already on PMT</span>
                          ) : null}
                        </td>
                        <td className={`px-3 py-2 align-top text-xs ${severityClass(issue.severity)}`}>
                          {issue.severity}
                        </td>
                        <td className="px-3 py-2 align-top">
                          <a
                            href={issue.url}
                            target="_blank"
                            rel="noreferrer"
                            className="text-zinc-900 hover:underline"
                          >
                            {issue.message}
                          </a>
                        </td>
                        <td className="px-3 py-2 align-top text-xs text-zinc-500">
                          {issue.component}
                          {issue.line ? `:${issue.line}` : ""}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            {hiddenIssues > 0 ? (
              <button
                type="button"
                className={`${quietButtonClass} mt-2`}
                onClick={() => setShowAllIssues(true)}
              >
                Show all {visibleIssues.length} issues · {hiddenIssues} hidden
              </button>
            ) : null}
          </div>

          {proposals.length > 0 ? (
            <div>
              <h3 className="text-sm font-medium text-zinc-900">
                Proposed tickets ({selectedProposalIds.size} of {proposals.length})
              </h3>
              <p className="mt-1 text-xs text-zinc-500">
                Gemini grouped related findings. Each description lists Sonar issues as [title](link).
                Uncheck any packet you do not want on PMT.
              </p>
              <div className="mt-3 space-y-3">
                {proposals.map((ticket) => (
                  <div key={ticket.id} className="rounded-md border border-zinc-200 p-3">
                    <div className="flex items-start gap-2">
                      <input
                        type="checkbox"
                        className="mt-1"
                        checked={selectedProposalIds.has(ticket.id)}
                        onChange={() => toggleProposal(ticket.id)}
                        aria-label={`Include ${ticket.title}`}
                      />
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-medium text-zinc-900">{ticket.title}</p>
                        <p className="mt-1 text-xs text-zinc-500">
                          {ticket.type} · {ticket.priority} · {ticket.issueKeys.length} issue
                          {ticket.issueKeys.length === 1 ? "" : "s"}
                          {ticket.existing ? (
                            <>
                              {" · "}
                              <a
                                href={ticket.existing.url}
                                target="_blank"
                                rel="noreferrer"
                                className="text-indigo-700 hover:underline"
                              >
                                Already on PMT #{ticket.existing.number}
                              </a>
                            </>
                          ) : null}
                        </p>
                        <div className="mt-2 whitespace-pre-wrap text-sm text-zinc-700">
                          <MarkdownLinks text={ticket.description} />
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : null}

          {proposals.length > 0 || result ? (
            <>
              {proposals.length > 0 ? (
                <label className="block">
                  <span className={labelClass}>Assign all</span>
                  <select
                    className={fieldClass}
                    value={assignAllUserId}
                    onChange={(event) => setAssignAllUserId(event.target.value)}
                  >
                    <option value="">Leave tickets unassigned</option>
                    {snapshot.users.map((user) => (
                      <option key={user.id} value={user.id}>
                        {user.displayName} · {userEmail(user)}
                      </option>
                    ))}
                  </select>
                  <span className="mt-1 block text-xs text-zinc-400">
                    {assignee
                      ? `Every created ticket will be assigned to ${assignee.displayName}, and they are @ mentioned in the message.`
                      : "Choose one licensed user to assign every selected ticket. They are @ mentioned in the Teams message."}
                  </span>
                </label>
              ) : null}

              <div>
                <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                  <h3 className="text-sm font-medium text-zinc-900">
                    Licensed users · group ({groupIds.size})
                  </h3>
                  <div className="flex gap-2">
                    <button
                      type="button"
                      className={quietButtonClass}
                      onClick={() => setGroupIds(new Set(snapshot.users.map((user) => user.id)))}
                    >
                      Tick all
                    </button>
                    <button
                      type="button"
                      className={quietButtonClass}
                      onClick={() =>
                        setGroupIds(
                          new Set([
                            ...ccUserIds(snapshot.users),
                            ...(assignAllUserId ? [assignAllUserId] : []),
                          ]),
                        )
                      }
                    >
                      Clear
                    </button>
                  </div>
                </div>
                <p className="mb-2 text-xs text-zinc-500">
                  Ticked people join a Teams group titled{" "}
                  <span className="font-medium text-zinc-700">
                    {displayBoardName(boardName) || "the PMT board"}
                  </span>
                  . Jehad and Arnold are CC’d for devops. The assignee is added automatically. You
                  are added automatically too.
                </p>
                <div className="grid max-h-56 gap-1 overflow-y-auto rounded-md border border-zinc-200 p-2 sm:grid-cols-2">
                  {snapshot.users.map((user) => (
                    <label
                      key={user.id}
                      className="flex cursor-pointer items-start gap-2 rounded-md px-2 py-1.5 text-sm text-zinc-800 hover:bg-zinc-50"
                    >
                      <input
                        type="checkbox"
                        className="mt-0.5"
                        checked={groupIds.has(user.id)}
                        disabled={isCcUser(user)}
                        onChange={() => toggleGroup(user.id)}
                      />
                      <span>
                        <span className="block font-medium">{user.displayName}</span>
                        <span className="block text-xs text-zinc-500">
                          {userEmail(user)}
                          {isCcUser(user) ? " · CC devops" : ""}
                          {user.id === assignAllUserId ? " · assignee" : ""}
                        </span>
                      </span>
                    </label>
                  ))}
                </div>
                {!signedIn ? (
                  <div className="mt-3 flex flex-wrap items-center gap-3">
                    <p className="text-sm text-zinc-600">
                      Sign in to Teams so you can share the PMT ticket links with the ticked people.
                    </p>
                    <button type="button" className={secondaryButtonClass} onClick={onTeamsSignIn}>
                      Sign in to Teams
                    </button>
                  </div>
                ) : null}
              </div>

              <label className="block">
                <span className={labelClass}>Message</span>
                <textarea
                  className={`${fieldClass} min-h-40`}
                  value={note}
                  onChange={(event) => {
                    setNoteTouched(true);
                    setNote(event.target.value);
                  }}
                  rows={10}
                />
                <span className="mt-1 block text-xs text-zinc-400">
                  Built from the proposed tickets. The assignee is @ mentioned, tickets are numbered
                  hyperlinks after PMT create, and Jehad and Arnold are CC’d for devops.
                </span>
                {note.trim() ? <NotePreview text={note} /> : null}
              </label>
            </>
          ) : null}

          {result ? (
            <div
              ref={shareRef}
              className="space-y-3 rounded-lg border border-zinc-200 bg-zinc-50 p-4"
            >
              {result.teams?.warning ? (
                <p className="text-sm text-amber-700">{result.teams.warning}</p>
              ) : null}
              <div>
                <h3 className="text-sm font-medium text-zinc-900">
                  {result.teams?.sent
                    ? `Shared to Teams · ${result.teams.topic}`
                    : `Created on ${displayBoardName(result.board.name || boardName)}`}
                </h3>
                <p className="mt-1 text-xs text-zinc-500">
                  {result.teams?.sent
                    ? "The ticked people are in that group, with PMT ticket links."
                    : "Tick people in the group list, then share these PMT tickets to Teams."}
                </p>
              </div>
              <ol className="list-decimal space-y-1 pl-5 text-sm text-zinc-700">
                {shownLinks.map((item) => (
                  <li key={item.externalKey}>
                    <a
                      href={item.url}
                      target="_blank"
                      rel="noreferrer"
                      className="text-indigo-700 hover:underline"
                    >
                      #{item.number}
                      {item.title ? ` ${item.title}` : ""}
                    </a>
                  </li>
                ))}
              </ol>
              {hiddenLinks > 0 && !showMoreLinks ? (
                <button
                  type="button"
                  className={quietButtonClass}
                  onClick={() => setShowMoreLinks(true)}
                >
                  More · {hiddenLinks} hidden
                </button>
              ) : null}
              <div className="flex flex-wrap items-center gap-3">
                {!signedIn ? (
                  <button type="button" className={secondaryButtonClass} onClick={onTeamsSignIn}>
                    Sign in to Teams
                  </button>
                ) : null}
                <button
                  type="button"
                  className={primaryButtonClass}
                  disabled={isSending || !canSendTeams}
                  onClick={() => void sendToTeams()}
                >
                  {isSending
                    ? "Sharing…"
                    : result.teams?.sent
                      ? "Share again to Teams"
                      : "Share to Teams"}
                </button>
                <p className="text-xs text-zinc-500">
                  {!signedIn
                    ? "Sign in to Teams, then share."
                    : groupIds.size < 1
                      ? "Tick at least one person for the group."
                      : `${groupIds.size} person${groupIds.size === 1 ? "" : "s"} will be in the group.`}
                </p>
              </div>
            </div>
          ) : null}

          <div className="sticky bottom-0 z-20 -mx-6 border-t border-zinc-200 bg-white/95 px-6 py-3 backdrop-blur">
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                className={proposals.length > 0 ? secondaryButtonClass : primaryButtonClass}
                disabled={isAnalyzing || isCreating || selectedKeys.size === 0}
                onClick={() => void analyzeTickets()}
              >
                {isAnalyzing
                  ? "Analyzing…"
                  : `Analyze ${selectedKeys.size} issue${selectedKeys.size === 1 ? "" : "s"}`}
              </button>
              <button
                type="button"
                className={result ? secondaryButtonClass : primaryButtonClass}
                disabled={
                  isCreating || isAnalyzing || freshSelected === 0 || !boardName
                }
                onClick={() => void createTickets()}
              >
                {isCreating
                  ? "Creating tickets…"
                  : `Create ${freshSelected} ticket${freshSelected === 1 ? "" : "s"} on PMT`}
              </button>
              {isCreating ? (
                <button type="button" className={secondaryButtonClass} onClick={stopCreating}>
                  Stop
                </button>
              ) : null}
              <button
                type="button"
                className={result && !result.teams?.sent ? primaryButtonClass : secondaryButtonClass}
                disabled={isSending || !canSendTeams}
                onClick={() => void sendToTeams()}
              >
                {isSending ? "Sharing…" : "Share to Teams"}
              </button>
              <p className="text-xs text-zinc-500">
                Analyze → create on PMT → share to the ticked people on Teams.
                {!result
                  ? " Create tickets first."
                  : !signedIn
                    ? " Sign in to Teams to share."
                    : groupIds.size < 1
                      ? " Tick people for the group."
                      : ""}
              </p>
            </div>
          </div>
        </>
      ) : null}
    </section>
  );
}
