"use client";

import { useCallback, useEffect, useState } from "react";
import {
  agingDayLabel,
  defaultAssigneeAgingNote,
  noteSegments,
} from "@/lib/analytics-brief";
import type {
  PmtAssigneeAgingProject,
  PmtAssigneeAgingReport,
  PmtAssigneeAgingTicket,
  TeamsAccountStatus,
} from "@/types/meeting";
import { fieldClass, labelClass, primaryButtonClass } from "@/components/ui-classes";

interface AgingTicketsPanelProps {
  teamsAccount: TeamsAccountStatus | null;
  onTeamsSignIn: () => void;
}

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

function ticketKey(projectId: number, ticket: PmtAssigneeAgingTicket): string {
  return `${projectId}:${ticket.number}`;
}

function selectedReport(
  report: PmtAssigneeAgingReport,
  included: ReadonlySet<string>,
): PmtAssigneeAgingReport {
  const projects = report.projects.flatMap((project) => {
    const tickets = project.tickets.filter((ticket) => included.has(ticketKey(project.id, ticket)));
    if (tickets.length === 0) {
      return [];
    }
    return [
      {
        ...project,
        tickets,
        more: project.more + (project.tickets.length - tickets.length),
      },
    ];
  });
  return { ...report, projects };
}

function withWeekOldTickets(data: PmtAssigneeAgingReport): PmtAssigneeAgingReport {
  const projects = data.projects.flatMap((project) => {
    const tickets = project.tickets.filter((ticket) => ticket.days >= 7);
    if (tickets.length === 0) {
      return [];
    }
    const droppedYoung = project.tickets.length - tickets.length;
    return [{ ...project, tickets, more: droppedYoung > 0 ? 0 : project.more }];
  });
  return { ...data, projects };
}

function dayPhrase(ticket: PmtAssigneeAgingTicket): string {
  if (!ticket.column || ticket.column === "No status") {
    return agingDayLabel(ticket.days);
  }
  return `${agingDayLabel(ticket.days)} in ${ticket.column}`;
}

export function AgingTicketsPanel({
  teamsAccount,
  onTeamsSignIn,
}: AgingTicketsPanelProps): React.ReactNode {
  const signedIn = Boolean(teamsAccount?.signedIn);
  const [report, setReport] = useState<PmtAssigneeAgingReport | null>(null);
  const [assignee, setAssignee] = useState("");
  const [included, setIncluded] = useState<Set<string>>(new Set());
  const [note, setNote] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isSending, setIsSending] = useState(false);

  const applyReport = useCallback((data: PmtAssigneeAgingReport, refreshNote: boolean): void => {
    const aged = withWeekOldTickets(data);
    setReport(aged);
    setAssignee(data.person?.email ?? "");
    const keys = aged.projects.flatMap((project) =>
      project.tickets.map((ticket) => ticketKey(project.id, ticket)),
    );
    setIncluded(new Set(keys));
    if (refreshNote && data.person) {
      setNote(defaultAssigneeAgingNote(data.person.name));
    }
  }, []);

  const loadReport = useCallback(
    async (assigneeEmail: string, refreshNote: boolean): Promise<void> => {
      setIsLoading(true);
      setError(null);
      setStatus(null);
      try {
        const query = assigneeEmail ? `?assignee=${encodeURIComponent(assigneeEmail)}` : "";
        const response = await fetch(`/api/pmt/aging${query}`);
        const data = (await response.json()) as PmtAssigneeAgingReport & ApiErrorResponse;
        if (!response.ok) {
          setReport(null);
          setError(data.error ?? "Could not load aging tickets.");
          return;
        }
        applyReport(data, refreshNote);
      } catch {
        setReport(null);
        setError("Could not load aging tickets.");
      } finally {
        setIsLoading(false);
      }
    },
    [applyReport],
  );

  useEffect(() => {
    if (!signedIn) {
      return;
    }
    void loadReport("", true);
  }, [signedIn, loadReport]);

  const preview = report ? selectedReport(report, included) : null;
  const selectedCount = preview?.projects.reduce((sum, project) => sum + project.tickets.length, 0) ?? 0;

  const sendPreview = useCallback(async (): Promise<void> => {
    if (!preview || selectedCount === 0) {
      return;
    }
    setIsSending(true);
    setStatus(null);
    setError(null);
    try {
      const response = await fetch("/api/pmt/aging/send", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ note, report: preview }),
      });
      const data = (await response.json()) as { message?: string } & ApiErrorResponse;
      if (!response.ok) {
        setError(data.error ?? "Could not send the aging tickets.");
        return;
      }
      setStatus(data.message ?? `Aging tickets sent to ${preview.person?.name ?? "that person"}.`);
    } catch {
      setError("Could not send the aging tickets.");
    } finally {
      setIsSending(false);
    }
  }, [note, preview, selectedCount]);

  function toggleTicket(project: PmtAssigneeAgingProject, ticket: PmtAssigneeAgingTicket): void {
    const key = ticketKey(project.id, ticket);
    setIncluded((current) => {
      const next = new Set(current);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }

  return (
    <section>
      <div className="flex flex-wrap items-end justify-between gap-4">
        <div>
          <h2 className="text-lg font-medium tracking-tight text-zinc-900">Aging tickets</h2>
          <p className="mt-1 max-w-xl text-sm text-zinc-500">
            Five oldest open tickets on each project, kept in the same column for 7 days or more.
          </p>
        </div>
        {signedIn ? null : (
          <button type="button" onClick={onTeamsSignIn} className={primaryButtonClass}>
            Sign in to Teams
          </button>
        )}
      </div>

      {signedIn ? (
        <div className="mt-8 grid gap-4">
          <label className="block max-w-sm">
            <span className={labelClass}>Person</span>
            <select
              value={assignee}
              onChange={(event) => {
                const email = event.target.value;
                setAssignee(email);
                void loadReport(email, true);
              }}
              disabled={isLoading || (report?.people.length ?? 0) === 0}
              className={fieldClass}
            >
              {(report?.people ?? []).map((person) => (
                <option key={person.email} value={person.email}>
                  {person.name}
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className={labelClass}>Message</span>
            <textarea
              value={note}
              onChange={(event) => {
                setNote(event.target.value);
              }}
              rows={4}
              className={fieldClass}
            />
          </label>
        </div>
      ) : (
        <p className="mt-8 text-sm text-zinc-500">Sign in to choose a person.</p>
      )}

      <div className="mt-8 border-t border-zinc-200 pt-6">
        <p className="text-xs font-medium text-zinc-400">Preview</p>
        <div className="mt-4">
          {isLoading ? (
            <p className="text-sm text-zinc-500">Loading…</p>
          ) : preview?.person ? (
            <div>
              {note.trim() ? (
                <p className="whitespace-pre-wrap text-sm leading-6 text-zinc-800">
                  {noteSegments(note.trim()).map((segment, index) =>
                    segment.mention ? (
                      <span key={`${index}-${segment.text}`} className="font-medium">
                        {segment.text}
                      </span>
                    ) : (
                      <span key={`${index}-text`}>{segment.text}</span>
                    ),
                  )}
                </p>
              ) : null}
              {report && report.projects.length === 0 ? (
                <p className="mt-4 text-sm text-zinc-500">
                  No tickets have stayed in the same column for 7 days or more.
                </p>
              ) : (
                (report?.projects ?? []).map((project) => (
                  <div key={project.id} className="mt-6">
                    <h3 className="text-sm font-medium text-zinc-900">
                      {project.name}
                    </h3>
                    <ol className="mt-2 list-decimal space-y-2 pl-5 text-sm">
                      {project.tickets.map((ticket) => {
                        const key = ticketKey(project.id, ticket);
                        const checked = included.has(key);
                        return (
                          <li key={key} className={checked ? "text-zinc-800" : "text-zinc-400"}>
                            <label className="inline-flex items-start gap-2">
                              <input
                                type="checkbox"
                                checked={checked}
                                onChange={() => toggleTicket(project, ticket)}
                                className="mt-1 accent-zinc-900"
                              />
                              <span>
                                <a
                                  href={ticket.url}
                                  className="font-medium underline decoration-zinc-300 underline-offset-2"
                                >
                                  #{ticket.number} {ticket.name}
                                </a>{" "}
                                <span className="text-zinc-400">({dayPhrase(ticket)})</span>
                              </span>
                            </label>
                          </li>
                        );
                      })}
                    </ol>
                    {project.more > 0 ? (
                      <p className="mt-2 text-sm text-zinc-400">
                        And {project.more} more open {project.more === 1 ? "ticket" : "tickets"}.
                      </p>
                    ) : null}
                  </div>
                ))
              )}
              {report && report.omittedProjects > 0 ? (
                <p className="mt-4 text-sm text-zinc-400">
                  And {report.omittedProjects} other {report.omittedProjects === 1 ? "project" : "projects"}.
                </p>
              ) : null}
              <p className="mt-6 text-xs text-zinc-400">
                Uncheck a ticket to leave it out of the message to {preview.person.name}.
              </p>
            </div>
          ) : (
            <p className="text-sm text-zinc-400">Choose a person.</p>
          )}
        </div>
      </div>

      {error ? <p className="mt-4 text-sm text-red-700">{error}</p> : null}
      {status ? <p className="mt-4 text-sm text-emerald-700">{status}</p> : null}

      <div className="mt-6 flex justify-end">
        <button
          type="button"
          onClick={() => {
            void sendPreview();
          }}
          disabled={!signedIn || !preview?.person || selectedCount === 0 || isSending}
          className={primaryButtonClass}
        >
          {isSending ? "Sending…" : `Send to ${preview?.person?.name ?? "person"}`}
        </button>
      </div>
    </section>
  );
}
