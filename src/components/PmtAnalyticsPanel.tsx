"use client";

import { useCallback, useEffect, useState } from "react";
import {
  analyticsAdvice,
  defaultAnalyticsNote,
  noteSegments,
  VISIBLE_AGING_TICKETS,
  type AnalyticsRange,
} from "@/lib/analytics-brief";
import type {
  MeetingChatOption,
  PmtAnalyticsReport,
  PmtBoardOption,
  TeamsAccountStatus,
} from "@/types/meeting";

interface PmtAnalyticsPanelProps {
  teamsAccount: TeamsAccountStatus | null;
  onTeamsSignIn: () => void;
}

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

const FIELD_CLASS =
  "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-200";

export function PmtAnalyticsPanel({
  teamsAccount,
  onTeamsSignIn,
}: PmtAnalyticsPanelProps): React.ReactNode {
  const signedIn = Boolean(teamsAccount?.signedIn);
  const [boards, setBoards] = useState<PmtBoardOption[]>([]);
  const [boardId, setBoardId] = useState("");
  const [range, setRange] = useState<AnalyticsRange>("week");
  const [chats, setChats] = useState<MeetingChatOption[]>([]);
  const [chatId, setChatId] = useState("");
  const [note, setNote] = useState("");
  const [noteTouched, setNoteTouched] = useState(false);
  const [report, setReport] = useState<PmtAnalyticsReport | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isLoadingBoards, setIsLoadingBoards] = useState(false);
  const [isLoadingReport, setIsLoadingReport] = useState(false);
  const [isSending, setIsSending] = useState(false);

  const loadChoices = useCallback(async (): Promise<void> => {
    setIsLoadingBoards(true);
    setError(null);
    try {
      const [boardResponse, chatResponse] = await Promise.all([
        fetch("/api/pmt/analytics"),
        fetch("/api/teams/meeting-chats"),
      ]);
      const boardData = (await boardResponse.json()) as { boards?: PmtBoardOption[] } & ApiErrorResponse;
      const chatData = (await chatResponse.json()) as { chats?: MeetingChatOption[] } & ApiErrorResponse;
      if (!boardResponse.ok) {
        setBoards([]);
        setError(boardData.error ?? "Could not load PMT boards.");
      } else {
        setBoards(boardData.boards ?? []);
      }
      if (!chatResponse.ok) {
        setChats([]);
        setError((current) => current ?? chatData.error ?? "Could not load meeting chats.");
      } else {
        setChats(chatData.chats ?? []);
      }
    } catch {
      setError("Could not load PMT boards.");
    } finally {
      setIsLoadingBoards(false);
    }
  }, []);

  useEffect(() => {
    if (!signedIn) {
      return;
    }
    void loadChoices();
  }, [signedIn, loadChoices]);

  const loadReport = useCallback(
    async (nextBoardId: string, nextRange: AnalyticsRange, refreshNote: boolean): Promise<void> => {
      setBoardId(nextBoardId);
      setReport(null);
      setStatus(null);
      if (!nextBoardId) {
        return;
      }
      setIsLoadingReport(true);
      setError(null);
      try {
        const response = await fetch(
          `/api/pmt/analytics?boardId=${encodeURIComponent(nextBoardId)}&range=${nextRange}`,
        );
        const data = (await response.json()) as PmtAnalyticsReport & ApiErrorResponse;
        if (!response.ok) {
          setError(data.error ?? "Could not load analytics.");
          return;
        }
        setReport(data);
        if (data.range !== nextRange) {
          setRange(data.range);
        }
        if (refreshNote) {
          setNote(defaultAnalyticsNote(data, data.range));
        }
      } catch {
        setError("Could not load analytics.");
      } finally {
        setIsLoadingReport(false);
      }
    },
    [],
  );

  const sendPreview = useCallback(async (): Promise<void> => {
    if (!report || !chatId) {
      return;
    }
    setIsSending(true);
    setStatus(null);
    setError(null);
    try {
      const response = await fetch("/api/pmt/analytics/send", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ chatId, note, report }),
      });
      const data = (await response.json()) as { message?: string } & ApiErrorResponse;
      if (!response.ok) {
        setError(data.error ?? "Could not send the preview.");
        return;
      }
      setStatus(data.message ?? "Analytics sent to the meeting chat.");
    } catch {
      setError("Could not send the preview.");
    } finally {
      setIsSending(false);
    }
  }, [chatId, note, report]);

  return (
    <section className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm lg:col-span-2 lg:p-5">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">PMT Analytics</h2>
          <p className="mt-1 text-sm text-slate-500">
            Where the work sits, then the items that have been waiting too long.
          </p>
        </div>
        {signedIn ? (
          <p className="text-xs text-slate-500">
            Boards for {teamsAccount?.displayName ?? teamsAccount?.username}
          </p>
        ) : (
          <button
            type="button"
            onClick={onTeamsSignIn}
            className="rounded-md bg-blue-600 px-3 py-2 text-sm font-medium text-white hover:bg-blue-700"
          >
            Sign in to Teams
          </button>
        )}
      </div>

      {signedIn ? (
        <div className="mt-4 grid gap-3 lg:grid-cols-2">
          <label className="block">
            <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
              Board
            </span>
            <select
              value={boardId}
              onChange={(event) => {
                setNoteTouched(false);
                void loadReport(event.target.value, range, true);
              }}
              disabled={isLoadingBoards || isLoadingReport}
              className={FIELD_CLASS}
            >
              <option value="">{isLoadingBoards ? "Loading boards…" : "Choose a board"}</option>
              {boards.map((board) => (
                <option key={board.id} value={board.id}>
                  {board.name}
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
              To
            </span>
            <select
              value={chatId}
              onChange={(event) => setChatId(event.target.value)}
              disabled={isLoadingBoards}
              className={FIELD_CLASS}
            >
              <option value="">{chats.length === 0 ? "No meeting chats" : "Choose a meeting chat"}</option>
              {chats.map((chat) => (
                <option key={chat.id} value={chat.id}>
                  {chat.topic}
                </option>
              ))}
            </select>
          </label>
          <div className="lg:col-span-2">
            <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
              Period
            </span>
            <div className="inline-flex rounded-md border border-slate-300 bg-white p-0.5">
              {(["week", "month"] as const).map((option) => (
                <button
                  key={option}
                  type="button"
                  onClick={() => {
                    setRange(option);
                    if (boardId) {
                      void loadReport(boardId, option, !noteTouched);
                    }
                  }}
                  className={`rounded px-3 py-1.5 text-sm ${
                    range === option ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-100"
                  }`}
                >
                  {option === "week" ? "1 week" : "1 month"}
                </button>
              ))}
            </div>
          </div>
          <label className="block lg:col-span-2">
            <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
              Message
            </span>
            <textarea
              value={note}
              onChange={(event) => {
                setNoteTouched(true);
                setNote(event.target.value);
              }}
              rows={7}
              placeholder="The weekly note appears here after you choose a board."
              className={FIELD_CLASS}
            />
          </label>
        </div>
      ) : (
        <p className="mt-4 text-sm text-slate-600">
          Sign in to load your PMT boards and the meeting chats you can send to.
        </p>
      )}

      <div className="mt-4 overflow-hidden rounded-xl border border-slate-200 bg-slate-50">
        <div className="border-b border-slate-200 bg-white px-4 py-2 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
          Preview
        </div>
        <div className="px-4 py-4">
          {isLoadingReport ? (
            <p className="text-sm text-slate-500">Loading analytics…</p>
          ) : report ? (
            <div className="mx-auto max-w-3xl rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              {note.trim() ? (
                <p className="whitespace-pre-wrap text-sm leading-6 text-slate-800">
                  {noteSegments(note.trim()).map((segment, index) =>
                    segment.mention ? (
                      <span key={`${index}-${segment.text}`} className="font-medium text-blue-700">
                        {segment.text}
                      </span>
                    ) : (
                      <span key={`${index}-text`}>{segment.text}</span>
                    ),
                  )}
                </p>
              ) : null}
              <h3 className="mt-4 text-lg font-semibold tracking-tight text-slate-900">{report.board.name}</h3>
              <p className="text-sm text-slate-500">
                {report.fromLabel} – {report.toLabel}
              </p>
              <div className="mt-4 grid grid-cols-2 gap-2 sm:grid-cols-4">
                {(
                  [
                    ["Open", report.open, false],
                    ["Closed", report.closed, false],
                    ["Overdue", report.overdue, report.overdue > 0],
                    ["Unassigned", report.unassigned, report.unassigned > 0],
                  ] as const
                ).map(([label, value, attention]) => (
                  <div key={label} className="rounded-lg border border-slate-200 px-3 py-2">
                    <p className="text-[11px] font-medium uppercase tracking-wide text-slate-500">{label}</p>
                    <p className={`mt-1 text-2xl font-semibold tabular-nums ${attention ? "text-rose-700" : "text-slate-900"}`}>
                      {value}
                    </p>
                  </div>
                ))}
              </div>
              <p className="mt-4 text-sm leading-6 text-slate-700">{analyticsAdvice(report)}</p>
              <h4 className="mt-5 text-sm font-semibold text-slate-900">{report.chartTitle}</h4>
              <img
                alt={report.chartTitle}
                src={`data:image/png;base64,${report.chartPng}`}
                className="mt-2 w-full bg-white"
              />
              <h4 className="mt-5 text-sm font-semibold text-slate-900">Aging tickets to act on</h4>
              {report.hiddenAging.length === 0 ? (
                <p className="mt-2 text-sm text-slate-500">
                  Every aging ticket is already named on the PMT page.
                </p>
              ) : (
                <ul className="mt-2 divide-y divide-slate-100">
                  {report.hiddenAging.slice(0, VISIBLE_AGING_TICKETS).map((ticket) => (
                    <li key={`${ticket.number}-${ticket.name}`} className="py-2 text-sm">
                      <a href={ticket.url} className="font-medium text-slate-900 hover:text-blue-700">
                        #{ticket.number} {ticket.name}
                      </a>
                      <p className="text-slate-500">
                        {ticket.column} ·{" "}
                        <span className={ticket.days >= 15 ? "font-medium text-rose-700" : ""}>
                          {ticket.days} days
                        </span>{" "}
                        · {ticket.owner}
                        {ticket.due ? ` · Due ${ticket.due}` : ""}
                      </p>
                    </li>
                  ))}
                </ul>
              )}
              {report.hiddenAging.length > VISIBLE_AGING_TICKETS ? (
                <p className="mt-2 text-sm text-slate-500">
                  And {report.hiddenAging.length - VISIBLE_AGING_TICKETS} more on the board.
                </p>
              ) : null}
              <a href={report.board.url} className="mt-4 inline-block text-sm text-blue-700 hover:underline">
                Open analytics
              </a>
            </div>
          ) : (
            <p className="text-sm text-slate-500">Choose a board to load the chart and the aging tickets.</p>
          )}
        </div>
      </div>

      {error ? <p className="mt-3 text-sm text-rose-700">{error}</p> : null}
      {status ? <p className="mt-3 text-sm text-emerald-700">{status}</p> : null}

      <div className="mt-4 flex justify-end">
        <button
          type="button"
          onClick={() => {
            void sendPreview();
          }}
          disabled={!signedIn || !report || !chatId || isSending}
          className="rounded-md bg-emerald-600 px-4 py-2 text-sm font-medium text-white hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300"
        >
          {isSending ? "Sending…" : "Send to chat"}
        </button>
      </div>
    </section>
  );
}
