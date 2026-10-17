"use client";

import { useEffect, useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import type { MeetingActionItem, PmtTicketResult, TeamsAccountStatus } from "@/types/meeting";
import { ActionItemsTable } from "@/components/ActionItemsTable";

interface AnalysisPanelProps {
  analysis: string | null;
  actionItems: MeetingActionItem[];
  onItemsChange: (items: MeetingActionItem[]) => void;
  modelUsed?: string | null;
  isLoading: boolean;
  streamPreview?: string;
  statusMessage?: string | null;
  hasTranscript: boolean;
  videoRequired?: boolean;
  onGenerate: () => void;
  onSendEmail: () => void;
  onSendTeams: (includeIris: boolean) => void;
  onSendPmt: () => void;
  onTeamsSignIn: () => void;
  onTeamsSignOut: () => void;
  isSendingEmail: boolean;
  isSendingTeams: boolean;
  emailSendSuccess: boolean;
  emailRecipients?: string[];
  emailRecipient: string;
  teamsAccount: TeamsAccountStatus | null;
  teamsSummary?: string | null;
  isSendingPmt: boolean;
  pmtSummary?: string | null;
  pmtItems?: PmtTicketResult[];
  includeVideo: boolean;
  onToggleIncludeVideo: (value: boolean) => void;
  videoUsed?: boolean;
  videoNotice?: string | null;
  restoredAt?: string | null;
}

function formatElapsed(seconds: number): string {
  const minutes = Math.floor(seconds / 60);
  const remainder = seconds % 60;
  return `${minutes}:${remainder.toString().padStart(2, "0")}`;
}

function useElapsedSeconds(): number {
  const [seconds, setSeconds] = useState(0);

  useEffect(() => {
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      setSeconds(Math.floor((Date.now() - startedAt) / 1000));
    }, 1000);

    return () => {
      window.clearInterval(timer);
    };
  }, []);

  return seconds;
}

function AnalysisSkeleton({
  includeVideo,
  statusMessage,
  streamPreview,
}: {
  includeVideo: boolean;
  statusMessage?: string | null;
  streamPreview?: string;
}): React.ReactNode {
  const elapsedSeconds = useElapsedSeconds();
  const expected = includeVideo
    ? "Video analysis often takes 2–5 minutes (upload, then Gemini)."
    : "Transcript analysis usually finishes in 30–90 seconds. If Pro is busy it walks down to Flash.";

  return (
    <div className="space-y-4 p-6">
      <div className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-sm text-blue-800">
        <p className="font-medium">
          Still working — {formatElapsed(elapsedSeconds)} elapsed
        </p>
        <p className="mt-1 text-xs text-blue-700">
          {statusMessage || expected}
        </p>
        {elapsedSeconds >= 180 ? (
          <p className="mt-1 text-xs text-blue-700">
            This is slow, but not stuck unless an error banner appears or the
            button returns to Generate Analysis.
          </p>
        ) : null}
      </div>
      {streamPreview ? (
        <pre className="max-h-48 overflow-auto whitespace-pre-wrap rounded-lg bg-slate-50 p-3 font-mono text-xs text-slate-600">
          {streamPreview.slice(-2500)}
        </pre>
      ) : (
        <div className="animate-pulse space-y-4">
          <div className="h-5 w-1/3 rounded bg-slate-200" />
          <div className="space-y-2">
            {Array.from({ length: 8 }).map((_, index) => (
              <div key={index} className="h-3 rounded bg-slate-100" />
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

export function AnalysisPanel({
  analysis,
  actionItems,
  onItemsChange,
  modelUsed,
  isLoading,
  streamPreview = "",
  statusMessage,
  hasTranscript,
  videoRequired = false,
  onGenerate,
  onSendEmail,
  onSendTeams,
  onSendPmt,
  onTeamsSignIn,
  onTeamsSignOut,
  isSendingEmail,
  isSendingTeams,
  emailSendSuccess,
  emailRecipients = [],
  emailRecipient,
  teamsAccount,
  teamsSummary,
  isSendingPmt,
  pmtSummary,
  pmtItems = [],
  includeVideo,
  onToggleIncludeVideo,
  videoUsed,
  videoNotice,
  restoredAt,
}: AnalysisPanelProps): React.ReactNode {
  const canAnalyze = hasTranscript || videoRequired;
  const showActionItems = !isLoading && analysis !== null;

  return (
    <section className="flex h-full min-h-0 flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
      <header className="flex flex-col gap-3 border-b border-slate-200 px-5 py-4">
        <div className="flex items-start justify-between gap-4">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              AI Analysis
            </h2>
            <p className="mt-1 text-sm text-slate-500">
              {modelUsed
                ? `Analyzed with ${modelUsed}${videoUsed ? " · recording video included" : ""}`
                : "Tries Gemini 3.8 Flash, then 3.6 Flash if that model is busy"}
            </p>
            {restoredAt ? (
              <p className="mt-1 text-xs text-emerald-700">
                Restored saved analysis from {new Date(restoredAt).toLocaleString()}
              </p>
            ) : null}
          </div>
          <button
            type="button"
            onClick={onGenerate}
            disabled={!canAnalyze || isLoading}
            className="shrink-0 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white shadow-sm transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
          >
            {isLoading
              ? includeVideo
                ? "Analyzing video…"
                : "Generating…"
              : analysis
                ? "Regenerate"
                : "Generate Analysis"}
          </button>
        </div>

        <label
          className={`flex items-start gap-2.5 rounded-lg border px-3 py-2.5 text-sm transition-colors ${
            includeVideo
              ? "border-indigo-200 bg-indigo-50"
              : "border-slate-200 bg-slate-50"
          } ${canAnalyze && !isLoading && !videoRequired ? "cursor-pointer" : "cursor-not-allowed opacity-60"}`}
        >
          <input
            type="checkbox"
            checked={includeVideo || videoRequired}
            disabled={!canAnalyze || isLoading || videoRequired}
            onChange={(event) => onToggleIncludeVideo(event.target.checked)}
            className="mt-0.5 h-4 w-4 shrink-0 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
          />
          <span>
            <span className="font-medium text-slate-800">
              Analyze recording video
            </span>
            <span className="mt-0.5 block text-xs text-slate-500">
              Triangulates transcript, chat, and on-screen content. Recordings
              over 80 MB are skipped automatically.
            </span>
          </span>
        </label>

        {videoNotice ? (
          <p className="rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-800">
            {videoNotice}
          </p>
        ) : null}
      </header>

      {showActionItems ? (
        <ActionItemsTable
          actionItems={actionItems}
          onItemsChange={onItemsChange}
          onSendEmail={onSendEmail}
          onSendTeams={onSendTeams}
          onSendPmt={onSendPmt}
          onTeamsSignIn={onTeamsSignIn}
          onTeamsSignOut={onTeamsSignOut}
          isSending={isSendingEmail}
          isSendingTeams={isSendingTeams}
          sendSuccess={emailSendSuccess}
          sentRecipients={emailRecipients}
          emailRecipient={emailRecipient}
          teamsAccount={teamsAccount}
          teamsSummary={teamsSummary}
          isSendingPmt={isSendingPmt}
          pmtSummary={pmtSummary}
          pmtLinks={pmtItems.map((item) => ({
            number: item.number,
            url: item.url,
            created: item.created,
          }))}
        />
      ) : null}

      <div className="flex-1 overflow-y-auto">
        {isLoading ? (
          <AnalysisSkeleton
            includeVideo={includeVideo || videoRequired}
            statusMessage={statusMessage}
            streamPreview={streamPreview}
          />
        ) : analysis ? (
          <article className="prose prose-slate max-w-none p-6 prose-headings:text-slate-900 prose-p:text-slate-700 prose-li:text-slate-700 prose-strong:text-slate-900">
            <ReactMarkdown remarkPlugins={[remarkGfm]}>
              {analysis}
            </ReactMarkdown>
          </article>
        ) : (
          <div className="flex h-full flex-col items-center justify-center gap-3 p-6 text-center text-sm text-slate-500">
            <p>
              {hasTranscript
                ? "Click Generate Analysis to build the action items table and full meeting insights."
                : "Select a meeting to load its transcript, then click Generate Analysis."}
            </p>
          </div>
        )}
      </div>
    </section>
  );
}
