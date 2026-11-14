"use client";

import { useEffect, useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import type { MeetingActionItem, PmtTicketResult, TeamsAccountStatus } from "@/types/meeting";
import { ActionItemsTable } from "@/components/ActionItemsTable";
import { primaryButtonClass } from "@/components/ui-classes";

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
      <div className="text-sm text-zinc-700">
        <p className="font-medium">Working · {formatElapsed(elapsedSeconds)}</p>
        <p className="mt-1 text-xs text-zinc-500">{statusMessage || expected}</p>
        {elapsedSeconds >= 180 ? (
          <p className="mt-1 text-xs text-zinc-500">
            Still running. An error banner appears if it stops.
          </p>
        ) : null}
      </div>
      {streamPreview ? (
        <pre className="max-h-48 overflow-auto whitespace-pre-wrap rounded-md bg-zinc-50 p-3 font-mono text-xs text-zinc-600">
          {streamPreview.slice(-2500)}
        </pre>
      ) : (
        <div className="animate-pulse space-y-4">
          <div className="h-5 w-1/3 rounded bg-zinc-200" />
          <div className="space-y-2">
            {Array.from({ length: 8 }).map((_, index) => (
              <div key={index} className="h-3 rounded bg-zinc-100" />
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
    <section className="flex h-full min-h-[24rem] flex-col overflow-hidden bg-white lg:min-h-0">
      <header className="flex flex-col gap-3 border-b border-zinc-200 px-5 py-4">
        <div className="flex items-start justify-between gap-4">
          <div>
            <h2 className="text-sm font-medium text-zinc-900">Analysis</h2>
            <p className="mt-1 text-sm text-zinc-500">
              {modelUsed
                ? `${modelUsed}${videoUsed ? " · video" : ""}`
                : "Gemini Flash"}
            </p>
            {restoredAt ? (
              <p className="mt-1 text-xs text-zinc-400">
                Saved {new Date(restoredAt).toLocaleString()}
              </p>
            ) : null}
          </div>
          <button
            type="button"
            onClick={onGenerate}
            disabled={!canAnalyze || isLoading}
            className={`${primaryButtonClass} shrink-0`}
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
          className={`flex items-start gap-2.5 text-sm ${
            canAnalyze && !isLoading && !videoRequired ? "cursor-pointer" : "cursor-not-allowed opacity-50"
          }`}
        >
          <input
            type="checkbox"
            checked={includeVideo || videoRequired}
            disabled={!canAnalyze || isLoading || videoRequired}
            onChange={(event) => onToggleIncludeVideo(event.target.checked)}
            className="mt-0.5 h-4 w-4 shrink-0 accent-zinc-900"
          />
          <span>
            <span className="font-medium text-zinc-800">Include the recording</span>
            <span className="mt-0.5 block text-xs text-zinc-500">
              Uses the video with the transcript and chat. Files over 150 MB are skipped.
            </span>
          </span>
        </label>

        {videoNotice ? (
          <p className="text-xs text-amber-800">{videoNotice}</p>
        ) : null}
      </header>

      <div className="analysis-scroll min-h-0 flex-1 overflow-y-scroll overscroll-contain">
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
        {isLoading ? (
          <AnalysisSkeleton
            includeVideo={includeVideo || videoRequired}
            statusMessage={statusMessage}
            streamPreview={streamPreview}
          />
        ) : analysis ? (
          <article className="prose prose-zinc max-w-none p-6 prose-headings:text-zinc-900 prose-p:text-zinc-700 prose-li:text-zinc-700 prose-strong:text-zinc-900">
            <ReactMarkdown remarkPlugins={[remarkGfm]}>
              {analysis}
            </ReactMarkdown>
          </article>
        ) : (
          <div className="flex h-full items-center justify-center p-6 text-center text-sm text-zinc-400">
            <p>
              {hasTranscript
                ? "Generate an analysis when you are ready."
                : "Select a meeting, then generate an analysis."}
            </p>
          </div>
        )}
      </div>
    </section>
  );
}
