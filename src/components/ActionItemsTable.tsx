"use client";

import { useState } from "react";
import { DEV_TAGS, normalizeDevTag } from "@/lib/dev-tags";
import type { MeetingActionItem, TeamsAccountStatus } from "@/types/meeting";

interface ActionItemsTableProps {
  actionItems: MeetingActionItem[];
  onItemsChange: (items: MeetingActionItem[]) => void;
  onSendEmail: () => void;
  onSendTeams: (includeIris: boolean) => void;
  onSendPmt: () => void;
  onTeamsSignIn: () => void;
  onTeamsSignOut: () => void;
  isSending: boolean;
  isSendingTeams: boolean;
  sendSuccess: boolean;
  sentRecipients?: string[];
  emailRecipient: string;
  teamsAccount: TeamsAccountStatus | null;
  teamsSummary?: string | null;
  isSendingPmt: boolean;
  pmtSummary?: string | null;
  pmtLinks?: Array<{ number: number; url: string; created: boolean }>;
}

const INPUT_CLASS =
  "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-200";

function updateItem(
  items: MeetingActionItem[],
  id: string,
  patch: Partial<MeetingActionItem>,
): MeetingActionItem[] {
  return items.map((item) => (item.id === id ? { ...item, ...patch } : item));
}

export function ActionItemsTable({
  actionItems,
  onItemsChange,
  onSendEmail,
  onSendTeams,
  onSendPmt,
  onTeamsSignIn,
  onTeamsSignOut,
  isSending,
  isSendingTeams,
  sendSuccess,
  sentRecipients = [],
  emailRecipient,
  teamsAccount,
  teamsSummary,
  isSendingPmt,
  pmtSummary,
  pmtLinks = [],
}: ActionItemsTableProps): React.ReactNode {
  const [includeIris, setIncludeIris] = useState(false);
  const selectedCount = actionItems.filter((item) => item.include).length;

  if (actionItems.length === 0) {
    return (
      <div className="border-b border-slate-200 bg-slate-50 px-5 py-4">
        <h3 className="text-sm font-semibold text-slate-900">Action Items</h3>
        <p className="mt-1 text-sm text-slate-500">
          No action items were identified for this meeting.
        </p>
      </div>
    );
  }

  return (
    <div className="border-b border-slate-200 bg-slate-50 px-5 py-4">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h3 className="text-sm font-semibold text-slate-900">Action Items</h3>
          <p className="mt-1 text-xs text-slate-500">
            {selectedCount} of {actionItems.length} selected · email{" "}
            {emailRecipient}. Edit before sending. Uncheck items to skip.
          </p>
          <p className="mt-1 text-xs text-slate-500">
            {teamsAccount?.signedIn
              ? `Teams chats send as ${teamsAccount.displayName ?? teamsAccount.username}.`
              : "Sign in once to send each ticket as a Teams chat from your account. Entra needs a Web redirect URI of http://localhost:3000/api/auth/teams/callback."}
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {sendSuccess && sentRecipients.length > 0 ? (
            <span className="rounded-lg border border-green-300 bg-green-50 px-3 py-2 text-sm font-medium text-green-800">
              Sent to {sentRecipients[0]}
            </span>
          ) : null}
          {teamsSummary ? (
            <span className="rounded-lg border border-green-300 bg-green-50 px-3 py-2 text-sm font-medium text-green-800">
              {teamsSummary}
            </span>
          ) : null}
          {pmtSummary ? (
            <div className="rounded-lg border border-green-300 bg-green-50 px-3 py-2 text-sm font-medium text-green-800">
              <p>{pmtSummary}</p>
              {pmtLinks.length > 0 ? (
                <p className="mt-1 flex flex-wrap gap-x-3 gap-y-1">
                  {pmtLinks.map((link) => (
                    <a
                      key={link.url}
                      href={link.url}
                      target="_blank"
                      rel="noreferrer"
                      className="underline"
                    >
                      {link.created ? "PMT" : "Already"} #{link.number}
                    </a>
                  ))}
                </p>
              ) : null}
            </div>
          ) : null}
          {teamsAccount?.signedIn ? (
            <button
              type="button"
              onClick={onTeamsSignOut}
              className="rounded-lg px-2 py-2 text-sm font-medium text-slate-500 hover:text-slate-800"
            >
              Sign out
            </button>
          ) : null}
          {teamsAccount?.signedIn ? (
            <div className="flex flex-wrap items-center gap-2">
              <label
                title="Add iris.ebale@sindbad.tech so the chat is a group of three"
                className="flex cursor-pointer items-center gap-2 rounded-lg border border-sky-200 bg-white px-3 py-2 text-sm font-medium text-sky-800 shadow-sm"
              >
                <input
                  type="checkbox"
                  checked={includeIris}
                  disabled={isSendingTeams}
                  onChange={(event) => setIncludeIris(event.target.checked)}
                  className="h-4 w-4 rounded border-slate-300 text-sky-700 focus:ring-sky-500 disabled:cursor-not-allowed"
                />
                include Iris
              </label>
              <button
                type="button"
                onClick={() => onSendTeams(includeIris)}
                disabled={isSendingTeams || selectedCount === 0}
                className="rounded-lg border border-sky-200 bg-white px-4 py-2 text-sm font-medium text-sky-800 shadow-sm transition-colors hover:bg-sky-50 disabled:cursor-not-allowed disabled:text-slate-400"
              >
                {isSendingTeams ? "Sending to Teams…" : "Send to Assignee"}
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={onTeamsSignIn}
              className="rounded-lg border border-sky-200 bg-white px-4 py-2 text-sm font-medium text-sky-800 shadow-sm transition-colors hover:bg-sky-50"
            >
              Sign in to Teams
            </button>
          )}
          <button
            type="button"
            onClick={onSendPmt}
            disabled={isSendingPmt || selectedCount === 0}
            className="rounded-lg bg-emerald-700 px-4 py-2 text-sm font-medium text-white shadow-sm transition-colors hover:bg-emerald-800 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
          >
            {isSendingPmt ? "Creating PMT tickets…" : "Create Ticket to PMT"}
          </button>
          <button
            type="button"
            onClick={onSendEmail}
            disabled={isSending || sendSuccess || selectedCount === 0}
            className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-medium text-white shadow-sm transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
          >
            {isSending ? "Sending…" : sendSuccess ? "Email Sent" : "Send via Email"}
          </button>
        </div>
      </div>

      <div className="space-y-3">
        {actionItems.map((item, index) => (
          <article
            key={item.id}
            className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm"
          >
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <label className="flex items-center gap-2 text-xs font-semibold text-indigo-600">
                <input
                  type="checkbox"
                  checked={item.include}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        include: event.target.checked,
                      }),
                    )
                  }
                  className="h-4 w-4 rounded border-slate-300 text-indigo-600"
                />
                Task #{index + 1} · {item.meetingTitle} · {item.meetingDate}
              </label>
              {item.alreadyFiled ? (
                <span className="rounded-full bg-amber-100 px-2 py-0.5 text-[11px] font-medium text-amber-800">
                  Already filed
                </span>
              ) : null}
              {item.pmtUrl && item.pmtNumber ? (
                <a
                  href={item.pmtUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="rounded-full bg-emerald-100 px-2 py-0.5 text-[11px] font-medium text-emerald-800"
                >
                  PMT #{item.pmtNumber}
                </a>
              ) : null}
            </div>

            <label className="block">
              <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                Title
              </span>
              <input
                value={item.title}
                onChange={(event) =>
                  onItemsChange(
                    updateItem(actionItems, item.id, { title: event.target.value }),
                  )
                }
                className={INPUT_CLASS}
              />
            </label>

            <div className="mt-3 grid grid-cols-2 gap-3">
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Type
                </span>
                <select
                  value={item.type}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        type: event.target.value === "Bug" ? "Bug" : "Task",
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                >
                  <option value="Task">Task</option>
                  <option value="Bug">Bug</option>
                </select>
              </label>
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Status
                </span>
                <input
                  value={item.status}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        status: event.target.value,
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                />
              </label>
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Priority
                </span>
                <select
                  value={item.priority}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        priority:
                          event.target.value === "High" ||
                          event.target.value === "Low"
                            ? event.target.value
                            : "Medium",
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                >
                  <option value="High">High</option>
                  <option value="Medium">Medium</option>
                  <option value="Low">Low</option>
                </select>
              </label>
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Due
                </span>
                <input
                  value={item.deadline}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        deadline: event.target.value,
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                />
              </label>
            </div>

            <label className="mt-3 block">
              <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                Assignee
              </span>
              <input
                value={item.pointPerson}
                onChange={(event) =>
                  onItemsChange(
                    updateItem(actionItems, item.id, {
                      pointPerson: event.target.value,
                    }),
                  )
                }
                className={INPUT_CLASS}
              />
              {item.assigneeUpn ? (
                <p className="mt-1 text-xs text-slate-500">{item.assigneeUpn}</p>
              ) : null}
            </label>

            <div className="mt-3 grid grid-cols-2 gap-3">
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Dev Tag
                </span>
                <select
                  value={normalizeDevTag(item.devTag)}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        devTag: normalizeDevTag(event.target.value),
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                >
                  {DEV_TAGS.map((tag) => (
                    <option key={tag} value={tag}>
                      {tag}
                    </option>
                  ))}
                </select>
              </label>
              <label className="block">
                <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Related To
                </span>
                <input
                  value={item.relatedTo}
                  onChange={(event) =>
                    onItemsChange(
                      updateItem(actionItems, item.id, {
                        relatedTo: event.target.value,
                      }),
                    )
                  }
                  className={INPUT_CLASS}
                />
              </label>
            </div>

            <label className="mt-3 block">
              <span className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                Description
              </span>
              <textarea
                value={item.description}
                onChange={(event) =>
                  onItemsChange(
                    updateItem(actionItems, item.id, {
                      description: event.target.value,
                    }),
                  )
                }
                rows={5}
                className={INPUT_CLASS}
              />
            </label>

            {(item.evidence ?? []).length > 0 ? (
              <div className="mt-3 rounded-md bg-slate-50 px-3 py-2">
                <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  Evidence
                </p>
                <ul className="mt-1 space-y-1 text-xs text-slate-600">
                  {(item.evidence ?? []).map((quote, quoteIndex) => (
                    <li key={`${item.id}-ev-${quoteIndex}`}>
                      <span className="font-medium text-slate-700">
                        {quote.timestamp || "time n/a"}
                        {quote.speaker ? ` · ${quote.speaker}` : ""}
                      </span>
                      {`: “${quote.quote}”`}
                    </li>
                  ))}
                </ul>
              </div>
            ) : null}
          </article>
        ))}
      </div>
    </div>
  );
}
