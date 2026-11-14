"use client";

import { useState } from "react";
import { DEV_TAGS, normalizeDevTag } from "@/lib/dev-tags";
import type { MeetingActionItem, TeamsAccountStatus } from "@/types/meeting";
import { fieldClass, labelClass, primaryButtonClass, secondaryButtonClass } from "@/components/ui-classes";

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
      <div className="border-b border-zinc-200 px-5 py-4">
        <h3 className="text-sm font-medium text-zinc-900">Action items</h3>
        <p className="mt-1 text-sm text-zinc-500">None found in this meeting.</p>
      </div>
    );
  }

  return (
    <div className="border-b border-zinc-200 px-5 py-4">
      <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
        <div>
          <h3 className="text-sm font-medium text-zinc-900">Action items</h3>
          <p className="mt-1 text-xs text-zinc-500">
            {selectedCount} of {actionItems.length} selected · {emailRecipient}
          </p>
          <p className="mt-1 text-xs text-zinc-400">
            {teamsAccount?.signedIn
              ? `Teams sends as ${teamsAccount.displayName ?? teamsAccount.username}.`
              : "Sign in to send each ticket from your Teams account."}
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {sendSuccess && sentRecipients.length > 0 ? (
            <span className="text-sm text-emerald-700">Sent to {sentRecipients[0]}</span>
          ) : null}
          {teamsSummary ? (
            <span className="text-sm text-emerald-700">{teamsSummary}</span>
          ) : null}
          {pmtSummary ? (
            <div className="text-sm text-emerald-700">
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
              className="rounded-md px-2 py-2 text-sm text-zinc-500 hover:text-zinc-900"
            >
              Sign out
            </button>
          ) : null}
          {teamsAccount?.signedIn ? (
            <div className="flex flex-wrap items-center gap-2">
              <label
                title="Add iris.ebale@sindbad.tech so the chat is a group of three"
                className="flex cursor-pointer items-center gap-2 text-sm text-zinc-600"
              >
                <input
                  type="checkbox"
                  checked={includeIris}
                  disabled={isSendingTeams}
                  onChange={(event) => setIncludeIris(event.target.checked)}
                  className="h-4 w-4 accent-zinc-900 disabled:cursor-not-allowed"
                />
                include Iris
              </label>
              <button
                type="button"
                onClick={() => onSendTeams(includeIris)}
                disabled={isSendingTeams || selectedCount === 0}
                className={secondaryButtonClass}
              >
                {isSendingTeams ? "Sending to Teams…" : "Send to Assignee"}
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={onTeamsSignIn}
              className={secondaryButtonClass}
            >
              Sign in to Teams
            </button>
          )}
          <button
            type="button"
            onClick={onSendPmt}
            disabled={isSendingPmt || selectedCount === 0}
            className={primaryButtonClass}
          >
            {isSendingPmt ? "Creating PMT tickets…" : "Create Ticket to PMT"}
          </button>
          <button
            type="button"
            onClick={onSendEmail}
            disabled={isSending || sendSuccess || selectedCount === 0}
            className={secondaryButtonClass}
          >
            {isSending ? "Sending…" : sendSuccess ? "Email Sent" : "Send via Email"}
          </button>
        </div>
      </div>

      <div className="space-y-3">
        {actionItems.map((item, index) => (
          <article
            key={item.id}
            className="rounded-md border border-zinc-200 bg-white p-4"
          >
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <label className="flex items-center gap-2 text-xs text-zinc-500">
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
                  className="h-4 w-4 accent-zinc-900"
                />
                Task #{index + 1} · {item.meetingTitle} · {item.meetingDate}
              </label>
              {item.alreadyFiled ? (
                <span className="text-[11px] text-amber-700">Already filed</span>
              ) : null}
              {item.pmtUrl && item.pmtNumber ? (
                <a
                  href={item.pmtUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="text-[11px] text-emerald-700 underline decoration-emerald-200 underline-offset-2"
                >
                  PMT #{item.pmtNumber}
                </a>
              ) : null}
            </div>

            <label className="block">
              <span className={labelClass}>
                Title
              </span>
              <input
                value={item.title}
                onChange={(event) =>
                  onItemsChange(
                    updateItem(actionItems, item.id, { title: event.target.value }),
                  )
                }
                className={fieldClass}
              />
            </label>

            <div className="mt-3 grid grid-cols-2 gap-3">
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                >
                  <option value="Task">Task</option>
                  <option value="Bug">Bug</option>
                </select>
              </label>
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                />
              </label>
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                >
                  <option value="High">High</option>
                  <option value="Medium">Medium</option>
                  <option value="Low">Low</option>
                </select>
              </label>
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                />
              </label>
            </div>

            <label className="mt-3 block">
              <span className={labelClass}>
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
                className={fieldClass}
              />
              {item.assigneeUpn ? (
                <p className="mt-1 text-xs text-zinc-400">{item.assigneeUpn}</p>
              ) : null}
            </label>

            <div className="mt-3 grid grid-cols-2 gap-3">
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                >
                  {DEV_TAGS.map((tag) => (
                    <option key={tag} value={tag}>
                      {tag}
                    </option>
                  ))}
                </select>
              </label>
              <label className="block">
                <span className={labelClass}>
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
                  className={fieldClass}
                />
              </label>
            </div>

            <label className="mt-3 block">
              <span className={labelClass}>
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
                className={fieldClass}
              />
              <p className="mt-1 text-[11px] text-zinc-400">
                Chat or screenshot sources use [title](link) from Teams Copy link.
              </p>
            </label>

            {(item.evidence ?? []).length > 0 ? (
              <div className="mt-3 border-t border-zinc-100 pt-3">
                <p className="text-xs font-medium text-zinc-400">Evidence</p>
                <ul className="mt-1 space-y-1 text-xs text-zinc-600">
                  {(item.evidence ?? []).map((quote, quoteIndex) => (
                    <li key={`${item.id}-ev-${quoteIndex}`}>
                      <span className="font-medium text-zinc-800">
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
