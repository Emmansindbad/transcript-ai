"use client";

import { useMemo, useState } from "react";
import type { MeetingSummary, TenantUser } from "@/types/meeting";
import {
  formatKsaDateKeyLabel,
  formatKsaDateTime,
  getKsaDateKey,
  getKsaDateRangeBounds,
} from "@/lib/ksa-time";

interface SidebarProps {
  users: TenantUser[];
  selectedUserId: string;
  onUserChange: (userId: string) => void;
  meetings: MeetingSummary[];
  selectedJoinUrl: string | null;
  onSelect: (meeting: MeetingSummary) => void;
  onClearSelection?: () => void;
  onFetchData: () => Promise<void>;
  isLoading: boolean;
  isLoadingUsers: boolean;
  hasLoadedMeetings: boolean;
  showOrganizer: boolean;
}

type SortOption = "date-desc" | "date-asc" | "name-asc" | "name-desc";

const SELECT_CLASS =
  "w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-200";

function getDateKey(isoDate: string): string {
  return getKsaDateKey(isoDate);
}

function formatMeetingDate(isoDate: string): string {
  return formatKsaDateTime(isoDate);
}

function getDateRangeBounds(): { min: string; max: string } {
  return getKsaDateRangeBounds(30);
}

function formatDateLabel(dateKey: string): string {
  return formatKsaDateKeyLabel(dateKey);
}

function sortMeetings(
  items: MeetingSummary[],
  sortBy: SortOption,
): MeetingSummary[] {
  const sorted = [...items];

  sorted.sort((left, right) => {
    switch (sortBy) {
      case "date-asc":
        return new Date(left.start).getTime() - new Date(right.start).getTime();
      case "name-asc":
        return left.subject.localeCompare(right.subject);
      case "name-desc":
        return right.subject.localeCompare(left.subject);
      case "date-desc":
      default:
        return new Date(right.start).getTime() - new Date(left.start).getTime();
    }
  });

  return sorted;
}

function getUniqueMeetingNames(items: MeetingSummary[]): string[] {
  const latestBySubject = new Map<string, number>();

  for (const meeting of items) {
    const existing = latestBySubject.get(meeting.subject);
    const startTime = new Date(meeting.start).getTime();
    if (existing === undefined || startTime > existing) {
      latestBySubject.set(meeting.subject, startTime);
    }
  }

  return Array.from(latestBySubject.entries())
    .sort(([, leftTime], [, rightTime]) => rightTime - leftTime)
    .map(([subject]) => subject);
}

function LoadingSpinner({
  className = "h-5 w-5",
}: {
  className?: string;
}): React.ReactNode {
  return (
    <svg
      className={`animate-spin text-blue-600 ${className}`}
      xmlns="http://www.w3.org/2000/svg"
      fill="none"
      viewBox="0 0 24 24"
      aria-hidden="true"
    >
      <circle
        className="opacity-25"
        cx="12"
        cy="12"
        r="10"
        stroke="currentColor"
        strokeWidth="4"
      />
      <path
        className="opacity-75"
        fill="currentColor"
        d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
      />
    </svg>
  );
}

function FieldLoadingIndicator(): React.ReactNode {
  return (
    <span className="pointer-events-none absolute inset-y-0 right-3 flex items-center">
      <LoadingSpinner className="h-4 w-4" />
    </span>
  );
}

export function Sidebar({
  users,
  selectedUserId,
  onUserChange,
  meetings,
  selectedJoinUrl,
  onSelect,
  onClearSelection,
  onFetchData,
  isLoading,
  isLoadingUsers,
  hasLoadedMeetings,
  showOrganizer,
}: SidebarProps): React.ReactNode {
  const [selectedDate, setSelectedDate] = useState("");
  const [selectedMeetingSubject, setSelectedMeetingSubject] = useState("");
  const [sortBy, setSortBy] = useState<SortOption>("date-desc");
  const [isBrowsePending, setIsBrowsePending] = useState(false);
  const dateBounds = getDateRangeBounds();
  const isBrowseLoading = isLoading || isBrowsePending;

  const meetingsForSelectedDate = useMemo(() => {
    if (!selectedDate) {
      return meetings;
    }

    return meetings.filter(
      (meeting) => getDateKey(meeting.start) === selectedDate,
    );
  }, [meetings, selectedDate]);

  const meetingNameOptions = useMemo(
    () => getUniqueMeetingNames(meetingsForSelectedDate),
    [meetingsForSelectedDate],
  );

  const effectiveMeetingSubject = useMemo(() => {
    if (!selectedMeetingSubject) {
      return "";
    }

    return meetingNameOptions.includes(selectedMeetingSubject)
      ? selectedMeetingSubject
      : "";
  }, [selectedMeetingSubject, meetingNameOptions]);

  const visibleMeetings = useMemo(() => {
    if (!hasLoadedMeetings && !selectedDate && !effectiveMeetingSubject) {
      return [];
    }

    let filtered = meetingsForSelectedDate;

    if (effectiveMeetingSubject) {
      filtered = filtered.filter(
        (meeting) => meeting.subject === effectiveMeetingSubject,
      );
    }

    return sortMeetings(filtered, sortBy);
  }, [
    hasLoadedMeetings,
    selectedDate,
    meetingsForSelectedDate,
    effectiveMeetingSubject,
    sortBy,
  ]);

  function handleMeetingSubjectChange(subject: string): void {
    setSelectedMeetingSubject(subject);
    onClearSelection?.();
  }

  function handleDateChange(dateKey: string): void {
    setSelectedDate(dateKey);
    onClearSelection?.();
  }

  async function handleFetchData(): Promise<void> {
    setIsBrowsePending(true);
    try {
      await onFetchData();
    } finally {
      setIsBrowsePending(false);
    }
  }

  return (
    <aside className="flex h-full w-full flex-col border-r border-slate-200 bg-slate-50 lg:w-80 lg:shrink-0">
      <div className="space-y-3 border-b border-slate-200 px-4 py-4">
        <div>
          <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-500">
            Meetings
          </h2>
          <p className="mt-1 text-xs text-slate-400">
            Fetch meetings from the last 30 days, then filter by date or name
          </p>
        </div>

        <button
          type="button"
          onClick={() => {
            void handleFetchData();
          }}
          disabled={isBrowseLoading || isLoadingUsers || selectedUserId === "all"}
          className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white shadow-sm transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:text-slate-500"
        >
          {isBrowseLoading ? <LoadingSpinner className="h-4 w-4 text-white" /> : null}
          Fetch Data
        </button>

        <label className="block">
          <span className="mb-1.5 block text-xs font-medium text-slate-600">
            Person
          </span>
          <select
            value={selectedUserId}
            onChange={(event) => onUserChange(event.target.value)}
            className={SELECT_CLASS}
            disabled={isLoadingUsers && users.length === 0}
          >
            {isLoadingUsers && users.length === 0 ? (
              <option value="all">Loading people…</option>
            ) : null}
            {!isLoadingUsers && users.length === 0 ? (
              <option value="all">No people found</option>
            ) : null}
            {users.map((user) => (
              <option key={user.id} value={user.userPrincipalName}>
                {user.displayName}
              </option>
            ))}
          </select>
        </label>

        <label className="block">
          <span className="mb-1.5 flex items-center gap-2 text-xs font-medium text-slate-600">
            Date
            {isBrowseLoading ? <LoadingSpinner className="h-3.5 w-3.5" /> : null}
          </span>
          <div className="relative">
            <input
              type="date"
              value={selectedDate}
              min={dateBounds.min}
              max={dateBounds.max}
              onChange={(event) => {
                handleDateChange(event.target.value);
              }}
              className={`${SELECT_CLASS}${isBrowseLoading ? " pr-10" : ""}`}
              disabled={isBrowseLoading}
              aria-busy={isBrowseLoading}
            />
            {isBrowseLoading ? <FieldLoadingIndicator /> : null}
          </div>
        </label>

        <label className="block">
          <span className="mb-1.5 flex items-center gap-2 text-xs font-medium text-slate-600">
            Meeting name
            {isBrowseLoading ? <LoadingSpinner className="h-3.5 w-3.5" /> : null}
          </span>
          <div className="relative">
            <select
              value={selectedMeetingSubject}
              onChange={(event) => {
                handleMeetingSubjectChange(event.target.value);
              }}
              className={`${SELECT_CLASS}${isBrowseLoading ? " pr-10" : ""}`}
              disabled={isBrowseLoading}
              aria-busy={isBrowseLoading}
            >
              <option value="">
                {isBrowseLoading
                  ? "Loading meetings…"
                  : hasLoadedMeetings
                    ? "Select a meeting"
                    : "Fetch data to list meetings"}
              </option>
              {meetingNameOptions.map((subject) => (
                <option key={subject} value={subject}>
                  {subject}
                </option>
              ))}
            </select>
            {isBrowseLoading ? <FieldLoadingIndicator /> : null}
          </div>
        </label>

        <label className="block">
          <span className="mb-1.5 block text-xs font-medium text-slate-600">
            Sort by
          </span>
          <select
            value={sortBy}
            onChange={(event) => setSortBy(event.target.value as SortOption)}
            className={SELECT_CLASS}
            disabled={isBrowseLoading || !hasLoadedMeetings}
          >
            <option value="date-desc">Date (newest first)</option>
            <option value="date-asc">Date (oldest first)</option>
            <option value="name-asc">Name (A–Z)</option>
            <option value="name-desc">Name (Z–A)</option>
          </select>
        </label>

        {isBrowseLoading ? (
          <div
            className="flex items-center gap-2 rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-sm text-blue-800"
            role="status"
            aria-live="polite"
          >
            <LoadingSpinner className="h-4 w-4 shrink-0" />
            <span>Loading meetings from Microsoft Graph…</span>
          </div>
        ) : null}
      </div>

      <div className="flex-1 overflow-y-auto">
        {isBrowseLoading ? (
          <div className="flex flex-col items-center justify-center gap-3 p-8 text-sm text-slate-500">
            <LoadingSpinner className="h-8 w-8" />
            <p>Fetching meetings with transcripts…</p>
          </div>
        ) : !hasLoadedMeetings ? (
          <div className="p-4 text-sm text-slate-500">
            Click Fetch Data to load meetings from the last 30 days.
          </div>
        ) : hasLoadedMeetings && meetings.length === 0 ? (
          <div className="p-4 text-sm text-slate-500">
            No meetings with transcripts found for the selected person in the
            last 30 days.
          </div>
        ) : visibleMeetings.length === 0 ? (
          <div className="p-4 text-sm text-slate-500">
            No meetings match the selected filters.
          </div>
        ) : (
          <>
            <p className="border-b border-slate-200 px-4 py-2 text-xs text-slate-500">
              {visibleMeetings.length}{" "}
              {visibleMeetings.length === 1 ? "meeting" : "meetings"}
              {effectiveMeetingSubject ? ` · ${effectiveMeetingSubject}` : ""}
              {selectedDate ? ` · ${formatDateLabel(selectedDate)}` : ""}
            </p>
            <ul className="space-y-2 p-3">
              {visibleMeetings.map((meeting) => {
                const isSelected = meeting.joinUrl === selectedJoinUrl;

                return (
                  <li key={`${meeting.userId}-${meeting.eventId}`}>
                    <button
                      type="button"
                      onClick={() => onSelect(meeting)}
                      className={`w-full rounded-lg border px-3 py-3 text-left transition-colors ${
                        isSelected
                          ? "border-blue-500 bg-blue-50 shadow-sm"
                          : "border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-100"
                      }`}
                    >
                      <p className="text-xs font-semibold uppercase tracking-wide text-blue-600">
                        {formatMeetingDate(meeting.start)}
                      </p>
                      <p className="mt-1 line-clamp-2 text-sm font-medium text-slate-900">
                        {meeting.subject}
                      </p>
                      <div className="mt-2 flex flex-wrap gap-1">
                        {meeting.hasTranscript ? (
                          <span className="rounded bg-blue-100 px-1.5 py-0.5 text-[10px] font-medium text-blue-800">
                            Transcript
                          </span>
                        ) : null}
                        {meeting.hasRecording ? (
                          <span className="rounded bg-indigo-100 px-1.5 py-0.5 text-[10px] font-medium text-indigo-800">
                            Recording
                          </span>
                        ) : null}
                      </div>
                      {showOrganizer ? (
                        <p className="mt-1 text-xs text-slate-500">
                          {meeting.userDisplayName}
                        </p>
                      ) : null}
                    </button>
                  </li>
                );
              })}
            </ul>
          </>
        )}
      </div>
    </aside>
  );
}
