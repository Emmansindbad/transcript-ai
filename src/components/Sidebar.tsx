"use client";

import { useMemo, useState } from "react";
import type { MeetingSummary, TenantUser } from "@/types/meeting";
import {
  formatKsaDateKeyLabel,
  formatKsaDateTime,
  getKsaDateKey,
  getKsaDateRangeBounds,
} from "@/lib/ksa-time";
import { fieldClass, labelClass, primaryButtonClass } from "@/components/ui-classes";

interface SidebarProps {
  users: TenantUser[];
  selectedUserId: string;
  onUserChange: (userId: string) => void;
  meetings: MeetingSummary[];
  selectedEventId: string | null;
  onSelect: (meeting: MeetingSummary) => void;
  onClearSelection?: () => void;
  onFetchData: (date: string) => Promise<void>;
  isLoading: boolean;
  isLoadingUsers: boolean;
  hasLoadedMeetings: boolean;
  showOrganizer: boolean;
  heading?: string;
  note?: string;
}

type SortOption = "date-desc" | "date-asc" | "name-asc" | "name-desc";

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
      className={`animate-spin text-zinc-700 ${className}`}
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
  selectedEventId,
  onSelect,
  onClearSelection,
  onFetchData,
  isLoading,
  isLoadingUsers,
  hasLoadedMeetings,
  showOrganizer,
  heading = "Meetings",
  note,
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

  const everyone = !isLoadingUsers && users.length > 0 && selectedUserId === "all";
  const needsDate = everyone && !selectedDate;

  async function handleFetchData(): Promise<void> {
    setIsBrowsePending(true);
    try {
      await onFetchData(selectedDate);
    } finally {
      setIsBrowsePending(false);
    }
  }

  return (
    <aside className="flex w-full shrink-0 flex-col border-b border-zinc-200 bg-zinc-50 lg:h-full lg:w-72 lg:border-r lg:border-b-0">
      <div className="space-y-3 border-b border-zinc-200 px-4 py-4">
        <div className="flex items-baseline justify-between gap-3">
          <h2 className="text-sm font-medium text-zinc-900">{heading}</h2>
          <p className="text-xs text-zinc-400">
            {selectedDate
              ? formatDateLabel(selectedDate)
              : everyone
                ? "Pick a date"
                : "Last 30 days"}
          </p>
        </div>
        {note ? <p className="text-xs leading-5 text-zinc-500">{note}</p> : null}

        <button
          type="button"
          onClick={() => {
            void handleFetchData();
          }}
          disabled={isBrowseLoading || isLoadingUsers || needsDate}
          className={`${primaryButtonClass} w-full gap-2`}
        >
          {isBrowseLoading ? <LoadingSpinner className="h-4 w-4 text-white" /> : null}
          Fetch data
        </button>
        {needsDate ? (
          <p className="text-xs leading-5 text-zinc-500">
            Choose a date to load every recorded meeting that day.
          </p>
        ) : null}

        <label className="block">
          <span className={labelClass}>Person</span>
          <select
            value={selectedUserId}
            onChange={(event) => onUserChange(event.target.value)}
            className={fieldClass}
            disabled={isLoadingUsers && users.length === 0}
          >
            {isLoadingUsers && users.length === 0 ? (
              <option value="all">Loading people…</option>
            ) : null}
            {!isLoadingUsers && users.length === 0 ? (
              <option value="all">No people found</option>
            ) : null}
            {!isLoadingUsers && users.length > 0 ? (
              <option value="all">Everyone</option>
            ) : null}
            {users.map((user) => (
              <option key={user.id} value={user.userPrincipalName}>
                {user.displayName}
              </option>
            ))}
          </select>
        </label>

        <label className="block">
          <span className={`${labelClass} flex items-center gap-2`}>
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
              className={`${fieldClass}${isBrowseLoading ? " pr-10" : ""}`}
              disabled={isBrowseLoading}
              aria-busy={isBrowseLoading}
            />
            {isBrowseLoading ? <FieldLoadingIndicator /> : null}
          </div>
        </label>

        <label className="block">
          <span className={`${labelClass} flex items-center gap-2`}>
            Meeting
            {isBrowseLoading ? <LoadingSpinner className="h-3.5 w-3.5" /> : null}
          </span>
          <div className="relative">
            <select
              value={selectedMeetingSubject}
              onChange={(event) => {
                handleMeetingSubjectChange(event.target.value);
              }}
              className={`${fieldClass}${isBrowseLoading ? " pr-10" : ""}`}
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
          <span className={labelClass}>Sort</span>
          <select
            value={sortBy}
            onChange={(event) => setSortBy(event.target.value as SortOption)}
            className={fieldClass}
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
            className="flex items-center gap-2 text-sm text-zinc-500"
            role="status"
            aria-live="polite"
          >
            <LoadingSpinner className="h-4 w-4 shrink-0" />
            <span>
              {everyone
                ? "Loading every recorded meeting for that day…"
                : "Loading meetings from Microsoft Graph…"}
            </span>
          </div>
        ) : null}
      </div>

      <div className="flex-1 overflow-y-auto">
        {isBrowseLoading ? (
          <div className="flex flex-col items-center justify-center gap-3 p-8 text-sm text-zinc-500">
            <LoadingSpinner className="h-8 w-8" />
            <p>
              {everyone
                ? "Fetching recorded meetings for the selected day…"
                : selectedDate
                  ? "Fetching recorded meetings for that day…"
                  : "Fetching meetings with transcripts…"}
            </p>
          </div>
        ) : !hasLoadedMeetings ? (
          <div className="p-4 text-sm text-zinc-500">
            {everyone
              ? "Choose a date, then fetch every recorded meeting that day."
              : "Fetch data to load the last 30 days, or pick a date first."}
          </div>
        ) : hasLoadedMeetings && meetings.length === 0 ? (
          <div className="p-4 text-sm text-zinc-500">
            {selectedDate
              ? "No recorded meetings on this day."
              : "No meetings with a transcript or recording in the last 30 days."}
          </div>
        ) : visibleMeetings.length === 0 ? (
          <div className="p-4 text-sm text-zinc-500">
            {selectedDate
              ? "No meetings on this day in the last fetch. Fetch data again for this date."
              : "Nothing matches these filters."}
          </div>
        ) : (
          <>
            <p className="border-b border-zinc-200 px-4 py-2 text-xs text-zinc-400">
              {visibleMeetings.length}{" "}
              {visibleMeetings.length === 1 ? "meeting" : "meetings"}
              {effectiveMeetingSubject ? ` · ${effectiveMeetingSubject}` : ""}
              {selectedDate ? ` · ${formatDateLabel(selectedDate)}` : ""}
            </p>
            <ul>
              {visibleMeetings.map((meeting) => {
                const isSelected = meeting.eventId === selectedEventId;
                const marks = [
                  meeting.hasTranscript ? "Transcript" : null,
                  meeting.hasRecording ? "Recording" : null,
                ].filter((mark): mark is string => mark !== null);

                return (
                  <li key={`${meeting.userId}-${meeting.eventId}`}>
                    <button
                      type="button"
                      onClick={() => onSelect(meeting)}
                      className={`w-full border-b px-4 py-3 text-left transition-colors ${
                        isSelected
                          ? "bg-zinc-900 text-white"
                          : "border-zinc-200 hover:bg-white"
                      }`}
                    >
                      <p className="text-xs text-zinc-400">
                        {formatMeetingDate(meeting.start)}
                      </p>
                      <p className="mt-0.5 line-clamp-2 text-sm font-medium">
                        {meeting.subject}
                      </p>
                      {marks.length > 0 || showOrganizer ? (
                        <p className={`mt-1 text-xs ${isSelected ? "text-zinc-400" : "text-zinc-500"}`}>
                          {showOrganizer ? meeting.userDisplayName : null}
                          {showOrganizer && marks.length > 0 ? " · " : null}
                          {marks.join(" · ")}
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
