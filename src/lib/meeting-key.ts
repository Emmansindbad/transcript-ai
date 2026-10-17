import type { MeetingSummary } from "@/types/meeting";
import { getKsaDateKey } from "@/lib/ksa-time";

export function getMeetingSeriesId(meeting: {
  joinUrl: string;
  seriesMasterId?: string;
}): string {
  const series = meeting.seriesMasterId?.trim();
  return series || meeting.joinUrl.trim();
}

export function buildMeetingKey(meeting: {
  joinUrl: string;
  start: string;
  seriesMasterId?: string;
  eventId?: string;
}): string {
  const day = getKsaDateKey(meeting.start);
  const series = getMeetingSeriesId(meeting);
  return `${series}|${day}|${meeting.eventId ?? ""}`;
}

export function isSameMeeting(
  left: MeetingSummary,
  right: MeetingSummary,
): boolean {
  return left.joinUrl === right.joinUrl && left.start === right.start;
}

export function meetingListCacheKey(userId: string, daysBack: number): string {
  return `meetings:${userId}:${daysBack}`;
}
