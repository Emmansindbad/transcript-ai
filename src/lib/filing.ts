import "server-only";

import { createHash } from "node:crypto";
import type { MeetingActionItem, MeetingSummary, ParsedActionItem } from "@/types/meeting";
import { formatKsaDate, getKsaDateKey } from "@/lib/ksa-time";
import { buildMeetingKey, getMeetingSeriesId } from "@/lib/meeting-key";

export function normalizeTicketTitle(title: string): string {
  return title
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function buildFilingKey(
  title: string,
  meetingSeriesId: string,
  ksaDay: string,
): string {
  const payload = `${meetingSeriesId}|${ksaDay}|${normalizeTicketTitle(title)}`;
  return createHash("sha256").update(payload).digest("hex").slice(0, 32);
}

export function toMeetingActionItems(
  meeting: MeetingSummary,
  parsedItems: ParsedActionItem[],
  filedKeys: Set<string>,
): MeetingActionItem[] {
  const meetingDate = formatKsaDate(meeting.start);
  const meetingKey = buildMeetingKey(meeting);
  const meetingSeriesId = getMeetingSeriesId(meeting);
  const ksaDay = getKsaDateKey(meeting.start);

  return parsedItems.map((item, index) => {
    const filingKey = buildFilingKey(item.title, meetingSeriesId, ksaDay);
    const alreadyFiled = filedKeys.has(filingKey);

    return {
      id: `${filingKey}-${index}`,
      meetingDate,
      meetingTitle: meeting.subject,
      meetingKey,
      meetingSeriesId,
      ksaDay,
      filingKey,
      include: !alreadyFiled,
      alreadyFiled,
      title: item.title,
      type: item.type,
      status: item.status,
      priority: item.priority,
      pointPerson: item.pointPerson,
      assigneeUpn: item.assigneeUpn,
      assigneeDisplayName: item.assigneeDisplayName,
      task: item.task,
      deadline: item.deadline,
      devTag: item.devTag,
      relatedTo: item.relatedTo,
      description: item.description,
      evidence: item.evidence,
    };
  });
}
