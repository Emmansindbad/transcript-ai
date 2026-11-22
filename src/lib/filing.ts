import "server-only";

import { createHash } from "node:crypto";
import type { MeetingActionItem, MeetingSummary, ParsedActionItem } from "@/types/meeting";
import { formatKsaDate, getKsaDateKey } from "@/lib/ksa-time";
import { getMeetingSeriesId } from "@/lib/meeting-key";
import { buildWorkspaceMeetingKey } from "@/lib/ticket-lane";
import type { TicketLane } from "@/types/meeting";

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
  lane: TicketLane = "development",
): string {
  const lanePart = lane === "sheets" ? "|sheets" : "";
  const payload = `${meetingSeriesId}|${ksaDay}|${normalizeTicketTitle(title)}${lanePart}`;
  return createHash("sha256").update(payload).digest("hex").slice(0, 32);
}

function sheetsRelatedTo(relatedTo: string, meetingSubject: string): string {
  if (!/^(none|n\/a|na|tbd|-)?$/i.test(relatedTo.trim())) {
    return relatedTo;
  }
  return meetingSubject.trim() || "None";
}

export function toMeetingActionItems(
  meeting: MeetingSummary,
  parsedItems: ParsedActionItem[],
  filedKeys: Set<string>,
  lane: TicketLane = "development",
): MeetingActionItem[] {
  const meetingDate = formatKsaDate(meeting.start);
  const meetingKey = buildWorkspaceMeetingKey(meeting, lane);
  const meetingSeriesId = getMeetingSeriesId(meeting);
  const ksaDay = getKsaDateKey(meeting.start);

  return parsedItems.map((item, index) => {
    const filingKey = buildFilingKey(item.title, meetingSeriesId, ksaDay, lane);
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
      devTag: lane === "sheets" ? "None" : item.devTag,
      relatedTo: lane === "sheets" ? sheetsRelatedTo(item.relatedTo, meeting.subject) : item.relatedTo,
      description: item.description,
      evidence: item.evidence,
      lane,
    };
  });
}
