import type { PmtAnalyticsReport, PmtHiddenAgingTicket } from "@/types/meeting";

export type AnalyticsRange = "week" | "month";

export function topAgingOwners(tickets: PmtHiddenAgingTicket[], limit = 5): string[] {
  const names: string[] = [];
  for (const ticket of tickets) {
    const owner = ticket.owner.trim();
    if (!owner || /^unassigned$/i.test(owner) || /^someone$/i.test(owner)) {
      continue;
    }
    if (names.some((name) => name.toLowerCase() === owner.toLowerCase())) {
      continue;
    }
    names.push(owner);
    if (names.length >= limit) {
      break;
    }
  }
  return names;
}

function bucketCount(report: Pick<PmtAnalyticsReport, "aging">, needle: string): number {
  return report.aging.find((bucket) => bucket.label.toLowerCase().includes(needle))?.count ?? 0;
}

export function analyticsAdvice(report: Pick<PmtAnalyticsReport, "aging" | "overdue" | "unassigned">): string {
  const over30 = bucketCount(report, "over 30");
  const waiting = bucketCount(report, "15");
  const sentences: string[] = [];
  if (over30 > 0) {
    sentences.push(
      over30 === 1
        ? "1 item has been in the same column for more than 30 days. Start there."
        : `${over30} items have been in the same column for more than 30 days. Start there.`,
    );
  } else if (waiting > 0) {
    sentences.push(
      waiting === 1
        ? "1 item has been waiting 15 to 30 days."
        : `${waiting} items have been waiting 15 to 30 days.`,
    );
  } else {
    sentences.push("No open item has been sitting in one column for more than two weeks.");
  }
  if (report.overdue > 0) {
    sentences.push(report.overdue === 1 ? "1 item is past its due date." : `${report.overdue} items are past their due date.`);
  }
  if (report.unassigned > 0) {
    sentences.push(
      report.unassigned === 1
        ? "1 open item has no owner."
        : `${report.unassigned} open items have no owner.`,
    );
  }
  return sentences.join(" ");
}

export function defaultAssigneeAgingNote(personName: string): string {
  const name = personName.replace(/\s+/g, " ").trim();
  return [
    `@${name}`,
    "These are your oldest open tickets, across every project, from the start until today.",
    "Please check them and keep the status current. If the work is done, move it to Ready for QA.",
  ].join("\n");
}

export function agingDayLabel(days: number): string {
  return days === 1 ? "1 day" : `${days} days`;
}

export function defaultAnalyticsNote(
  report: Pick<PmtAnalyticsReport, "board" | "fromLabel" | "toLabel" | "hiddenAging">,
  range: AnalyticsRange,
): string {
  const period = range === "month" ? "month" : "week";
  const owners = topAgingOwners(report.hiddenAging);
  const callout =
    owners.length > 0
      ? `${owners.map((name) => `@${name.replace(/\s+/g, " ").trim()}`).join("\n")}\nPlease check the tickets assigned to you.`
      : "Please check the aging tickets and put an owner on any that are still unassigned.";
  return [
    `@Everyone — this is the analytics for the ${period} of ${report.board.name} (${report.fromLabel} – ${report.toLabel}).`,
    "Please check the tickets on your side and make sure they are up to date. If the work is done, move it to Ready for QA.",
    "There are aging items on this project. Please take action on them.",
    callout,
  ].join("\n\n");
}

export interface NoteSegment {
  text: string;
  mention: boolean;
}

export const VISIBLE_AGING_TICKETS = 8;

export function noteSegments(note: string): NoteSegment[] {
  const pattern = /@Everyone|@[^\n]+?(?=\s@|\s—|\n|$)/g;
  const segments: NoteSegment[] = [];
  let cursor = 0;
  for (const match of note.matchAll(pattern)) {
    const index = match.index ?? 0;
    if (index > cursor) {
      segments.push({ text: note.slice(cursor, index), mention: false });
    }
    segments.push({ text: match[0], mention: true });
    cursor = index + match[0].length;
  }
  if (cursor < note.length) {
    segments.push({ text: note.slice(cursor), mention: false });
  }
  return segments;
}
