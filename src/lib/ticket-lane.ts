import { buildMeetingKey } from "@/lib/meeting-key";
import type { TicketLane } from "@/types/meeting";

export type { TicketLane };

export function ticketLaneOf(item: { lane?: TicketLane }): TicketLane {
  return item.lane === "sheets" ? "sheets" : "development";
}

export function buildWorkspaceMeetingKey(
  meeting: {
    joinUrl: string;
    start: string;
    seriesMasterId?: string;
    eventId?: string;
  },
  lane: TicketLane,
): string {
  const key = buildMeetingKey(meeting);
  return lane === "sheets" ? `${key}|sheets` : key;
}
