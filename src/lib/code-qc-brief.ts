export const VISIBLE_CODE_QC_TICKETS = 12;
export const CODE_QC_MORE_BATCH = 40;
export const CODE_QC_CREATE_CHUNK = 20;
export const CODE_QC_MAX_GROUPED_TICKETS = 40;

export const CODE_QC_DEVOPS_CC = [
  { displayName: "Jehad Jamel Alfarasani", email: "jehad.alfarasani@sindbad.tech" },
  { displayName: "Arnold Gutib", email: "arnold.gutib@sindbad.tech" },
] as const;

export interface CodeQcNoteTicket {
  title: string;
  number?: number;
  url?: string;
}

export function displayBoardName(name: string): string {
  return name.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
}

function numberedTicketLine(index: number, ticket: CodeQcNoteTicket): string {
  const title = ticket.title.replace(/\s+/g, " ").trim() || "CODE QC ticket";
  if (ticket.url && ticket.number && ticket.number > 0) {
    const label = `#${ticket.number} ${title}`.replace(/[\[\]]/g, "");
    return `${index}. [${label}](${ticket.url})`;
  }
  return `${index}. ${title}`;
}

export function defaultCodeQcNote(input: {
  boardName: string;
  assigneeName?: string;
  tickets: CodeQcNoteTicket[];
}): string {
  const board = displayBoardName(input.boardName) || "this PMT board";
  const assignee = input.assigneeName?.replace(/\s+/g, " ").trim() ?? "";
  const greeting = assignee ? `@${assignee}` : "Hi team";
  const heading =
    input.tickets.length === 1
      ? `Please fix this CODE QC ticket on ${board}.`
      : input.tickets.length > 1
        ? `Please fix these ${input.tickets.length} CODE QC tickets on ${board}.`
        : `Please fix the CODE QC tickets on ${board}.`;
  const listed =
    input.tickets.length > 0
      ? input.tickets.map((ticket, index) => numberedTicketLine(index + 1, ticket)).join("\n")
      : "Tickets will be listed here after Gemini proposes them.";
  const cc = CODE_QC_DEVOPS_CC.map((person) => `@${person.displayName}`).join(" and ");
  return [greeting, "", heading, "", listed, "", `CC ${cc} for devops.`].join("\n");
}
