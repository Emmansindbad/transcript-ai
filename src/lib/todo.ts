import "server-only";

import type { Client } from "@microsoft/microsoft-graph-client";
import type { MeetingActionItem } from "@/types/meeting";
import { normalizeDevTag } from "@/lib/dev-tags";
import { getTargetUserId } from "@/lib/env";
import { GraphError } from "@/lib/errors";
import { getGraphClient, mapGraphError, resetGraphAuthCache } from "@/lib/m365";

const MEETING_TODO_LIST_NAME = "Meeting Action Items";
export const MICROSOFT_TODO_URL = "https://to-do.office.com/tasks/";

interface TodoList {
  id?: string;
  displayName?: string;
  wellknownListName?: string;
}

interface TodoListsResponse {
  value?: TodoList[];
}

function parseDeadline(deadline: string): string | null {
  if (!deadline || deadline.toUpperCase() === "TBD") {
    return null;
  }

  const parsed = new Date(deadline);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed.toISOString().split("T")[0];
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function truncateText(text: string, maxLength: number): string {
  const cleaned = text.trim().replace(/\s+/g, " ");
  if (cleaned.length <= maxLength) {
    return cleaned;
  }

  const truncated = cleaned.slice(0, maxLength - 1);
  const lastSpace = truncated.lastIndexOf(" ");
  if (lastSpace > maxLength * 0.6) {
    return `${truncated.slice(0, lastSpace)}…`;
  }

  return `${truncated}…`;
}

function buildTaskTitle(item: MeetingActionItem): string {
  return truncateText(item.title || item.task, 100);
}

function buildTaskBody(item: MeetingActionItem): string {
  const deadline =
    item.deadline.trim().toUpperCase() === "TBD"
      ? "No deadline set"
      : item.deadline.trim();

  return `<div style="font-family:Segoe UI,sans-serif;font-size:14px;line-height:1.5;color:#242424;">
  <p style="margin:0 0 12px;font-size:15px;font-weight:600;">${escapeHtml(item.title || item.task)}</p>
  <p style="margin:0 0 12px;white-space:pre-wrap;">${escapeHtml(item.description || item.task)}</p>
  <table style="border-collapse:collapse;width:100%;">
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Type</td>
      <td style="padding:6px 0;">${escapeHtml(item.type)}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Status</td>
      <td style="padding:6px 0;">${escapeHtml(item.status || "New")}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Priority</td>
      <td style="padding:6px 0;">${escapeHtml(item.priority)}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Assignee</td>
      <td style="padding:6px 0;font-weight:600;">${escapeHtml(item.pointPerson)}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Dev tag</td>
      <td style="padding:6px 0;">${escapeHtml(normalizeDevTag(item.devTag))}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Related to</td>
      <td style="padding:6px 0;">${escapeHtml(item.relatedTo || "None")}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Meeting</td>
      <td style="padding:6px 0;">${escapeHtml(item.meetingTitle)}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Meeting date</td>
      <td style="padding:6px 0;">${escapeHtml(item.meetingDate)}</td>
    </tr>
    <tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Due</td>
      <td style="padding:6px 0;">${escapeHtml(deadline)}</td>
    </tr>
    ${
      item.assigneeUpn
        ? `<tr>
      <td style="padding:6px 16px 6px 0;color:#616161;white-space:nowrap;vertical-align:top;">Entra</td>
      <td style="padding:6px 0;">${escapeHtml(item.assigneeUpn)}</td>
    </tr>`
        : ""
    }
  </table>
  ${
    item.evidence?.length
      ? `<p style="margin:12px 0 6px;font-weight:600;">Evidence</p>${item.evidence
          .map(
            (quote) =>
              `<p style="margin:0 0 6px;">${escapeHtml(quote.timestamp || "")} ${escapeHtml(quote.speaker || "")}: ${escapeHtml(quote.quote)}</p>`,
          )
          .join("")}`
      : ""
  }
</div>`;
}

function buildTaskPayload(item: MeetingActionItem): Record<string, unknown> {
  const dueDate = parseDeadline(item.deadline);
  const payload: Record<string, unknown> = {
    title: buildTaskTitle(item),
    body: {
      contentType: "html",
      content: buildTaskBody(item),
    },
    importance:
      item.priority === "High"
        ? "high"
        : item.priority === "Low"
          ? "low"
          : "normal",
  };

  if (dueDate) {
    payload.dueDateTime = {
      dateTime: `${dueDate}T09:00:00`,
      timeZone: "UTC",
    };
  }

  return payload;
}

function isStaleTokenError(error: GraphError): boolean {
  return error.statusCode === 401 || error.code === "MISSING_TASKS_PERMISSION";
}

async function getOrCreateMeetingTodoList(
  client: Client,
  userId: string,
): Promise<string> {
  try {
    const response = (await client
      .api(`/users/${userId}/todo/lists`)
      .get()) as TodoListsResponse;

    const existing = (response.value ?? []).find(
      (list) => list.displayName === MEETING_TODO_LIST_NAME,
    );

    if (existing?.id) {
      return existing.id;
    }

    const created = (await client.api(`/users/${userId}/todo/lists`).post({
      displayName: MEETING_TODO_LIST_NAME,
    })) as TodoList;

    if (!created.id) {
      throw new GraphError(
        "Failed to create Microsoft To Do list for meeting action items.",
        500,
        "TODO_LIST_CREATE_FAILED",
      );
    }

    return created.id;
  } catch (error) {
    throw mapGraphError(error, "Failed to access Microsoft To Do.");
  }
}

async function exportActionItemsWithClient(
  actionItems: MeetingActionItem[],
): Promise<{ created: number; listName: string }> {
  const userId = getTargetUserId();
  const client = getGraphClient();
  const listId = await getOrCreateMeetingTodoList(client, userId);

  let created = 0;

  for (const item of actionItems) {
    const taskPayload = buildTaskPayload(item);

    try {
      await client
        .api(`/users/${userId}/todo/lists/${listId}/tasks`)
        .post(taskPayload);
    } catch (error) {
      throw mapGraphError(error, "Failed to create a task in Microsoft To Do.");
    }

    created += 1;
  }

  return { created, listName: MEETING_TODO_LIST_NAME };
}

export async function exportActionItemsToTodo(
  actionItems: MeetingActionItem[],
): Promise<{ created: number; listName: string }> {
  if (actionItems.length === 0) {
    throw new GraphError(
      "No action items to export to Microsoft To Do.",
      400,
      "NO_ACTION_ITEMS",
    );
  }

  try {
    return await exportActionItemsWithClient(actionItems);
  } catch (error) {
    const mapped =
      error instanceof GraphError
        ? error
        : mapGraphError(error, "Failed to export action items to Microsoft To Do.");

    if (!isStaleTokenError(mapped)) {
      throw mapped;
    }

    resetGraphAuthCache();

    try {
      return await exportActionItemsWithClient(actionItems);
    } catch (retryError) {
      throw retryError instanceof GraphError
        ? retryError
        : mapGraphError(retryError, "Failed to export action items to Microsoft To Do.");
    }
  }
}
