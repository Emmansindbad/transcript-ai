import "server-only";

import { marked } from "marked";
import type { MeetingActionItem } from "@/types/meeting";
import { normalizeDevTag } from "@/lib/dev-tags";
import { getTargetUserId, getSummaryEmailRecipient } from "@/lib/env";
import { AppError, GraphError } from "@/lib/errors";
import { getGraphClient, mapGraphError, resetGraphAuthCache } from "@/lib/m365";
import { renderMarkdownLinksAsHtml, retainCrucialSourceLinks } from "@/lib/chat-links";
import { assertPmtLinks } from "@/lib/pmt";

export interface EmailSendResult {
  sent: number;
  recipients: string[];
  unmatched: string[];
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderMarkdownToHtml(markdown: string): string {
  // marked.parse is synchronous when not using async extensions
  const html = marked.parse(markdown, { async: false }) as string;
  return html;
}

function fieldLabel(text: string): string {
  return `<span style="display:block;margin:0 0 6px;font-size:11px;font-weight:700;letter-spacing:0.04em;color:#64748b;text-transform:uppercase;">${escapeHtml(text)}</span>`;
}

function fieldBox(value: string): string {
  return `<div style="min-height:38px;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;background:#ffffff;color:#0f172a;font-size:14px;">${escapeHtml(value)}</div>`;
}

function titleBox(item: MeetingActionItem): string {
  const title = item.title.trim() || item.task.trim() || "Untitled ticket";
  if (!item.pmtUrl?.trim()) {
    return fieldBox(title);
  }
  const label = item.pmtNumber ? `${title} · PMT #${item.pmtNumber}` : title;
  return `<div style="min-height:38px;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;background:#ffffff;font-size:14px;"><a href="${escapeHtml(item.pmtUrl)}" style="color:#4338ca;font-weight:600;">${escapeHtml(label)}</a></div>`;
}

function renderDescription(text: string): string {
  return renderMarkdownLinksAsHtml(retainCrucialSourceLinks(text)).replace(
    /<a href="/g,
    '<a style="color:#4338ca;font-weight:600;" href="',
  );
}

function buildActionItemCards(items: MeetingActionItem[]): string {
  return items
    .map((item, index) => {
      const due =
        item.deadline.trim().toUpperCase() === "TBD"
          ? "mm/dd/yyyy"
          : item.deadline.trim();

      return `
      <div style="margin:0 0 20px;padding:20px;border:1px solid #e2e8f0;border-radius:10px;background:#f8fafc;">
        <p style="margin:0 0 16px;font-size:13px;font-weight:700;color:#4f46e5;">Task #${index + 1} · ${escapeHtml(item.meetingTitle)} · ${escapeHtml(item.meetingDate)}</p>
        <div style="margin-bottom:14px;">
          ${fieldLabel("Title")}
          ${titleBox(item)}
        </div>
        <table style="width:100%;border-collapse:separate;border-spacing:0 12px;">
          <tr>
            <td style="width:50%;padding-right:8px;vertical-align:top;">
              ${fieldLabel("Type")}
              ${fieldBox(item.type || "Task")}
            </td>
            <td style="width:50%;padding-left:8px;vertical-align:top;">
              ${fieldLabel("Status")}
              ${fieldBox(item.status || "New")}
            </td>
          </tr>
          <tr>
            <td style="padding-right:8px;vertical-align:top;">
              ${fieldLabel("Priority")}
              ${fieldBox(item.priority || "Medium")}
            </td>
            <td style="padding-left:8px;vertical-align:top;">
              ${fieldLabel("Due")}
              ${fieldBox(due)}
            </td>
          </tr>
          <tr>
            <td colspan="2" style="vertical-align:top;">
              ${fieldLabel("Assignee")}
              ${fieldBox(item.pointPerson)}
            </td>
          </tr>
          <tr>
            <td style="padding-right:8px;vertical-align:top;">
              ${fieldLabel("Dev Tag")}
              ${fieldBox(normalizeDevTag(item.devTag))}
            </td>
            <td style="padding-left:8px;vertical-align:top;">
              ${fieldLabel("Related To")}
              ${fieldBox(item.relatedTo || "None")}
            </td>
          </tr>
        </table>
        <div style="margin-top:4px;">
          ${fieldLabel("Description")}
          <div style="min-height:88px;padding:10px 12px;border:1px solid #cbd5e1;border-radius:6px;background:#ffffff;color:#334155;font-size:14px;line-height:1.55;">${renderDescription(item.description || item.task)}</div>
        </div>
        ${
          item.assigneeUpn
            ? `<p style="margin:10px 0 0;font-size:12px;color:#64748b;">Entra: ${escapeHtml(item.assigneeDisplayName ?? item.pointPerson)} (${escapeHtml(item.assigneeUpn)})</p>`
            : ""
        }
        ${
          item.evidence?.length
            ? `<div style="margin-top:12px;">${fieldLabel("Evidence")}${item.evidence
                .map(
                  (quote) =>
                    `<p style="margin:0 0 6px;font-size:13px;color:#475569;"><strong>${escapeHtml(quote.timestamp || "time n/a")}${quote.speaker ? ` · ${escapeHtml(quote.speaker)}` : ""}</strong> — “${escapeHtml(quote.quote)}”</p>`,
                )
                .join("")}</div>`
            : ""
        }
      </div>`;
    })
    .join("");
}

function buildEmailBody(
  items: MeetingActionItem[],
  analysisHtml: string,
): string {
  const uniqueMeetings = [...new Set(items.map((i) => i.meetingTitle))];
  const uniquePersons = [...new Set(items.map((i) => i.pointPerson))];

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<style>
  /* Prose styles for the rendered markdown section */
  .analysis h1{font-size:18px;font-weight:700;color:#0f172a;margin:24px 0 8px}
  .analysis h2{font-size:16px;font-weight:700;color:#1e293b;margin:20px 0 6px;padding-bottom:4px;border-bottom:1px solid #e2e8f0}
  .analysis h3{font-size:14px;font-weight:700;color:#334155;margin:16px 0 4px}
  .analysis p{margin:0 0 10px;color:#475569;line-height:1.6}
  .analysis ul,.analysis ol{margin:0 0 10px;padding-left:20px;color:#475569;line-height:1.7}
  .analysis li{margin-bottom:2px}
  .analysis strong{color:#1e293b;font-weight:600}
  .analysis em{font-style:italic}
  .analysis blockquote{margin:0 0 10px;padding:8px 14px;border-left:3px solid #4f46e5;background:#f8f7ff;color:#475569}
  .analysis code{background:#f1f5f9;padding:1px 5px;border-radius:4px;font-size:12px;font-family:monospace}
  .analysis hr{border:none;border-top:1px solid #e2e8f0;margin:20px 0}
</style>
</head>
<body style="font-family:Segoe UI,Arial,sans-serif;color:#1e293b;background:#f8fafc;margin:0;padding:0;">
  <div style="max-width:760px;margin:32px auto;background:#ffffff;border-radius:12px;border:1px solid #e2e8f0;overflow:hidden;">

    <!-- Header -->
    <div style="background:#4f46e5;padding:28px 32px;">
      <h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:600;">Meeting Action Items &amp; Summary</h1>
      <p style="margin:8px 0 0;color:#c7d2fe;font-size:14px;">
        ${items.length} item${items.length === 1 ? "" : "s"} &nbsp;·&nbsp;
        ${uniquePersons.length} assignee${uniquePersons.length === 1 ? "" : "s"} &nbsp;·&nbsp;
        ${uniqueMeetings.length} meeting${uniqueMeetings.length === 1 ? "" : "s"}
      </p>
    </div>

    <!-- Action Items -->
    <div style="padding:28px 32px 8px;">
      <h2 style="margin:0 0 14px;font-size:15px;font-weight:700;color:#0f172a;letter-spacing:-0.01em;">Action Items</h2>
      ${buildActionItemCards(items)}
    </div>

    <!-- Divider -->
    <div style="margin:8px 32px 0;border-top:2px solid #e2e8f0;"></div>

    <!-- AI Analysis -->
    <div style="padding:24px 32px 28px;" class="analysis">
      <h2 style="margin:0 0 16px;font-size:15px;font-weight:700;color:#0f172a;letter-spacing:-0.01em;">AI Analysis</h2>
      ${analysisHtml}
    </div>

    <!-- Footer -->
    <div style="padding:16px 32px;background:#f8fafc;border-top:1px solid #e2e8f0;">
      <p style="margin:0;color:#94a3b8;font-size:12px;">Generated by M365 Transcript AI Analyzer</p>
    </div>

  </div>
</body>
</html>`;
}

function buildEmailSubject(items: MeetingActionItem[]): string {
  const titles = [...new Set(items.map((i) => i.meetingTitle))];
  if (titles.length === 1) {
    return `Meeting Summary — ${titles[0]}`;
  }
  return `Meeting Summary — ${titles.length} Meetings`;
}

async function sendEmailWithClient(
  actionItems: MeetingActionItem[],
  analysis: string,
): Promise<EmailSendResult> {
  assertPmtLinks(actionItems);
  const senderId = getTargetUserId();
  const recipient = getSummaryEmailRecipient();
  const client = getGraphClient();

  const subject = buildEmailSubject(actionItems);
  const analysisHtml = renderMarkdownToHtml(analysis);
  const htmlBody = buildEmailBody(actionItems, analysisHtml);

  try {
    await client.api(`/users/${senderId}/sendMail`).post({
      message: {
        subject,
        body: { contentType: "HTML", content: htmlBody },
        toRecipients: [{ emailAddress: { address: recipient } }],
      },
      saveToSentItems: true,
    });
  } catch (error) {
    throw mapGraphError(error, `Failed to send email to ${recipient}.`);
  }

  return { sent: 1, recipients: [recipient], unmatched: [] };
}

export async function sendActionItemsEmail(
  actionItems: MeetingActionItem[],
  analysis: string,
): Promise<EmailSendResult> {
  if (actionItems.length === 0) {
    throw new GraphError(
      "No action items to send via email.",
      400,
      "NO_ACTION_ITEMS",
    );
  }

  try {
    return await sendEmailWithClient(actionItems, analysis);
  } catch (error) {
    if (error instanceof AppError && !(error instanceof GraphError)) {
      throw error;
    }
    const mapped =
      error instanceof GraphError
        ? error
        : mapGraphError(error, "Failed to send action items email.");

    if (mapped.statusCode !== 401) {
      throw mapped;
    }

    resetGraphAuthCache();

    try {
      return await sendEmailWithClient(actionItems, analysis);
    } catch (retryError) {
      if (retryError instanceof AppError) {
        throw retryError;
      }
      throw mapGraphError(retryError, "Failed to send action items email.");
    }
  }
}
