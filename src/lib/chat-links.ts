import type { ParsedActionItem } from "@/types/meeting";

const MARKDOWN_LINK = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g;
const IMAGE_NAME = /\.(png|jpe?g|gif|webp|bmp|heic|svg)$/i;
const TEAMS_MESSAGE_CONTEXT = encodeURIComponent('{"contextType":"chat"}');
const TEAMS_HOST = /(?:^|\.)teams\.(?:microsoft\.com|cloud\.microsoft)$/i;
const TEAMS_MESSAGE_PATH = /\/l\/message\/([^/?#]+)\/(\d+)/i;
const BARE_URL = /https?:\/\/[^\s)\]]+/g;

export function isSafeHttpUrl(url: string): boolean {
  try {
    const parsed = new URL(url.trim());
    return parsed.protocol === "https:" || parsed.protocol === "http:";
  } catch {
    return false;
  }
}

export function toTeamsMessagePermalink(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) {
    return "";
  }

  let source = trimmed;
  try {
    const parsed = new URL(trimmed);
    if (TEAMS_HOST.test(parsed.hostname) && parsed.pathname.includes("/dl/launcher/launcher.html")) {
      const inner = parsed.searchParams.get("url") ?? "";
      if (inner) {
        source = inner;
        try {
          if (/%[0-9A-Fa-f]{2}/.test(source)) {
            source = decodeURIComponent(source);
          }
        } catch {
          // Keep the once-decoded launcher path.
        }
      }
    }
  } catch {
    return trimmed;
  }

  const match = source.match(TEAMS_MESSAGE_PATH);
  const chatId = match?.[1] ? decodeURIComponent(match[1]) : "";
  const messageId = match?.[2] ?? "";
  if (!chatId || !messageId) {
    return trimmed;
  }
  return `https://teams.cloud.microsoft/l/message/${chatId}/${messageId}?context=${TEAMS_MESSAGE_CONTEXT}`;
}

export function rewriteTeamsPermalinks(text: string): string {
  return text.replace(BARE_URL, (url) => {
    const trailing = url.match(/[.,;]+$/)?.[0] ?? "";
    const href = trailing ? url.slice(0, -trailing.length) : url;
    return `${toTeamsMessagePermalink(href)}${trailing}`;
  });
}

export function markdownLink(title: string, url: string): string {
  const href = toTeamsMessagePermalink(url);
  if (!isSafeHttpUrl(href)) {
    return "";
  }
  const label = title.replace(/[\[\]]/g, "").replace(/\s+/g, " ").trim() || "Link";
  return `[${label}](${href})`;
}

export function imageLinkTitle(name: string): string {
  const cleaned = name.replace(/^screenshot\s*[·\-:]\s*/i, "").trim();
  if (!cleaned || /^screenshot$/i.test(cleaned) || /^image$/i.test(cleaned)) {
    return "Screenshot in chat";
  }
  return cleaned.slice(0, 80);
}

export function formatChatSourceLinks(webUrl: string, images: string[]): string {
  if (!isSafeHttpUrl(webUrl)) {
    return "";
  }
  if (images.length > 0) {
    return images
      .map((name) => markdownLink(imageLinkTitle(name), webUrl))
      .filter(Boolean)
      .join(" ");
  }
  return markdownLink("Chat message", webUrl);
}

export function isImageAttachmentName(name: string, contentType?: string): boolean {
  const type = contentType?.toLowerCase() ?? "";
  if (type.startsWith("image/")) {
    return true;
  }
  return IMAGE_NAME.test(name);
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function renderMarkdownLinksAsHtml(text: string): string {
  const parts: string[] = [];
  const matcher = new RegExp(MARKDOWN_LINK.source, "g");
  let last = 0;
  let match = matcher.exec(text);
  while (match) {
    parts.push(escapeHtml(text.slice(last, match.index)).replace(/\r\n|\n|\r/g, "<br/>"));
    parts.push(
      `<a href="${escapeHtml(match[2] ?? "")}">${escapeHtml(match[1] ?? "Link")}</a>`,
    );
    last = match.index + match[0].length;
    match = matcher.exec(text);
  }
  parts.push(escapeHtml(text.slice(last)).replace(/\r\n|\n|\r/g, "<br/>"));
  return parts.join("");
}

export function extractMarkdownLinks(text: string): Array<{ title: string; url: string }> {
  const links: Array<{ title: string; url: string }> = [];
  const seen = new Set<string>();
  const matcher = new RegExp(MARKDOWN_LINK.source, "g");
  let match = matcher.exec(text);
  while (match) {
    const title = match[1]?.trim() ?? "";
    const url = match[2]?.trim() ?? "";
    if (title && isSafeHttpUrl(url) && !seen.has(url)) {
      seen.add(url);
      links.push({ title, url });
    }
    match = matcher.exec(text);
  }
  return links;
}

function appendSourceLinks(description: string, links: string[]): string {
  const unique = [...new Set(links.filter(Boolean))];
  const missing = unique.filter((link) => {
    const url = extractMarkdownLinks(link)[0]?.url;
    return url ? !description.includes(url) : !description.includes(link);
  });
  if (missing.length === 0) {
    return description;
  }
  return `${description.trim()}\n\n${missing.join("\n")}`;
}

interface ParsedChatLine {
  speaker: string;
  text: string;
  links: string[];
}

function parseFormattedChat(chat: string): ParsedChatLine[] {
  return chat
    .split(/\r?\n/)
    .map((raw) => raw.trim())
    .filter(Boolean)
    .map((line) => {
      const speakerMatch = line.match(/\bKSA\s+(.+?):\s*/);
      const speaker = speakerMatch?.[1]?.trim() ?? "";
      const afterSpeaker = speakerMatch ? line.slice(line.indexOf(speakerMatch[0]) + speakerMatch[0].length) : line;
      const links = extractMarkdownLinks(line).map((link) => markdownLink(link.title, link.url));
      const text = afterSpeaker.replace(MARKDOWN_LINK, "").replace(/\s+/g, " ").trim();
      return { speaker, text, links };
    })
    .filter((line) => line.links.length > 0);
}

function normalize(value: string): string {
  return value.toLowerCase().replace(/\s+/g, " ").trim();
}

function lineMatchesItem(line: ParsedChatLine, item: ParsedActionItem): boolean {
  const hay = normalize(
    `${item.title} ${item.task} ${item.description} ${item.evidence.map((entry) => entry.quote).join(" ")}`,
  );
  if (line.text) {
    const snippet = normalize(line.text).slice(0, 48);
    if (snippet.length >= 12 && hay.includes(snippet)) {
      return true;
    }
  }

  for (const evidence of item.evidence) {
    const quote = normalize(evidence.quote);
    if (!quote) {
      continue;
    }
    const quoteSnippet = quote.slice(0, 36);
    const speakerMatch =
      !evidence.speaker.trim() ||
      !line.speaker ||
      normalize(evidence.speaker) === normalize(line.speaker);
    if (speakerMatch && line.text && (normalize(line.text).includes(quoteSnippet) || quote.includes(normalize(line.text).slice(0, 36)))) {
      return true;
    }
  }

  for (const link of line.links) {
    const title = extractMarkdownLinks(link)[0]?.title ?? "";
    const needle = normalize(imageLinkTitle(title));
    if (needle.length >= 6 && hay.includes(needle)) {
      return true;
    }
  }

  const mentionsVisual = /\b(screenshot|image|photo|attached|attachment|picture|pic)\b/i.test(hay);
  if (mentionsVisual && line.links.some((link) => /screenshot|image|\.(png|jpe?g|gif|webp)/i.test(link))) {
    return item.evidence.some(
      (entry) => entry.speaker.trim() && normalize(entry.speaker) === normalize(line.speaker),
    );
  }

  return false;
}

export function attachChatSourceLinks(
  items: ParsedActionItem[],
  chat: string,
): ParsedActionItem[] {
  const lines = parseFormattedChat(chat);
  if (lines.length === 0) {
    return items;
  }

  return items.map((item) => {
    const links = lines.filter((line) => lineMatchesItem(line, item)).flatMap((line) => line.links);
    const description = rewriteTeamsPermalinks(
      links.length > 0 ? appendSourceLinks(item.description, links) : item.description,
    );
    if (description === item.description) {
      return item;
    }
    return {
      ...item,
      description,
    };
  });
}
