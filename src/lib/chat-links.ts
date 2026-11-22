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

const GENERIC_SOURCE_TITLE =
  /^(chat message|link|message|teams message|screenshot in chat|screenshot|image|photo|picture|pic)$/i;

export function isGenericSourceTitle(title: string): boolean {
  return GENERIC_SOURCE_TITLE.test(title.trim());
}

function isVisualLinkTitle(title: string): boolean {
  return IMAGE_NAME.test(title) || /\b(screenshot|image|photo|picture|attachment)\b/i.test(title);
}

function isTeamsMessagePermalink(url: string): boolean {
  const href = toTeamsMessagePermalink(url);
  try {
    const parsed = new URL(href);
    return TEAMS_HOST.test(parsed.hostname) && TEAMS_MESSAGE_PATH.test(parsed.pathname);
  } catch {
    return false;
  }
}

function collapseSourceSpacing(text: string): string {
  return text
    .split(/\r?\n/)
    .map((line) => line.replace(/[ \t]{2,}/g, " ").trimEnd())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function markdownLinkPattern(): RegExp {
  return new RegExp(MARKDOWN_LINK.source, "g");
}

interface FoundMarkdownLink {
  index: number;
  length: number;
  title: string;
  url: string;
  teams: boolean;
}

function findMarkdownLinks(text: string): FoundMarkdownLink[] {
  const found: FoundMarkdownLink[] = [];
  const matcher = markdownLinkPattern();
  let match = matcher.exec(text);
  while (match) {
    const title = match[1]?.trim() ?? "";
    const url = match[2]?.trim() ?? "";
    if (title && isSafeHttpUrl(url)) {
      found.push({
        index: match.index,
        length: match[0].length,
        title,
        url: toTeamsMessagePermalink(url),
        teams: isTeamsMessagePermalink(url),
      });
    }
    match = matcher.exec(text);
  }
  return found;
}

/**
 * Tickets keep one Teams source link. A named screenshot or file the ticket
 * actually mentions wins. Otherwise the first source link stays, and every
 * extra "Chat message" link is removed. Links that are not Teams messages stay.
 */
export function retainCrucialSourceLinks(text: string): string {
  if (!text.trim()) {
    return "";
  }
  const rewritten = rewriteTeamsPermalinks(text);
  const found = findMarkdownLinks(rewritten);
  const teams = found.filter((link) => link.teams);
  const withoutLinks = rewritten.replace(markdownLinkPattern(), " ");
  const named = teams.find((link) => {
    if (isGenericSourceTitle(link.title)) {
      return false;
    }
    const needle = normalize(imageLinkTitle(link.title));
    return needle.length >= 6 && normalize(withoutLinks).includes(needle);
  });
  const winner = named ?? teams[0];

  let result = "";
  let cursor = 0;
  for (const link of found) {
    const keep = !link.teams || link === winner;
    result += rewritten.slice(cursor, link.index);
    if (keep) {
      result += markdownLink(link.title, link.url);
    }
    cursor = link.index + link.length;
  }
  result += rewritten.slice(cursor);
  return collapseSourceSpacing(result);
}

export function formatChatSourceLinks(webUrl: string, images: string[]): string {
  if (!isSafeHttpUrl(webUrl)) {
    return "";
  }
  if (images.length === 0) {
    return markdownLink("Chat message", webUrl);
  }
  const named = images
    .map((name) => imageLinkTitle(name))
    .find((title) => !isGenericSourceTitle(title));
  return markdownLink(named || "Screenshot in chat", webUrl);
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
  const matcher = markdownLinkPattern();
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
  const matcher = markdownLinkPattern();
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
      const text = afterSpeaker.replace(markdownLinkPattern(), "").replace(/\s+/g, " ").trim();
      return { speaker, text, links };
    })
    .filter((line) => line.links.length > 0);
}

function normalize(value: string): string {
  return value.toLowerCase().replace(/\s+/g, " ").trim();
}

function lineMatchScore(line: ParsedChatLine, item: ParsedActionItem): number {
  const hay = normalize(
    `${item.title} ${item.task} ${item.description} ${item.evidence.map((entry) => entry.quote).join(" ")}`,
  );
  let score = 0;
  if (line.text) {
    const snippet = normalize(line.text).slice(0, 48);
    if (snippet.length >= 20 && hay.includes(snippet)) {
      score += 5;
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
    const lineText = normalize(line.text);
    if (
      speakerMatch &&
      lineText &&
      quoteSnippet.length >= 12 &&
      (lineText.includes(quoteSnippet) || quote.includes(lineText.slice(0, 36)))
    ) {
      score += 8;
    }
  }

  for (const link of line.links) {
    const title = extractMarkdownLinks(link)[0]?.title ?? "";
    if (isGenericSourceTitle(title) || isGenericSourceTitle(imageLinkTitle(title))) {
      continue;
    }
    const needle = normalize(imageLinkTitle(title));
    if (needle.length >= 6 && hay.includes(needle)) {
      score += 6;
    }
  }

  const mentionsVisual = /\b(screenshot|image|photo|attached|attachment|picture|pic)\b/i.test(hay);
  if (
    mentionsVisual &&
    line.links.some((link) => isVisualLinkTitle(extractMarkdownLinks(link)[0]?.title ?? "")) &&
    item.evidence.some(
      (entry) => entry.speaker.trim() && normalize(entry.speaker) === normalize(line.speaker),
    )
  ) {
    score += 4;
  }

  return score;
}

function bestMatchingLine(lines: ParsedChatLine[], item: ParsedActionItem): ParsedChatLine | null {
  let best: ParsedChatLine | null = null;
  let bestScore = 0;
  for (const line of lines) {
    const score = lineMatchScore(line, item);
    if (score > bestScore) {
      best = line;
      bestScore = score;
    }
  }
  return best;
}

function preferredLineLink(line: ParsedChatLine): string {
  const visual = line.links.find((link) =>
    isVisualLinkTitle(extractMarkdownLinks(link)[0]?.title ?? ""),
  );
  return visual ?? line.links[0] ?? "";
}

function descriptionHasTeamsLink(description: string): boolean {
  return extractMarkdownLinks(description).some((link) => isTeamsMessagePermalink(link.url));
}

export function attachChatSourceLinks(
  items: ParsedActionItem[],
  chat: string,
): ParsedActionItem[] {
  const lines = parseFormattedChat(chat);

  return items.map((item) => {
    let description = retainCrucialSourceLinks(item.description);
    if (!descriptionHasTeamsLink(description)) {
      const best = bestMatchingLine(lines, item);
      const link = best ? preferredLineLink(best) : "";
      if (link) {
        description = retainCrucialSourceLinks(appendSourceLinks(description, [link]));
      }
    }
    if (description === item.description) {
      return item;
    }
    return {
      ...item,
      description,
    };
  });
}
