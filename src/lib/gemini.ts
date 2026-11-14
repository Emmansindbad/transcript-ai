import "server-only";

import {
  GoogleGenAI,
  createPartFromUri,
  createUserContent,
  FileState,
  type ContentListUnion,
} from "@google/genai";
import type { CodeQcIssue } from "@/types/code-qc";
import type {
  ActionItemEvidence,
  ActionItemPriority,
  ActionItemType,
  ParsedActionItem,
} from "@/types/meeting";
import { attachChatSourceLinks } from "@/lib/chat-links";
import { DEV_TAGS, normalizeDevTag } from "@/lib/dev-tags";
import { getEnvVar } from "@/lib/env";
import { GeminiError } from "@/lib/errors";

export const SYSTEM_PROMPT = `You are an elite Technical Project Manager and Systems Architect.

Your job is to read the meeting source carefully and produce tickets plus analysis that a delivery lead can act on without re-watching the meeting.

GROUNDING
- Use only what is in the spoken transcript, the meeting chat, and the recording when provided.
- Prefer names, owners, dates, numbers, thresholds, product names, and use-case IDs that were actually said, typed in chat, or shown.
- Do not invent owners, deadlines, systems, or acceptance criteria.
- If the transcript is thin or missing, use the meeting chat and recording.
- Cover the whole meeting, not only the last few minutes.
- Split distinct pieces of work into separate tickets. Merge only when they are clearly the same job.

REDUNDANCY FOR THAT DAY
- All times are Arabia Standard Time (KSA, Asia/Riyadh, UTC+3).
- The comparison day is the KSA calendar date when the transcript or recording was created, not the viewer's local time.
- Compare spoken transcript and meeting chat only for that KSA day.
- If the same action, decision, or bug appears in both, create ONE ticket and note that it was repeated in chat and speech.
- Chat-only items still become tickets. Transcript-only items still become tickets.
- Do not emit duplicate tickets for the same work.

STEP 1 — Return ONE JSON object only (no markdown fences, no text before or after):
{
  "actionItems": [
    {
      "title": "Inactive User Redemption Trigger",
      "type": "Task",
      "status": "New",
      "priority": "Medium",
      "pointPerson": "Mohammed Agbawi",
      "task": "Implement inactive-user redemption when portfolio value falls below 100K.",
      "deadline": "TBD",
      "devTag": "#backend",
      "relatedTo": "Use Case 7",
      "description": "Build the inactive-user redemption rule discussed in the meeting.\\n1. Trigger redemption when portfolio value (Purchasing Power + Equity) is below 100K.\\n2. Confirm the threshold and owner with product before release.\\n\\n[Chat message](https://teams.cloud.microsoft/l/message/19:example@thread.v2/123?context=%7B%22contextType%22%3A%22chat%22%7D)",
      "evidence": [
        {
          "timestamp": "00:12:04",
          "speaker": "Mohammed Agbawi",
          "quote": "Trigger redemption when Purchasing Power plus Equity is below 100K."
        }
      ]
    }
  ],
  "analysis": "## Executive Summary\\n...markdown for THIS meeting..."
}

Be exhaustive in actionItems. Create a ticket for every:
- committed task or owner assignment
- implied follow-up ("we should", "need to", "next step")
- bug, defect, wrong behavior, or fix
- investigation, blocker, or missing decision that creates work
- use case, feature, or rule that must be implemented

Ticket field rules:
- title and description MUST be different. Never copy one into the other.
- title: 4–12 word headline that names the outcome or defect. Summarize the description; do not paste it. No numbered lists in the title.
- type: classify each item as "Task" or "Bug". Decide from the content, not from habit. Do not mark everything as Task.
  - Bug: something is broken, incorrect, failing, crashing, missing expected behavior, a regression, or needs a fix.
  - Task: new work, a feature, a use case, an enhancement, research, follow-up, or implementation.
- status: "New" unless the source says it is already in progress or done.
- priority: "High", "Medium", or "Low" from urgency, risk, customer impact, and deadlines in the meeting.
- pointPerson: assignee full name from the source. If several people are named, pick the person who owns the work. Use "Unassigned" only when no owner is identifiable.
- task: one actionable sentence. Not the title. Not the full description.
- deadline: exact date if mentioned, otherwise "TBD".
- devTag: exactly one of ${DEV_TAGS.join(", ")}. Choose the closest tag for the work. Use "None" when none of them fit. Never invent a module, squad, or system name.
- relatedTo: related feature, use case, epic, or meeting topic if mentioned, otherwise "None".
- description: full ticket body. At least two sentences or a numbered list. Include context, who is affected, trigger/acceptance criteria, and every figure, threshold, or date from the source. Use line breaks. Do not paste the title.
- evidence: 1–2 short quotes from the spoken transcript or chat that ground the ticket. Copy timestamp and speaker from lines like "[00:12:04] Name: ...". If no timestamp exists, use "".
- SOURCE LINKS: Meeting chat lines may already include Teams permalinks as markdown, e.g. [Chat message](https://teams.cloud.microsoft/l/message/...) or [screenshot.png](https://teams.cloud.microsoft/l/message/...). When a ticket is grounded in that chat message, or an image/screenshot attached there is the evidence, copy the exact [title](url) onto its own line at the end of description. Use a short accurate title (file name, "Screenshot in chat", or "Chat message"). Only add a link when it is necessary and traceable to that ticket. Never invent a URL. Never put links in the title. Never use teams.microsoft.com/dl/launcher URLs.
- If there is truly no work, return "actionItems": []

The "analysis" string is markdown for THIS meeting. Do not paste the JSON tickets again.

## Executive Summary
Three sentences: what was decided, what is blocked, and what must happen next.

## Software & App Development Tasks
Prioritized work with owners and timing when known. Call out bugs separately from new features.

## Fintech & Business Logic Integrations
Rules, calculations, thresholds, portfolio logic, payments, and other business behavior that was discussed.

## Transcript vs Chat (that day)
Use KSA time. Compare spoken discussion and meeting chat for the KSA calendar date when the transcript/recording was created: what was unique to speech, unique to chat, and repeated in both. List any redundant follow-ups that should stay as a single ticket.

## Strategic Advice
Blind spots, risks, missing owners, and recommendations grounded in the meeting.`;

export const VIDEO_SYSTEM_PROMPT = `${SYSTEM_PROMPT}

You are ALSO given the meeting recording. Treat the recording as a first-class source, especially if the transcript is empty or incomplete.
Triangulate spoken discussion with what is on screen:
- Shared slides, dashboards, tickets, spreadsheets, code, diagrams, and whiteboards
- On-screen text: titles, use-case names, numbers, owners, dates, and acceptance criteria
- Demonstrated bugs, wrong UI states, and product behavior
If the video, transcript, and chat conflict, prefer what is clearly shown on screen, then chat text, and note the conflict in Strategic Advice and Transcript vs Chat (that day).
Fold visual findings into the same JSON tickets and markdown sections. Do not add extra sections.`;

/**
 * Best-first models from ListModels on this API key (Oct 2026).
 * Default start is Gemini 3.8 Flash: current Flash for agents, video, and
 * structured extraction. 3.6 Flash is the previous generation and the next
 * fallback. 3.1 Pro stays above 3.8 and is used only when GEMINI_MODEL
 * points at it (heavier science/math reasoning, still preview).
 * gemini-3-pro-preview and gemini-2.0-* are shut down.
 * gemini-3.5-flash is deprecated and routed to 3.6; we call 3.8/3.6 directly.
 */
const DEFAULT_MODEL = "gemini-3.8-flash";

const MODEL_TIER = [
  "gemini-3.1-pro-preview",
  "gemini-pro-latest",
  "gemini-3.8-flash",
  "gemini-3.6-flash",
  "gemini-3-flash-preview",
  "gemini-flash-latest",
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-2.5-pro",
  "gemini-2.5-flash",
  "gemini-2.5-flash-lite",
] as const;

const MAX_RETRIES_PER_MODEL = 2;
const INITIAL_RETRY_DELAY_MS = 1000;
const MAX_OUTPUT_TOKENS = 65536;

const OUTPUT_BUDGET_RULES = `
OUTPUT BUDGET
- Return one complete JSON object. Never stop mid-string or omit the closing braces.
- Keep "analysis" under 500 words. Prefer short bullets over long paragraphs.
- Evidence quotes: one short sentence each.
- Prefer many complete tickets over a long write-up.`;

const COMPACT_OUTPUT_RULES = `
The previous attempt was truncated. Return COMPLETE valid JSON only.
- At most 12 actionItems (highest impact).
- "analysis" at most 250 words, short bullets.
- At most 1 evidence quote per ticket, under 15 words.
- Finish the JSON object.`;

const ANALYSIS_RESPONSE_SCHEMA = {
  type: "OBJECT",
  properties: {
    actionItems: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          title: { type: "STRING" },
          type: { type: "STRING" },
          status: { type: "STRING" },
          priority: { type: "STRING" },
          pointPerson: { type: "STRING" },
          task: { type: "STRING" },
          deadline: { type: "STRING" },
          devTag: { type: "STRING", enum: [...DEV_TAGS] },
          relatedTo: { type: "STRING" },
          description: { type: "STRING" },
          evidence: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                timestamp: { type: "STRING" },
                speaker: { type: "STRING" },
                quote: { type: "STRING" },
              },
              required: ["quote"],
            },
          },
        },
        required: ["title", "type", "description"],
      },
    },
    analysis: { type: "STRING" },
  },
  required: ["actionItems", "analysis"],
} as const;

export interface GeminiAnalysisResult {
  actionItems: ParsedActionItem[];
  analysis: string;
  modelUsed: string;
}

let geminiClient: GoogleGenAI | null = null;

function getGeminiClient(): GoogleGenAI {
  if (!geminiClient) {
    geminiClient = new GoogleGenAI({
      apiKey: getEnvVar("GEMINI_API_KEY"),
    });
  }
  return geminiClient;
}

function getModelTierRank(model: string): number {
  const rank = MODEL_TIER.indexOf(model as (typeof MODEL_TIER)[number]);
  return rank === -1 ? MODEL_TIER.length : rank;
}

/** Models ordered best-first; skips tiers above GEMINI_MODEL. Unset starts at 3.8 Flash. */
function getOrderedModels(): string[] {
  const configured = process.env.GEMINI_MODEL?.trim() || DEFAULT_MODEL;
  let pool: string[] = [...MODEL_TIER];

  if (configured && !pool.includes(configured)) {
    pool = [configured, ...pool];
  }

  if (configured) {
    const ceilingRank = getModelTierRank(configured);
    pool = pool.filter((model) => getModelTierRank(model) >= ceilingRank);
  }

  return [...new Set(pool)].sort(
    (left, right) => getModelTierRank(left) - getModelTierRank(right),
  );
}

function shouldFailFastToNextModel(error: GeminiError): boolean {
  return (
    error.code === "MODEL_OVERLOADED" ||
    error.code === "MODEL_NOT_FOUND" ||
    error.statusCode === 404 ||
    isHighDemandMessage(error.message)
  );
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isHighDemandMessage(message: string): boolean {
  const lower = message.toLowerCase();
  return (
    lower.includes("high demand") ||
    lower.includes("overloaded") ||
    lower.includes("unavailable") ||
    lower.includes("capacity")
  );
}

function extractGeminiErrorMessage(message: string): string {
  try {
    const parsed = JSON.parse(message) as {
      error?: { message?: string };
    };
    if (parsed.error?.message) {
      return parsed.error.message.split("\n")[0];
    }
  } catch {
    // Not JSON — use raw message.
  }
  return message;
}

function isQuotaOrRateLimitError(status: number, message: string): boolean {
  return (
    status === 429 ||
    message.toLowerCase().includes("quota") ||
    message.toLowerCase().includes("rate limit") ||
    message.toLowerCase().includes("resource_exhausted")
  );
}

function isIncompleteJsonMessage(message: string): boolean {
  const lower = message.toLowerCase();
  return (
    lower.includes("incomplete json segment") ||
    lower.includes("unexpected end-of-input") ||
    lower.includes("unterminated string")
  );
}

function isRetryableError(status: number, message: string, code: string): boolean {
  return (
    status === 429 ||
    status === 502 ||
    status === 503 ||
    code === "MODEL_OVERLOADED" ||
    code === "RATE_LIMIT" ||
    code === "EMPTY_RESPONSE" ||
    code === "INCOMPLETE_JSON" ||
    isIncompleteJsonMessage(message) ||
    isHighDemandMessage(message)
  );
}

function isFallbackEligible(status: number, message: string, code: string): boolean {
  return (
    isRetryableError(status, message, code) ||
    status === 404 ||
    code === "MODEL_NOT_FOUND" ||
    code === "INCOMPLETE_JSON"
  );
}

function mapGeminiError(error: unknown): GeminiError {
  if (error instanceof GeminiError) {
    return error;
  }

  const geminiError = error as Error & {
    status?: number;
    code?: string;
    cause?: unknown;
  };

  const rawMessage = geminiError.message ?? "Gemini analysis failed.";
  const message = extractGeminiErrorMessage(rawMessage);
  const status = geminiError.status ?? 500;

  if (
    rawMessage.toLowerCase().includes("fetch failed") ||
    geminiError.name === "TypeError"
  ) {
    const cause = geminiError.cause as Error | undefined;
    return new GeminiError(
      `Gemini API connection failed${cause?.message ? `: ${cause.message}` : ". Check your API key and network connection."}`,
      503,
      "GEMINI_NETWORK_ERROR",
    );
  }

  if (isQuotaOrRateLimitError(status, message)) {
    return new GeminiError(
      message.includes("quota")
        ? message
        : "Gemini API rate limit exceeded. Please wait and try again.",
      429,
      "RATE_LIMIT",
    );
  }

  if (status === 503 || isHighDemandMessage(message)) {
    return new GeminiError(message, 503, "MODEL_OVERLOADED");
  }

  if (status === 403 || status === 401) {
    return new GeminiError(
      "Gemini API authentication failed. Check your GEMINI_API_KEY in .env.local.",
      status,
      "AUTH_FAILED",
    );
  }

  if (status === 404) {
    return new GeminiError(
      `Gemini model not found or unavailable: ${message}`,
      404,
      "MODEL_NOT_FOUND",
    );
  }

  if (isIncompleteJsonMessage(message) || geminiError.code === "INCOMPLETE_JSON") {
    return new GeminiError(
      "Gemini ran out of output space before finishing the ticket JSON.",
      502,
      "INCOMPLETE_JSON",
    );
  }

  return new GeminiError(message, status, geminiError.code ?? "GEMINI_ERROR");
}

function asTrimmedString(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() ? value.trim() : fallback;
}

function looksLikeBug(text: string): boolean {
  return /\b(bug|defect|broken|crash(?:es|ed|ing)?|error|incorrect|wrong|fails?|failure|fix|hotfix|regression|not working|doesn'?t work|does not work|issue)\b/i.test(
    text,
  );
}

function asTaskType(
  value: unknown,
  title: string,
  description: string,
  task: string,
): ActionItemType {
  const declared = typeof value === "string" ? value.trim().toLowerCase() : "";
  const inferredBug = looksLikeBug(`${title} ${task} ${description}`);

  if (declared === "bug" || inferredBug) {
    return "Bug";
  }

  return "Task";
}

function shortenToTitle(text: string): string {
  const firstLine = text
    .split(/\n/)[0]
    ?.replace(/^[\d.)\-\s]+/, "")
    .trim() ?? "";
  const words = firstLine.split(/\s+/).filter(Boolean).slice(0, 10);
  let title = words.join(" ");

  if (title.length > 72) {
    title = title.slice(0, 72);
    const lastSpace = title.lastIndexOf(" ");
    if (lastSpace > 40) {
      title = title.slice(0, lastSpace);
    }
  }

  return title.replace(/[.:]+$/, "").trim();
}

function separateTitleAndDescription(
  title: string,
  description: string,
  task: string,
): { title: string; description: string } {
  const compactTitle = title.replace(/\s+/g, " ").trim();
  const compactDescription = description.trim();
  const normalizedTitle = compactTitle.toLowerCase();
  const normalizedDescription = compactDescription.replace(/\s+/g, " ").toLowerCase();

  let nextTitle = compactTitle;
  let nextDescription = compactDescription || task || compactTitle;

  const sameAsTitle =
    !compactDescription ||
    normalizedDescription === normalizedTitle ||
    normalizedDescription === task.replace(/\s+/g, " ").toLowerCase();

  if (sameAsTitle) {
    nextDescription = compactDescription || task || compactTitle;
    nextTitle = shortenToTitle(task && task !== compactTitle ? task : nextDescription);
  }

  if (nextTitle.toLowerCase() === nextDescription.replace(/\s+/g, " ").toLowerCase()) {
    nextTitle = shortenToTitle(nextDescription);
  }

  if (!nextTitle) {
    nextTitle = shortenToTitle(nextDescription) || "Follow-up item";
  }

  return { title: nextTitle, description: nextDescription };
}

function asPriority(value: unknown): ActionItemPriority {
  const normalized = typeof value === "string" ? value.trim().toLowerCase() : "";
  if (normalized === "high" || normalized === "critical" || normalized === "urgent") {
    return "High";
  }
  if (normalized === "low") {
    return "Low";
  }
  return "Medium";
}

function parseEvidence(value: unknown): ActionItemEvidence[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.flatMap((raw) => {
    if (typeof raw !== "object" || raw === null) {
      return [];
    }
    const item = raw as Record<string, unknown>;
    const quote = asTrimmedString(item.quote, "");
    if (!quote) {
      return [];
    }
    return [
      {
        timestamp: asTrimmedString(item.timestamp, ""),
        speaker: asTrimmedString(item.speaker, ""),
        quote,
      },
    ];
  });
}

function extractJsonText(text: string): string {
  const fenced = text.match(/```json\s*([\s\S]*?)\s*```/i);
  if (fenced?.[1]) {
    return fenced[1].trim();
  }

  const trimmed = text.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed;
  }

  const start = trimmed.indexOf("{");
  const end = trimmed.lastIndexOf("}");
  if (start >= 0 && end > start) {
    return trimmed.slice(start, end + 1);
  }

  return trimmed;
}

function repairIncompleteJson(text: string): string | null {
  const start = text.indexOf("{");
  if (start < 0) {
    return null;
  }

  let source = text.slice(start);
  if (source.endsWith("\\")) {
    source = source.slice(0, -1);
  }

  const stack: Array<"{" | "["> = [];
  let inString = false;
  let escaped = false;

  for (const character of source) {
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (character === "\\") {
        escaped = true;
        continue;
      }
      if (character === '"') {
        inString = false;
      }
      continue;
    }

    if (character === '"') {
      inString = true;
      continue;
    }
    if (character === "{" || character === "[") {
      stack.push(character);
      continue;
    }
    if (character === "}" || character === "]") {
      stack.pop();
    }
  }

  if (escaped) {
    source += " ";
  }
  if (inString) {
    source += '"';
  }
  while (stack.length > 0) {
    const open = stack.pop();
    source += open === "{" ? "}" : "]";
  }

  try {
    JSON.parse(source);
    return source;
  } catch {
    return null;
  }
}

function mapRawActionItem(raw: unknown): ParsedActionItem | null {
  if (typeof raw !== "object" || raw === null) {
    return null;
  }

  const item = raw as Record<string, unknown>;
  const task = asTrimmedString(item.task, "");
  const rawTitle = asTrimmedString(item.title, "");
  const rawDescription = asTrimmedString(item.description, "");

  if (!rawTitle && !task && !rawDescription) {
    return null;
  }

  const separated = separateTitleAndDescription(
    rawTitle,
    rawDescription,
    task,
  );

  return {
    title: separated.title,
    type: asTaskType(item.type, separated.title, separated.description, task),
    status: asTrimmedString(item.status, "New"),
    priority: asPriority(item.priority),
    pointPerson: asTrimmedString(item.pointPerson, "Unassigned"),
    task: task || separated.title,
    deadline: asTrimmedString(item.deadline, "TBD"),
    devTag: normalizeDevTag(asTrimmedString(item.devTag, "None")),
    relatedTo: asTrimmedString(item.relatedTo, "None"),
    description: separated.description,
    evidence: parseEvidence(item.evidence),
  };
}

function parseGeminiJson(
  text: string,
): Omit<GeminiAnalysisResult, "modelUsed"> {
  const extracted = extractJsonText(text);
  let jsonText = extracted;

  try {
    JSON.parse(jsonText);
  } catch {
    jsonText = repairIncompleteJson(text) ?? extracted;
  }

  try {
    const parsed = JSON.parse(jsonText) as {
      actionItems?: unknown;
      analysis?: unknown;
    };
    const rawItems = Array.isArray(parsed.actionItems)
      ? parsed.actionItems
      : [];
    const actionItems = rawItems
      .map(mapRawActionItem)
      .filter((item): item is ParsedActionItem => item !== null);
    const analysisFromModel =
      typeof parsed.analysis === "string" && parsed.analysis.trim()
        ? parsed.analysis.trim()
        : "";
    const analysis =
      analysisFromModel ||
      (actionItems.length > 0
        ? `## Executive Summary\nGemini finished ${actionItems.length} tickets but the written analysis was cut off. Review the tickets below.`
        : text.replace(extracted, "").trim());

    return { actionItems, analysis };
  } catch {
    return { actionItems: [], analysis: text.trim() };
  }
}

interface GenerateOptions {
  onChunk?: (text: string) => void;
  compact?: boolean;
  responseSchema?: object;
}

interface AnalysisGenerationConfig {
  systemInstruction: string;
  responseMimeType: "application/json";
  maxOutputTokens: number;
  responseSchema?: object;
}

function getGenerationConfig(
  systemInstruction: string,
  compact: boolean,
  responseSchema?: object,
): AnalysisGenerationConfig {
  const config: AnalysisGenerationConfig = {
    systemInstruction,
    responseMimeType: "application/json",
    maxOutputTokens: MAX_OUTPUT_TOKENS,
  };
  if (responseSchema) {
    config.responseSchema = responseSchema;
  } else if (!compact) {
    config.responseSchema = ANALYSIS_RESPONSE_SCHEMA;
  }
  return config;
}

async function streamModelText(
  model: string,
  contents: ContentListUnion,
  systemInstruction: string,
  options: GenerateOptions = {},
): Promise<string> {
  const ai = getGeminiClient();
  const request = {
    model,
    contents,
    config: getGenerationConfig(
      systemInstruction,
      options.compact === true,
      options.responseSchema,
    ),
  };
  const models = ai.models as unknown as {
    generateContentStream?: (
      payload: typeof request,
    ) => Promise<AsyncIterable<{ text?: string }>>;
    generateContent: (payload: typeof request) => Promise<{ text?: string }>;
  };

  if (typeof models.generateContentStream === "function") {
    let accumulated = "";
    try {
      const stream = await models.generateContentStream(request);
      for await (const chunk of stream) {
        const text = chunk.text ?? "";
        if (!text) {
          continue;
        }
        accumulated += text;
        options.onChunk?.(text);
      }
      return accumulated;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (isIncompleteJsonMessage(message) && accumulated.trim()) {
        return accumulated;
      }
      throw error;
    }
  }

  const response = await models.generateContent(request);
  const raw = response.text ?? "";
  if (raw) {
    options.onChunk?.(raw);
  }
  return raw;
}

async function generateAnalysis(
  model: string,
  contents: ContentListUnion,
  systemInstruction: string,
  options: GenerateOptions = {},
): Promise<Omit<GeminiAnalysisResult, "modelUsed">> {
  let raw = "";
  try {
    raw = await streamModelText(model, contents, systemInstruction, options);
  } catch (error) {
    const mapped = mapGeminiError(error);
    if (mapped.code !== "INCOMPLETE_JSON" || !raw.trim()) {
      throw mapped;
    }
  }

  if (!raw.trim()) {
    throw new GeminiError(
      "Gemini returned an empty analysis response.",
      502,
      "EMPTY_RESPONSE",
    );
  }

  const parsed = parseGeminiJson(raw);
  if (parsed.actionItems.length === 0 && !parsed.analysis.trim()) {
    throw new GeminiError(
      "Gemini ran out of output space before finishing the ticket JSON.",
      502,
      "INCOMPLETE_JSON",
    );
  }
  return parsed;
}

async function generateAnalysisWithRetries(
  model: string,
  contents: ContentListUnion,
  systemInstruction: string,
  options: GenerateOptions = {},
): Promise<GeminiAnalysisResult> {
  let lastError: GeminiError | null = null;
  let compact = options.compact === true;

  for (let attempt = 0; attempt < MAX_RETRIES_PER_MODEL; attempt++) {
    const instruction = compact
      ? `${systemInstruction}\n${COMPACT_OUTPUT_RULES}`
      : `${systemInstruction}\n${OUTPUT_BUDGET_RULES}`;

    try {
      const result = await generateAnalysis(model, contents, instruction, {
        ...options,
        compact,
      });
      return { ...result, modelUsed: model };
    } catch (error) {
      const mapped = mapGeminiError(error);
      lastError = mapped;

      if (mapped.code === "INCOMPLETE_JSON") {
        compact = true;
        if (attempt < MAX_RETRIES_PER_MODEL - 1) {
          await sleep(INITIAL_RETRY_DELAY_MS * 2 ** attempt);
          continue;
        }
      }

      if (!isRetryableError(mapped.statusCode, mapped.message, mapped.code)) {
        throw mapped;
      }

      if (shouldFailFastToNextModel(mapped)) {
        throw mapped;
      }

      if (attempt < MAX_RETRIES_PER_MODEL - 1) {
        await sleep(INITIAL_RETRY_DELAY_MS * 2 ** attempt);
      }
    }
  }

  throw (
    lastError ??
    new GeminiError("Gemini analysis failed.", 503, "MODEL_OVERLOADED")
  );
}

async function runModelTier(
  contents: ContentListUnion,
  systemInstruction: string,
  options: GenerateOptions = {},
): Promise<GeminiAnalysisResult> {
  const models = getOrderedModels();
  let lastError: GeminiError | null = null;

  for (let index = 0; index < models.length; index++) {
    const model = models[index];

    try {
      return await generateAnalysisWithRetries(
        model,
        contents,
        systemInstruction,
        options,
      );
    } catch (error) {
      const mapped = mapGeminiError(error);
      lastError = mapped;

      const hasLowerTierModel = index < models.length - 1;
      if (
        !hasLowerTierModel ||
        !isFallbackEligible(mapped.statusCode, mapped.message, mapped.code)
      ) {
        throw mapped;
      }
    }
  }

  throw (
    lastError ??
    new GeminiError(
      "All Gemini models are busy. Wait a minute and try again.",
      503,
      "MODEL_OVERLOADED",
    )
  );
}

export interface MeetingAnalysisSource {
  transcript: string;
  chat?: string;
  meetingDate?: string;
  artifactDateLabel?: string;
  artifactSource?: string;
}

function buildMeetingAnalysisPrompt(source: MeetingAnalysisSource): string {
  const meetingDate = source.meetingDate?.trim() || "this meeting day";
  const artifactDate =
    source.artifactDateLabel?.trim() || meetingDate;
  const artifactSource = source.artifactSource?.trim() || "meeting start";
  const transcript =
    source.transcript.trim() || "(No spoken transcript was available.)";
  const chat =
    source.chat?.trim() ||
    "(No meeting chat messages were available for that KSA day.)";

  return `Analyze this Teams meeting comprehensively using BOTH sources.

Timezone: Arabia Standard Time (KSA, Asia/Riyadh, UTC+3).
Calendar event: ${meetingDate}
Comparison day: ${artifactDate}
Day source: KSA calendar date when the ${artifactSource} was created.

Read the spoken transcript AND the meeting chat. Compare them for redundancy on that KSA day only.
- Extract work from both sources.
- If the same action appears in both, create ONE ticket and say it was repeated.
- Still capture chat-only and transcript-only work.
- Ignore chat from other KSA dates; it has already been filtered out when possible.
- Identify each item as Task or Bug.
- Keep title and description different: the title is a short headline for the description.
- Ground each ticket with evidence quotes and timestamps from lines like "[00:12:04] Speaker: ...".
- If a chat line includes a markdown [title](url) permalink or image link and the ticket comes from that message or image, put that same [title](url) at the end of the description. Use teams.cloud.microsoft/l/message links only, never launcher.html.

--- Spoken transcript ---
${transcript}

--- Meeting chat (same KSA day) ---
${chat}`;
}

export async function analyzeTranscript(
  source: MeetingAnalysisSource,
  options: GenerateOptions = {},
): Promise<GeminiAnalysisResult> {
  const result = await runModelTier(
    buildMeetingAnalysisPrompt(source),
    SYSTEM_PROMPT,
    options,
  );
  return {
    ...result,
    actionItems: attachChatSourceLinks(result.actionItems, source.chat ?? ""),
  };
}

const FILE_ACTIVE_POLL_INTERVAL_MS = 3000;
const FILE_ACTIVE_TIMEOUT_MS = 5 * 60 * 1000;

async function uploadVideoAndWaitActive(video: {
  data: Buffer;
  mimeType: string;
}): Promise<{ uri: string; mimeType: string; name?: string }> {
  const ai = getGeminiClient();

  let uploaded;
  try {
    const blob = new Blob([new Uint8Array(video.data)], {
      type: video.mimeType,
    });
    uploaded = await ai.files.upload({
      file: blob,
      config: { mimeType: video.mimeType },
    });
  } catch (error) {
    throw mapGeminiError(error);
  }

  const fileName = uploaded.name;
  if (!fileName) {
    throw new GeminiError(
      "Gemini did not return a file handle for the uploaded recording.",
      502,
      "VIDEO_UPLOAD_FAILED",
    );
  }

  const deadline = Date.now() + FILE_ACTIVE_TIMEOUT_MS;
  let current = uploaded;

  while (current.state === FileState.PROCESSING) {
    if (Date.now() > deadline) {
      throw new GeminiError(
        "Timed out waiting for Gemini to process the recording video.",
        504,
        "VIDEO_PROCESSING_TIMEOUT",
      );
    }

    await sleep(FILE_ACTIVE_POLL_INTERVAL_MS);
    current = await ai.files.get({ name: fileName });
  }

  if (current.state === FileState.FAILED) {
    throw new GeminiError(
      current.error?.message ?? "Gemini failed to process the recording video.",
      502,
      "VIDEO_PROCESSING_FAILED",
    );
  }

  if (!current.uri || !current.mimeType) {
    throw new GeminiError(
      "Uploaded recording is missing a usable file URI.",
      502,
      "VIDEO_UPLOAD_FAILED",
    );
  }

  return { uri: current.uri, mimeType: current.mimeType, name: fileName };
}

export async function analyzeTranscriptWithVideo(
  source: MeetingAnalysisSource,
  video: { data: Buffer; mimeType: string },
  options: GenerateOptions = {},
): Promise<GeminiAnalysisResult> {
  const file = await uploadVideoAndWaitActive(video);

  try {
    const contents = createUserContent([
      createPartFromUri(file.uri, file.mimeType),
      `${buildMeetingAnalysisPrompt(source)}\n\nA meeting recording is also attached. Use it with the transcript and chat. If the transcript is empty, rely on the recording and chat.`,
    ]);

    const result = await runModelTier(contents, VIDEO_SYSTEM_PROMPT, options);
    return {
      ...result,
      actionItems: attachChatSourceLinks(result.actionItems, source.chat ?? ""),
    };
  } finally {
    if (file.name) {
      // Best-effort cleanup; ignore failures.
      try {
        await getGeminiClient().files.delete({ name: file.name });
      } catch {
        // Uploaded files auto-expire after 48h if delete fails.
      }
    }
  }
}

export function getActiveGeminiModelLabel(): string {
  const models = getOrderedModels();
  if (models.length === 1) {
    return models[0];
  }
  return `${models[0]} → ${models[models.length - 1]}`;
}

const CODE_QC_RESPONSE_SCHEMA = {
  type: "OBJECT",
  properties: {
    tickets: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          title: { type: "STRING" },
          type: { type: "STRING" },
          priority: { type: "STRING" },
          issueIds: { type: "ARRAY", items: { type: "STRING" } },
          reason: { type: "STRING" },
        },
        required: ["title", "issueIds"],
      },
    },
  },
  required: ["tickets"],
} as const;

const CODE_QC_SYSTEM_PROMPT = `You are grouping SonarQube findings into PMT work packets for a software team.

Return ONE JSON object only:
{
  "tickets": [
    {
      "title": "Fix hardcoded token in supabaseClient.js",
      "type": "Bug",
      "priority": "High",
      "issueIds": ["A1", "A2"],
      "reason": "Same secret-detection findings in one file"
    }
  ]
}

Rules:
- Group issues that a developer would fix in one sitting: same file, or the same rule across the repo when it is a cleanup (prefer-const, unused vars).
- Keep BLOCKER, CRITICAL, and VULNERABILITY findings on their own ticket unless they are clearly the same defect.
- Distinct bugs stay separate. Do not dump the whole repo into one ticket.
- Every input id must appear in exactly one ticket. Do not invent ids.
- Prefer 8–40 tickets. Never more than 40.
- title: 4–12 words, outcome-focused. No issue ids in the title.
- type: "Bug" or "Task". Vulnerabilities and broken behavior are Bug. Code smells and cleanups are Task.
- priority: "High", "Medium", or "Low" from the worst severity in the group (BLOCKER/CRITICAL = High, MAJOR = Medium, else Low).
- issueIds: the A1/A2 ids from the catalog, not Sonar keys.
- reason: one short sentence.`;

const CODE_QC_COMPACT_RULES = `
The previous attempt was truncated. Return COMPLETE valid JSON only.
- At most 20 tickets covering the highest-severity work first.
- Still assign every remaining id to some ticket (file-level groups are fine).
- Finish the JSON object.`;

export interface CodeQcGeminiTicket {
  title: string;
  type: "Task" | "Bug";
  priority: "High" | "Medium" | "Low";
  issueKeys: string[];
  reason: string;
}

function parseCodeQcTickets(text: string): CodeQcGeminiTicket[] {
  const extracted = extractJsonText(text);
  let jsonText = extracted;
  try {
    JSON.parse(jsonText);
  } catch {
    jsonText = repairIncompleteJson(text) ?? extracted;
  }

  const parsed = JSON.parse(jsonText) as { tickets?: unknown };
  if (!Array.isArray(parsed.tickets)) {
    return [];
  }

  return parsed.tickets.flatMap((raw) => {
    if (typeof raw !== "object" || raw === null) {
      return [];
    }
    const item = raw as Record<string, unknown>;
    const title = asTrimmedString(item.title, "");
    const ids = Array.isArray(item.issueIds)
      ? item.issueIds.filter((id): id is string => typeof id === "string" && id.trim().length > 0)
      : [];
    if (!title || ids.length === 0) {
      return [];
    }
    const declared = asTrimmedString(item.type, "Task");
    return [
      {
        title,
        type: declared.toLowerCase() === "bug" ? "Bug" : "Task",
        priority: asPriority(item.priority),
        issueKeys: ids.map((id) => id.trim().toUpperCase()),
        reason: asTrimmedString(item.reason, ""),
      },
    ];
  });
}

function codeQcCatalog(issues: CodeQcIssue[]): { lines: string[]; byId: Map<string, CodeQcIssue> } {
  const byId = new Map<string, CodeQcIssue>();
  const lines = issues.map((issue, index) => {
    const id = `A${index + 1}`;
    byId.set(id, issue);
    const line = typeof issue.line === "number" ? `:${issue.line}` : "";
    const message = issue.message.replace(/\s+/g, " ").trim().slice(0, 140);
    return `${id}\t${issue.type}\t${issue.severity}\t${issue.rule ?? "-"}\t${issue.component}${line}\t${message}`;
  });
  return { lines, byId };
}

export async function analyzeCodeQcIssues(
  issues: CodeQcIssue[],
): Promise<{ tickets: CodeQcGeminiTicket[]; modelUsed: string }> {
  if (issues.length === 0) {
    return { tickets: [], modelUsed: getOrderedModels()[0] ?? DEFAULT_MODEL };
  }

  const { lines, byId } = codeQcCatalog(issues);
  const prompt = `Group these SonarQube issues from the repo into PMT tickets.

Catalog columns: id, type, severity, rule, file, message.
Refer to issues only by id.

--- Issues ---
${lines.join("\n")}`;

  const models = getOrderedModels();
  let lastError: GeminiError | null = null;

  for (let index = 0; index < models.length; index++) {
    const model = models[index];
    let compact = false;
    for (let attempt = 0; attempt < MAX_RETRIES_PER_MODEL; attempt++) {
      const instruction = compact
        ? `${CODE_QC_SYSTEM_PROMPT}\n${CODE_QC_COMPACT_RULES}`
        : `${CODE_QC_SYSTEM_PROMPT}\n${OUTPUT_BUDGET_RULES}`;
      try {
        const raw = await streamModelText(model, prompt, instruction, {
          compact,
          responseSchema: CODE_QC_RESPONSE_SCHEMA,
        });
        if (!raw.trim()) {
          throw new GeminiError("Gemini returned an empty CODE QC grouping.", 502, "EMPTY_RESPONSE");
        }
        let parsed: CodeQcGeminiTicket[] = [];
        try {
          parsed = parseCodeQcTickets(raw);
        } catch {
          throw new GeminiError(
            "Gemini ran out of output space before finishing the CODE QC tickets.",
            502,
            "INCOMPLETE_JSON",
          );
        }
        const byKey = new Map(issues.map((issue) => [issue.key, issue]));
        const tickets = parsed
          .map((ticket) => ({
            ...ticket,
            issueKeys: [
              ...new Set(
                ticket.issueKeys.flatMap((id) => {
                  const fromCatalog = byId.get(id)?.key;
                  if (fromCatalog) {
                    return [fromCatalog];
                  }
                  return byKey.has(id) ? [id] : [];
                }),
              ),
            ],
          }))
          .filter((ticket) => ticket.issueKeys.length > 0);
        if (tickets.length === 0) {
          throw new GeminiError(
            "Gemini ran out of output space before finishing the CODE QC tickets.",
            502,
            "INCOMPLETE_JSON",
          );
        }
        return { tickets, modelUsed: model };
      } catch (error) {
        const mapped = mapGeminiError(error);
        lastError = mapped;
        if (mapped.code === "INCOMPLETE_JSON") {
          compact = true;
          if (attempt < MAX_RETRIES_PER_MODEL - 1) {
            await sleep(INITIAL_RETRY_DELAY_MS * 2 ** attempt);
            continue;
          }
        }
        if (!isRetryableError(mapped.statusCode, mapped.message, mapped.code)) {
          throw mapped;
        }
        if (shouldFailFastToNextModel(mapped) || index === models.length - 1) {
          if (index === models.length - 1) {
            throw mapped;
          }
          break;
        }
      }
    }
  }

  throw (
    lastError ??
    new GeminiError("Gemini could not group the CODE QC issues.", 503, "MODEL_OVERLOADED")
  );
}
