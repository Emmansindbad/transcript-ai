import "server-only";

import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import type { MeetingActionItem, SavedAnalysis } from "@/types/meeting";

const STORE_VERSION = 1;
const STORE_DIR = path.join(process.cwd(), "data");
const STORE_PATH = path.join(STORE_DIR, "analyzer-store.json");

interface AnalyzerStoreFile {
  version: number;
  analyses: Record<string, SavedAnalysis>;
  filedKeys: string[];
}

let writeChain: Promise<void> = Promise.resolve();

function emptyStore(): AnalyzerStoreFile {
  return { version: STORE_VERSION, analyses: {}, filedKeys: [] };
}

async function readStore(): Promise<AnalyzerStoreFile> {
  try {
    const raw = await readFile(STORE_PATH, "utf8");
    const parsed = JSON.parse(raw) as Partial<AnalyzerStoreFile>;
    return {
      version: STORE_VERSION,
      analyses:
        parsed.analyses && typeof parsed.analyses === "object"
          ? parsed.analyses
          : {},
      filedKeys: Array.isArray(parsed.filedKeys)
        ? parsed.filedKeys.filter((key): key is string => typeof key === "string")
        : [],
    };
  } catch {
    return emptyStore();
  }
}

async function writeStore(store: AnalyzerStoreFile): Promise<void> {
  await mkdir(STORE_DIR, { recursive: true });
  await writeFile(STORE_PATH, JSON.stringify(store, null, 2), "utf8");
}

function withLock<T>(fn: () => Promise<T>): Promise<T> {
  const run = writeChain.then(fn, fn);
  writeChain = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

export async function getSavedAnalysis(
  meetingKey: string,
): Promise<SavedAnalysis | null> {
  const store = await readStore();
  return store.analyses[meetingKey] ?? null;
}

export async function saveAnalysis(record: SavedAnalysis): Promise<SavedAnalysis> {
  return withLock(async () => {
    const store = await readStore();
    const filed = new Set(store.filedKeys);
    const actionItems = applyFiledState(record.actionItems, filed);
    const saved: SavedAnalysis = {
      ...record,
      savedAt: new Date().toISOString(),
      actionItems,
    };
    store.analyses[record.meetingKey] = saved;
    await writeStore(store);
    return saved;
  });
}

export async function listFiledKeys(): Promise<Set<string>> {
  const store = await readStore();
  return new Set(store.filedKeys);
}

export async function markItemsFiled(filingKeys: string[]): Promise<string[]> {
  const unique = [...new Set(filingKeys.filter(Boolean))];
  if (unique.length === 0) {
    return [];
  }

  return withLock(async () => {
    const store = await readStore();
    const filed = new Set(store.filedKeys);
    for (const key of unique) {
      filed.add(key);
    }
    store.filedKeys = [...filed];

    for (const analysis of Object.values(store.analyses)) {
      analysis.actionItems = analysis.actionItems.map((item) =>
        unique.includes(item.filingKey)
          ? { ...item, alreadyFiled: true, include: false }
          : item,
      );
    }

    await writeStore(store);
    return unique;
  });
}

export function applyFiledState(
  items: MeetingActionItem[],
  filedKeys: Set<string>,
): MeetingActionItem[] {
  return items.map((item) => ({
    ...item,
    alreadyFiled: item.alreadyFiled || filedKeys.has(item.filingKey),
  }));
}
