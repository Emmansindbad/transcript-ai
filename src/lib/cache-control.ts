import "server-only";

import { readdir, rm, stat } from "node:fs/promises";
import path from "node:path";
import { clearGraphDataCaches } from "@/lib/m365";
import { clearRememberedGroupChats } from "@/lib/teams-message";
import { releaseDelegatedClient } from "@/lib/teams-auth";

const DATA_DIR = path.join(process.cwd(), "data");

const DATA_CACHE_FILES = ["meetings-cache.json", "users-cache.json"];

export interface CacheClearResult {
  freedBytes: number;
  clearedBuildCache: boolean;
}

async function removeTree(target: string): Promise<number> {
  let freed = 0;
  let info: Awaited<ReturnType<typeof stat>>;
  try {
    info = await stat(target);
  } catch {
    return 0;
  }
  if (info.isDirectory()) {
    const entries = await readdir(target, { withFileTypes: true });
    for (const entry of entries) {
      freed += await removeTree(path.join(target, entry.name));
    }
    try {
      await rm(target, { force: true });
    } catch {
      // A file still in use stays until the next restart.
    }
    return freed;
  }
  try {
    await rm(target, { force: true });
    return info.size;
  } catch {
    return 0;
  }
}

export async function clearAnalyzerCaches(): Promise<CacheClearResult> {
  clearGraphDataCaches();
  releaseDelegatedClient();
  await clearRememberedGroupChats();

  let freedBytes = 0;
  for (const name of DATA_CACHE_FILES) {
    freedBytes += await removeTree(path.join(DATA_DIR, name));
  }

  return { freedBytes, clearedBuildCache: false };
}
