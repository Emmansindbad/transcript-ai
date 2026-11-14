"use client";

import { useState } from "react";
import { quietButtonClass } from "@/components/ui-classes";

function formatFreed(bytes: number): string {
  if (bytes >= 1024 * 1024) {
    return `${Math.round(bytes / (1024 * 1024))} MB`;
  }
  if (bytes >= 1024) {
    return `${Math.round(bytes / 1024)} KB`;
  }
  return `${bytes} B`;
}

export function ClearCacheButton(): React.ReactNode {
  const [clearing, setClearing] = useState(false);

  async function handleClear(): Promise<void> {
    const confirmed = window.confirm(
      "Clear the saved meeting list and people list? The next Fetch Data loads them again from Microsoft 365. You stay signed in to Teams. The page will reload.",
    );
    if (!confirmed) {
      return;
    }

    setClearing(true);
    try {
      const response = await fetch("/api/cache", { method: "POST" });
      const data = (await response.json()) as { freedBytes?: number; error?: string };
      if (!response.ok) {
        window.alert(data.error ?? "Could not clear the cache.");
        setClearing(false);
        return;
      }
      const freed = typeof data.freedBytes === "number" ? formatFreed(data.freedBytes) : "";
      window.alert(freed ? `Cleared ${freed}. The page will reload.` : "Cache cleared. The page will reload.");
      window.location.reload();
    } catch {
      window.alert("Could not clear the cache.");
      setClearing(false);
    }
  }

  return (
    <button
      type="button"
      onClick={() => {
        void handleClear();
      }}
      disabled={clearing}
      title="Clear the saved meeting list and people list"
      aria-label="Clear cache"
      className={quietButtonClass}
    >
      {clearing ? "Clearing…" : "Clear cache"}
    </button>
  );
}
