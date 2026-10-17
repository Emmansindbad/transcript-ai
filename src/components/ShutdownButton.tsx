"use client";

import { useState } from "react";

export function ShutdownButton(): React.ReactNode {
  const [stopping, setStopping] = useState(false);
  const [stopped, setStopped] = useState(false);

  async function handleEnd(): Promise<void> {
    const confirmed = window.confirm(
      "Shut down the analyzer on this PC? Close this tab afterward. Start it again from the desktop shortcut.",
    );
    if (!confirmed) {
      return;
    }

    setStopping(true);
    try {
      await fetch("/api/shutdown", { method: "POST" });
    } catch {
      // The server closes the connection as it exits.
    }
    setStopped(true);
    setStopping(false);
  }

  if (stopped) {
    return (
      <p className="max-w-xs text-right text-xs font-medium text-slate-600">
        Analyzer stopped. Close this tab. Use the desktop shortcut to start it
        again.
      </p>
    );
  }

  return (
    <button
      type="button"
      onClick={() => {
        void handleEnd();
      }}
      disabled={stopping}
      title="Shut down the analyzer on this PC"
      aria-label="Shut down the analyzer on this PC"
      className="shrink-0 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm font-medium text-red-800 hover:bg-red-100 disabled:cursor-not-allowed disabled:opacity-60"
    >
      {stopping ? "Stopping…" : "End"}
    </button>
  );
}
