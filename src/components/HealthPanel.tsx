"use client";

import { useState } from "react";
import type { GraphHealthReport, ProbeStatus } from "@/types/meeting";
import { quietButtonClass } from "@/components/ui-classes";

function dotClass(status: ProbeStatus): string {
  switch (status) {
    case "ok":
      return "bg-emerald-500";
    case "missing":
      return "bg-amber-500";
    case "denied":
    case "failed":
      return "bg-red-500";
    default:
      return "bg-zinc-300";
  }
}

function ProbeChip({
  label,
  status,
  detail,
}: {
  label: string;
  status: ProbeStatus;
  detail: string;
}): React.ReactNode {
  return (
    <span title={detail} className="inline-flex items-center gap-1.5 text-xs text-zinc-600">
      <span className={`h-1.5 w-1.5 rounded-full ${dotClass(status)}`} />
      {label}
    </span>
  );
}

export function HealthPanel(): React.ReactNode {
  const [report, setReport] = useState<GraphHealthReport | null>(null);
  const [open, setOpen] = useState(false);
  const [refreshing, setRefreshing] = useState(false);

  async function loadHealth(refresh: boolean): Promise<void> {
    setRefreshing(true);
    try {
      const response = await fetch(refresh ? "/api/health?refresh=1" : "/api/health");
      const data = (await response.json()) as GraphHealthReport;
      if (response.ok) {
        setReport(data);
      }
    } catch {
      // Health is advisory; the dashboard still works.
    } finally {
      setRefreshing(false);
    }
  }

  function toggleHealth(): void {
    const next = !open;
    setOpen(next);
    if (next && !report && !refreshing) {
      void loadHealth(false);
    }
  }

  const probes = report?.probes;

  return (
    <div className="relative">
      <button type="button" onClick={toggleHealth} className={quietButtonClass}>
        Health
      </button>
      {open ? (
        <div className="fixed top-14 right-4 z-30 w-80 max-w-[calc(100vw-2rem)] rounded-md border border-zinc-200 bg-white p-4 shadow-sm">
          {report && probes ? (
            <div className="space-y-3">
              <div className="flex flex-wrap gap-x-3 gap-y-2">
                <ProbeChip label="Token" status={report.token} detail="Graph token" />
                <ProbeChip
                  label="Calendar"
                  status={probes.calendar.status}
                  detail={probes.calendar.detail}
                />
                <ProbeChip
                  label="Transcripts"
                  status={probes.transcripts.status}
                  detail={probes.transcripts.detail}
                />
                <ProbeChip
                  label="Recordings"
                  status={probes.recordings.status}
                  detail={probes.recordings.detail}
                />
                <ProbeChip label="Chat" status={probes.chat.status} detail={probes.chat.detail} />
                <ProbeChip label="Mail" status={probes.mail.status} detail={probes.mail.detail} />
                <ProbeChip label="To Do" status={probes.todo.status} detail={probes.todo.detail} />
              </div>
              <p className="break-all font-mono text-[11px] text-zinc-400">{report.clientId}</p>
              {report.missingRoles.length > 0 ? (
                <p className="text-xs text-amber-800">
                  Missing from this app&apos;s token: {report.missingRoles.join(", ")}
                </p>
              ) : null}
              {(report.hints.length > 0 ? report.hints : report.hint ? [report.hint] : []).map(
                (item) => (
                  <p key={item} className="text-xs leading-5 text-zinc-600">
                    {item}
                  </p>
                ),
              )}
              <p className="text-xs text-zinc-500">Emails go to {report.summaryEmailRecipient}</p>
              <button
                type="button"
                onClick={() => void loadHealth(true)}
                disabled={refreshing}
                className="text-xs font-medium text-zinc-700 hover:text-zinc-900 disabled:opacity-60"
              >
                {refreshing ? "Refreshing…" : "Refresh token"}
              </button>
            </div>
          ) : (
            <p className="text-xs text-zinc-500">
              {refreshing ? "Checking Graph…" : "Graph health is not loaded yet."}
            </p>
          )}
        </div>
      ) : null}
    </div>
  );
}
