"use client";

import { useState } from "react";
import type { GraphHealthReport, ProbeStatus } from "@/types/meeting";

function statusClass(status: ProbeStatus): string {
  switch (status) {
    case "ok":
      return "border-emerald-200 bg-emerald-50 text-emerald-800";
    case "denied":
    case "failed":
      return "border-red-200 bg-red-50 text-red-800";
    case "missing":
      return "border-amber-200 bg-amber-50 text-amber-900";
    default:
      return "border-slate-200 bg-slate-50 text-slate-600";
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
    <span
      title={detail}
      className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[11px] font-medium ${statusClass(status)}`}
    >
      {label}: {status}
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
    <div className="mt-3">
      <button
        type="button"
        onClick={toggleHealth}
        className="text-xs font-medium text-slate-500 hover:text-slate-800"
      >
        Graph health {open ? "▾" : "▸"}
      </button>
      {open && report && probes ? (
        <div className="mt-2 space-y-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2">
          <div className="flex flex-wrap gap-1.5">
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
            <ProbeChip
              label="Chat"
              status={probes.chat.status}
              detail={probes.chat.detail}
            />
            <ProbeChip label="Mail" status={probes.mail.status} detail={probes.mail.detail} />
            <ProbeChip label="To Do" status={probes.todo.status} detail={probes.todo.detail} />
          </div>
          <p className="text-xs text-slate-600">
            App registration:{" "}
            <span className="font-mono text-[11px] text-slate-800">{report.clientId}</span>
          </p>
          {report.missingRoles.length > 0 ? (
            <p className="text-xs text-amber-800">
              Missing from this app&apos;s token: {report.missingRoles.join(", ")}
            </p>
          ) : null}
          {(report.hints.length > 0 ? report.hints : report.hint ? [report.hint] : []).map(
            (item) => (
              <p key={item} className="text-xs text-slate-600">
                {item}
              </p>
            ),
          )}
          <p className="text-xs text-slate-500">
            Emails go to {report.summaryEmailRecipient}
          </p>
          <button
            type="button"
            onClick={() => void loadHealth(true)}
            disabled={refreshing}
            className="rounded-md border border-slate-300 bg-white px-2 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100 disabled:opacity-60"
          >
            {refreshing ? "Refreshing token…" : "Refresh token"}
          </button>
        </div>
      ) : null}
      {open && !report ? (
        <p className="mt-1 text-xs text-slate-500">
          {refreshing ? "Checking Graph…" : "Graph health is not loaded yet."}
        </p>
      ) : null}
    </div>
  );
}
