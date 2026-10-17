import { ClearCacheButton } from "@/components/ClearCacheButton";
import { Dashboard } from "@/components/Dashboard";
import { HealthPanel } from "@/components/HealthPanel";
import { ShutdownButton } from "@/components/ShutdownButton";

export default function Home(): React.ReactNode {
  return (
    <div className="flex min-h-screen flex-col bg-slate-100">
      <header className="border-b border-slate-200 bg-white px-4 py-5 shadow-sm lg:px-6">
        <div className="mx-auto flex max-w-7xl flex-col gap-1">
          <p className="text-xs font-semibold uppercase tracking-widest text-blue-600">
            SindbadTech
          </p>
          <div className="flex flex-wrap items-center gap-3">
            <h1 className="text-2xl font-bold text-slate-900">
              M365 Transcript AI Analyzer
            </h1>
            <ClearCacheButton />
            <ShutdownButton />
          </div>
          <p className="text-sm text-slate-500">
            Fetch Teams transcripts, chat, and recordings from Microsoft Graph
            and turn them into tickets with Gemini 3.8 Flash.
          </p>
          <HealthPanel />
        </div>
      </header>

      <Dashboard />
    </div>
  );
}
