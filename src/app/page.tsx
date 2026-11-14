import { ClearCacheButton } from "@/components/ClearCacheButton";
import { Dashboard } from "@/components/Dashboard";
import { HealthPanel } from "@/components/HealthPanel";
import { ShutdownButton } from "@/components/ShutdownButton";

export default function Home(): React.ReactNode {
  return (
    <div className="flex min-h-screen flex-col bg-zinc-100 text-zinc-900 lg:h-screen lg:overflow-hidden">
      <header className="flex flex-wrap items-center justify-between gap-x-4 gap-y-1 border-b border-zinc-200 bg-white px-4 py-3 sm:px-5">
        <div className="min-w-0">
          <p className="text-xs text-zinc-400">Sindbad</p>
          <h1 className="text-sm font-semibold tracking-tight">Transcript analyzer</h1>
        </div>
        <div className="flex items-center gap-1">
          <HealthPanel />
          <ClearCacheButton />
          <ShutdownButton />
        </div>
      </header>

      <Dashboard />
    </div>
  );
}
