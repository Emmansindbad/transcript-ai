import type { ChatFetchStatus } from "@/types/meeting";

interface TranscriptPanelProps {
  transcript: string | null;
  chat?: string | null;
  chatStatus?: ChatFetchStatus | null;
  chatError?: string | null;
  isLoading: boolean;
  meetingSubject: string | null;
  recordingFallback?: boolean;
  artifactDateLabel?: string | null;
  artifactSource?: string | null;
}

function TranscriptSkeleton(): React.ReactNode {
  return (
    <div className="animate-pulse space-y-3 p-6">
      {Array.from({ length: 12 }).map((_, index) => (
        <div
          key={index}
          className="h-3 rounded bg-slate-200"
          style={{ width: `${70 + (index % 4) * 7}%` }}
        />
      ))}
    </div>
  );
}

function chatStatusLabel(
  chat: string | null,
  chatStatus: ChatFetchStatus | null,
): string {
  if (chat?.trim()) {
    return "Meeting chat loaded";
  }
  if (chatStatus === "forbidden") {
    return "Meeting chat denied (missing Chat.Read.All)";
  }
  if (chatStatus === "unavailable") {
    return "Meeting chat unavailable";
  }
  return "No meeting chat";
}

export function TranscriptPanel({
  transcript,
  chat = null,
  chatStatus = null,
  chatError = null,
  isLoading,
  meetingSubject,
  recordingFallback = false,
  artifactDateLabel = null,
  artifactSource = null,
}: TranscriptPanelProps): React.ReactNode {
  return (
    <section className="flex h-full min-h-0 flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
      <header className="border-b border-slate-200 px-5 py-4">
        <h2 className="text-base font-semibold text-slate-900">
          Transcript
        </h2>
        {meetingSubject ? (
          <p className="mt-1 text-sm text-slate-500">{meetingSubject}</p>
        ) : null}
        {transcript || chat ? (
          <p className="mt-1 text-xs text-slate-400">
            {transcript ? "Spoken transcript loaded" : "No spoken transcript"}
            {" · "}
            {chatStatusLabel(chat, chatStatus)}
            {artifactDateLabel
              ? ` · ${artifactDateLabel} (${artifactSource ?? "meeting"} created, KSA)`
              : " · times in KSA"}
          </p>
        ) : null}
      </header>

      <div className="flex-1 overflow-y-auto">
        {isLoading ? (
          <TranscriptSkeleton />
        ) : transcript || chat ? (
          <div className="space-y-6 p-6">
            {transcript ? (
              <section>
                <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
                  Spoken transcript
                </h3>
                <pre className="whitespace-pre-wrap font-mono text-sm leading-relaxed text-slate-700">
                  {transcript}
                </pre>
              </section>
            ) : null}
            {chat?.trim() ? (
              <section>
                <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
                  Meeting chat
                </h3>
                <pre className="whitespace-pre-wrap font-mono text-sm leading-relaxed text-slate-700">
                  {chat}
                </pre>
              </section>
            ) : (
              <p className="text-sm text-slate-500">
                {chatStatus === "forbidden"
                  ? chatError ??
                    "Graph denied meeting chat. Grant Chat.Read.All and ChatMessage.Read.All application permissions, then restart the app."
                  : chatStatus === "unavailable"
                    ? chatError ?? "Meeting chat was not available for this meeting."
                    : "No meeting chat messages were found for this KSA day."}
              </p>
            )}
          </div>
        ) : recordingFallback ? (
          <div className="flex h-full flex-col items-center justify-center gap-2 p-6 text-center text-sm text-slate-500">
            <p>
              Graph cannot return this transcript yet. A meeting recording is
              available — click Generate Analysis to analyze the video.
            </p>
            {chatStatus === "forbidden" ? (
              <p className="text-amber-800">
                {chatError ??
                  "Meeting chat was also denied. Grant Chat.Read.All to include chat."}
              </p>
            ) : null}
          </div>
        ) : (
          <div className="flex h-full items-center justify-center p-6 text-sm text-slate-500">
            Select a meeting to view its parsed transcript.
          </div>
        )}
      </div>
    </section>
  );
}
