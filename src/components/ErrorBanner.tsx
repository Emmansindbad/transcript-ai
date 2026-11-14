interface ErrorBannerProps {
  message: string | null;
  code?: string | null;
  onDismiss?: () => void;
}

const ERROR_HINTS: Record<string, string> = {
  MISSING_APP_PERMISSIONS:
    "In Azure Portal: App registrations → your app → API permissions → Add Application permissions for Calendars.Read, OnlineMeetings.Read.All, OnlineMeetingTranscript.Read.All, OnlineMeetingRecording.Read.All, User.Read.All, and Mail.Send, then click Grant admin consent. Or run scripts/grant-graph-permissions.ps1 in PowerShell.",
  NO_ACTION_ITEMS:
    "Generate analysis first to extract action items before sending emails.",
  NO_RECORDING:
    "No recording exists for this meeting. Recording must be enabled during the Teams meeting. The app analyzed the transcript only.",
  VIDEO_PROCESSING_TIMEOUT:
    "Gemini took too long to process the recording. Try again, or uncheck 'Analyze recording video' to use the transcript only.",
  VIDEO_PROCESSING_FAILED:
    "Gemini could not process the recording video. Try again later or uncheck 'Analyze recording video'.",
  VIDEO_UPLOAD_FAILED:
    "Uploading the recording to Gemini failed. Check your network and GEMINI_API_KEY, then retry.",
  FORBIDDEN:
    "Verify Azure app permissions and Application Access Policy for the target user.",
  MEETING_NOT_FOUND:
    "Graph could not find this meeting. Ensure an Application Access Policy covers the meeting organizer (not just the calendar owner).",
  NO_TRANSCRIPT:
    "Transcription must be enabled during the Teams meeting for a transcript to exist.",
  CHAT_FORBIDDEN:
    "Confirm Chat.Read.All and ChatMessage.Read.All are Application permissions on the same app as AZURE_CLIENT_ID (not a different registration), Grant admin consent, then open Graph health and click Refresh token. Until Graph accepts chat, analysis uses the spoken transcript and optional recording only.",
  VIDEO_TOO_LARGE:
    "Recordings over 150 MB are skipped so the server does not run out of memory. Uncheck Analyze recording video, or analyze from transcript and chat.",
  TRANSCRIPTS_DISABLED:
    "Microsoft turned this tenant control on by default at the end of July 2026. The Teams admin toggle can take hours to reach Graph. Force it with: Connect-MicrosoftTeams; Set-CsTeamsMeetingConfiguration -EnableGraphTranscriptAccess $true -EnableAttributedTranscripts $true -Identity Global. Then confirm with Get-CsTeamsMeetingConfiguration. Until Graph accepts transcripts, this app lists calendar meetings and can analyze recordings.",
  RATE_LIMIT:
    "Your Gemini API quota may be exceeded. Wait a minute and retry, enable billing in Google AI Studio, or set GEMINI_MODEL=gemini-2.0-flash in .env.local.",
  MODEL_OVERLOADED:
    "Gemini servers are busy. The app already retried and switched models automatically — wait a minute and click Generate Analysis again. For fewer interruptions, enable billing in Google AI Studio.",
  MODEL_NOT_FOUND:
    "That Gemini model is unavailable on your API key. The app will try the next model in the tier list automatically; restart the dev server if this persists.",
  GRAPH_PAGINATION:
    "Click Fetch Data again. Recurring meetings now load only that date's transcript, not the whole series.",
  INCOMPLETE_JSON:
    "Gemini hit its output limit while writing tickets. The app now retries with a shorter JSON shape automatically — click Generate Analysis again if this banner still appears.",
  GEMINI_NETWORK_ERROR:
    "Could not reach the Gemini API. Check your internet connection and GEMINI_API_KEY.",
  AUTH_FAILED: "Check your GEMINI_API_KEY in .env.local (from https://aistudio.google.com/apikey).",
};

export function ErrorBanner({
  message,
  code,
  onDismiss,
}: ErrorBannerProps): React.ReactNode {
  if (!message) {
    return null;
  }

  const hint = code ? ERROR_HINTS[code] : undefined;
  const isWarning =
    code === "TRANSCRIPTS_DISABLED" ||
    code === "CHAT_FORBIDDEN" ||
    code === "VIDEO_TOO_LARGE";

  return (
    <div
      role="alert"
      className={
        isWarning
          ? "border-b border-zinc-200 bg-white px-5 py-3 text-sm text-zinc-800"
          : "border-b border-red-200 bg-red-50 px-5 py-3 text-sm text-red-900"
      }
    >
      <div className="mx-auto flex max-w-7xl items-start justify-between gap-4">
        <div>
          <p className="font-medium">{message}</p>
          {hint ? <p className="mt-1 text-xs leading-5 text-zinc-500">{hint}</p> : null}
          {code ? <p className="mt-1 font-mono text-[11px] text-zinc-400">{code}</p> : null}
        </div>
        {onDismiss ? (
          <button
            type="button"
            onClick={onDismiss}
            className="shrink-0 rounded-md px-2 py-1 text-xs text-zinc-500 hover:bg-zinc-100 hover:text-zinc-900"
            aria-label="Dismiss error"
          >
            Dismiss
          </button>
        ) : null}
      </div>
    </div>
  );
}
