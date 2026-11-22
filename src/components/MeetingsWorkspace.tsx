"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import type {
  ChatFetchStatus,
  EmailSendResult,
  MeetingActionItem,
  MeetingSummary,
  PmtBoardOption,
  PmtCreateResult,
  SavedAnalysis,
  TeamsAccountStatus,
  TeamsSendResult,
  TenantUser,
  TicketLane,
} from "@/types/meeting";
import { AnalysisPanel } from "@/components/AnalysisPanel";
import { ErrorBanner } from "@/components/ErrorBanner";
import { Sidebar } from "@/components/Sidebar";
import { TranscriptPanel } from "@/components/TranscriptPanel";
import { formatKsaDate } from "@/lib/ksa-time";
import { assigneeFieldsForName } from "@/lib/assignees";
import { buildWorkspaceMeetingKey } from "@/lib/ticket-lane";
import { readSseEvents } from "@/lib/sse-client";

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

interface MeetingsWorkspaceProps {
  lane: TicketLane;
  users: TenantUser[];
  isLoadingUsers: boolean;
  selectedUserId: string;
  onUserChange: (userId: string) => void;
  emailRecipient: string;
  teamsAccount: TeamsAccountStatus | null;
  onTeamsAccountChange: (account: TeamsAccountStatus) => void;
}

function formatMeetingDateLabel(isoDate: string): string {
  return formatKsaDate(isoDate);
}

export function MeetingsWorkspace({
  lane,
  users,
  isLoadingUsers,
  selectedUserId,
  onUserChange,
  emailRecipient,
  teamsAccount,
  onTeamsAccountChange,
}: MeetingsWorkspaceProps): React.ReactNode {
  const sheets = lane === "sheets";
  const [meetings, setMeetings] = useState<MeetingSummary[]>([]);
  const [selectedMeeting, setSelectedMeeting] = useState<MeetingSummary | null>(null);
  const [transcript, setTranscript] = useState<string | null>(null);
  const [meetingChat, setMeetingChat] = useState<string | null>(null);
  const [chatStatus, setChatStatus] = useState<ChatFetchStatus | null>(null);
  const [chatError, setChatError] = useState<string | null>(null);
  const [artifactDateKsa, setArtifactDateKsa] = useState<string | null>(null);
  const [artifactDateLabel, setArtifactDateLabel] = useState<string | null>(null);
  const [artifactSource, setArtifactSource] = useState<string | null>(null);
  const [analysis, setAnalysis] = useState<string | null>(null);
  const [analysisModel, setAnalysisModel] = useState<string | null>(null);
  const [actionItems, setActionItems] = useState<MeetingActionItem[]>([]);
  const [isLoadingMeetings, setIsLoadingMeetings] = useState(false);
  const [loadedForUserId, setLoadedForUserId] = useState<string | null>(null);
  const [isLoadingTranscript, setIsLoadingTranscript] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [streamPreview, setStreamPreview] = useState("");
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [isSendingEmail, setIsSendingEmail] = useState(false);
  const [emailSendSuccess, setEmailSendSuccess] = useState(false);
  const [emailRecipients, setEmailRecipients] = useState<string[]>([]);
  const [isSendingTeams, setIsSendingTeams] = useState(false);
  const [teamsSummary, setTeamsSummary] = useState<string | null>(null);
  const [isSendingPmt, setIsSendingPmt] = useState(false);
  const [pmtSummary, setPmtSummary] = useState<string | null>(null);
  const [pmtItems, setPmtItems] = useState<PmtCreateResult["items"]>([]);
  const [destinations, setDestinations] = useState<PmtBoardOption[]>([]);
  const [destinationId, setDestinationId] = useState("");
  const [destinationsLoading, setDestinationsLoading] = useState(false);
  const [destinationsError, setDestinationsError] = useState<string | null>(null);
  const [includeVideo, setIncludeVideo] = useState(false);
  const [videoUsed, setVideoUsed] = useState(false);
  const [videoNotice, setVideoNotice] = useState<string | null>(null);
  const [recordingFallback, setRecordingFallback] = useState(false);
  const [restoredAt, setRestoredAt] = useState<string | null>(null);
  const [error, setError] = useState<ApiErrorResponse | null>(null);
  const persistTimer = useRef<number | null>(null);
  const meetingsRequest = useRef<AbortController | null>(null);
  const previousUserId = useRef(selectedUserId);

  useEffect(() => {
    const controller = new AbortController();
    setDestinationsLoading(true);
    setDestinationsError(null);

    void (async () => {
      try {
        const response = await fetch(`/api/pmt/destinations?lane=${lane}`, {
          signal: controller.signal,
        });
        const data = (await response.json()) as {
          destinations?: PmtBoardOption[];
          error?: string;
        };
        if (!response.ok) {
          setDestinations([]);
          setDestinationId("");
          setDestinationsError(data.error ?? "Could not load the list.");
          return;
        }
        const list = Array.isArray(data.destinations) ? data.destinations : [];
        setDestinations(list);
        setDestinationId((current) => {
          if (current && list.some((item) => String(item.id) === current)) {
            return current;
          }
          if (list.length === 1) {
            return String(list[0]?.id ?? "");
          }
          return "";
        });
      } catch (error) {
        if (error instanceof DOMException && error.name === "AbortError") {
          return;
        }
        setDestinations([]);
        setDestinationsError(sheets ? "Could not load sheets." : "Could not load boards.");
      } finally {
        if (!controller.signal.aborted) {
          setDestinationsLoading(false);
        }
      }
    })();

    return () => {
      controller.abort();
    };
  }, [lane, sheets, teamsAccount?.signedIn, teamsAccount?.username]);

  useEffect(() => {
    if (destinationId || !selectedMeeting) {
      return;
    }
    const subject = selectedMeeting.subject.trim().toLowerCase();
    const match = destinations.find((item) => item.name.trim().toLowerCase() === subject);
    if (match) {
      setDestinationId(String(match.id));
    }
  }, [destinationId, destinations, selectedMeeting]);

  const resetSendState = useCallback((): void => {
    setEmailSendSuccess(false);
    setEmailRecipients([]);
    setTeamsSummary(null);
    setPmtSummary(null);
    setPmtItems([]);
    setVideoUsed(false);
    setVideoNotice(null);
    setRestoredAt(null);
    setStreamPreview("");
    setStatusMessage(null);
  }, []);

  const clearMeeting = useCallback((): void => {
    setSelectedMeeting(null);
    setTranscript(null);
    setMeetingChat(null);
    setChatStatus(null);
    setChatError(null);
    setArtifactDateKsa(null);
    setArtifactDateLabel(null);
    setArtifactSource(null);
    setRecordingFallback(false);
    setAnalysis(null);
    setAnalysisModel(null);
    setActionItems([]);
    resetSendState();
  }, [resetSendState]);

  useEffect(() => {
    if (previousUserId.current === selectedUserId) {
      return;
    }
    previousUserId.current = selectedUserId;
    setMeetings([]);
    setLoadedForUserId(null);
    clearMeeting();
    setError(null);
  }, [clearMeeting, selectedUserId]);

  const persistAnalysis = useCallback(
    (nextItems: MeetingActionItem[], nextAnalysis: string | null): void => {
      if (!selectedMeeting || nextAnalysis === null) {
        return;
      }
      if (persistTimer.current) {
        window.clearTimeout(persistTimer.current);
      }
      persistTimer.current = window.setTimeout(() => {
        void fetch("/api/analyses", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            meetingKey: buildWorkspaceMeetingKey(selectedMeeting, lane),
            analysis: nextAnalysis,
            actionItems: nextItems,
            modelUsed: analysisModel,
            videoUsed,
            videoNotice,
          }),
        });
      }, 800);
    },
    [analysisModel, lane, selectedMeeting, videoNotice, videoUsed],
  );

  const handleItemsChange = useCallback(
    (items: MeetingActionItem[]): void => {
      setActionItems(items);
      persistAnalysis(items, analysis);
    },
    [analysis, persistAnalysis],
  );

  useEffect(() => {
    if (users.length === 0 || actionItems.length === 0) {
      return;
    }
    let changed = false;
    const next = actionItems.map((item) => {
      const fields = assigneeFieldsForName(item.pointPerson, users);
      if (
        fields.assigneeUpn === (item.assigneeUpn ?? "") &&
        fields.assigneeDisplayName === (item.assigneeDisplayName ?? "")
      ) {
        return item;
      }
      changed = true;
      return { ...item, ...fields };
    });
    if (changed) {
      handleItemsChange(next);
    }
  }, [actionItems, handleItemsChange, users]);

  const fetchMeetings = useCallback(async (userId: string, date: string): Promise<void> => {
    meetingsRequest.current?.abort();
    const controller = new AbortController();
    meetingsRequest.current = controller;

    setIsLoadingMeetings(true);
    setError(null);
    clearMeeting();

    try {
      const params = new URLSearchParams();
      params.set("userId", userId || "all");
      if (date) {
        params.set("date", date);
      }
      const response = await fetch(`/api/meetings?${params.toString()}`, {
        signal: controller.signal,
      });
      const data = (await response.json()) as {
        meetings?: MeetingSummary[];
        transcriptAccessDisabled?: boolean;
      } & ApiErrorResponse;

      if (!response.ok) {
        setError({
          error: data.error ?? "Failed to load meetings.",
          code: data.code,
        });
        setMeetings([]);
        return;
      }

      setMeetings(data.meetings ?? []);
      setLoadedForUserId(userId);
      if (data.transcriptAccessDisabled) {
        setError({
          error:
            data.error ??
            "Graph transcript access is still disabled. Showing calendar meetings; recordings can still be analyzed.",
          code: data.code ?? "TRANSCRIPTS_DISABLED",
        });
      }
    } catch (fetchError) {
      if (fetchError instanceof DOMException && fetchError.name === "AbortError") {
        return;
      }
      setError({
        error: "Unable to reach the meetings API.",
        code: "NETWORK_ERROR",
      });
      setMeetings([]);
    } finally {
      if (meetingsRequest.current === controller) {
        setIsLoadingMeetings(false);
      }
    }
  }, [clearMeeting]);

  const handleSelectMeeting = useCallback(
    async (meeting: MeetingSummary): Promise<void> => {
      setSelectedMeeting(meeting);
      setTranscript(null);
      setMeetingChat(null);
      setChatStatus(null);
      setChatError(null);
      setArtifactDateKsa(null);
      setArtifactDateLabel(null);
      setArtifactSource(null);
      setRecordingFallback(false);
      setAnalysis(null);
      setAnalysisModel(null);
      setActionItems([]);
      resetSendState();
      setIsLoadingTranscript(true);
      setError(null);

      try {
        const [transcriptResponse, savedResponse] = await Promise.all([
          fetch("/api/transcript", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              joinUrl: meeting.joinUrl,
              subject: meeting.subject,
              organizerId: meeting.organizerId,
              onlineMeetingId: meeting.onlineMeetingId,
              latestTranscriptId: meeting.latestTranscriptId,
              meetingStart: meeting.start,
              meetingEnd: meeting.end,
            }),
          }),
          fetch(
            `/api/analyses?meetingKey=${encodeURIComponent(buildWorkspaceMeetingKey(meeting, lane))}`,
          ),
        ]);

        const data = (await transcriptResponse.json()) as {
          transcript?: string;
          chat?: string;
          chatStatus?: ChatFetchStatus;
          chatError?: string;
          recordingAvailable?: boolean;
          artifactDateKsa?: string;
          artifactDateLabel?: string;
          artifactSource?: string;
        } & ApiErrorResponse;

        const savedData = (await savedResponse.json()) as {
          analysis?: SavedAnalysis | null;
        };

        if (savedResponse.ok && savedData.analysis) {
          setAnalysis(savedData.analysis.analysis);
          setActionItems(
            savedData.analysis.actionItems.map((item) => ({ ...item, lane })),
          );
          setAnalysisModel(savedData.analysis.modelUsed ?? null);
          setVideoUsed(Boolean(savedData.analysis.videoUsed));
          setVideoNotice(savedData.analysis.videoNotice ?? null);
          setRestoredAt(savedData.analysis.savedAt);
        }

        if (!transcriptResponse.ok) {
          if (data.chat?.trim()) {
            setMeetingChat(data.chat);
          }
          setChatStatus(data.chatStatus ?? null);
          setChatError(data.chatError ?? null);
          if (data.code === "TRANSCRIPTS_DISABLED" && data.recordingAvailable) {
            setRecordingFallback(true);
            setIncludeVideo(true);
            setError({
              error:
                "Graph cannot return this transcript yet. The meeting recording is available — Generate Analysis will use the video.",
              code: data.code,
            });
            return;
          }

          setError({
            error: data.error ?? "Failed to load transcript.",
            code: data.code,
          });
          return;
        }

        setTranscript(data.transcript ?? null);
        setMeetingChat(data.chat ?? null);
        setChatStatus(data.chatStatus ?? null);
        setChatError(data.chatError ?? null);
        setArtifactDateKsa(data.artifactDateKsa ?? null);
        setArtifactDateLabel(data.artifactDateLabel ?? null);
        setArtifactSource(data.artifactSource ?? null);
        if (data.chatStatus === "forbidden") {
          setError({
            error:
              data.chatError ??
              "Meeting chat was denied. Analysis can still use the transcript and recording.",
            code: "CHAT_FORBIDDEN",
          });
        }
      } catch {
        setError({
          error: "Unable to reach the transcript API.",
          code: "NETWORK_ERROR",
        });
      } finally {
        setIsLoadingTranscript(false);
      }
    },
    [lane, resetSendState],
  );

  const handleGenerateAnalysis = useCallback(async (): Promise<void> => {
    if (
      (!transcript?.trim() && !meetingChat?.trim() && !recordingFallback) ||
      !selectedMeeting
    ) {
      return;
    }

    setIsAnalyzing(true);
    setError(null);
    resetSendState();
    setStreamPreview("");
    setStatusMessage("Starting Gemini analysis…");

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          transcript: transcript ?? "",
          chat: meetingChat ?? "",
          meetingDate: formatMeetingDateLabel(selectedMeeting.start),
          artifactDateKsa,
          artifactDateLabel,
          artifactSource,
          includeVideo: includeVideo || recordingFallback,
          joinUrl: selectedMeeting.joinUrl,
          organizerId: selectedMeeting.organizerId,
          onlineMeetingId: selectedMeeting.onlineMeetingId,
          meeting: selectedMeeting,
          lane,
        }),
      });

      if (!response.ok && response.headers.get("content-type")?.includes("application/json")) {
        const data = (await response.json()) as ApiErrorResponse;
        setError({
          error: data.error ?? "Failed to generate analysis.",
          code: data.code,
        });
        return;
      }

      await readSseEvents(response, (event) => {
        if (event.type === "status" && typeof event.message === "string") {
          setStatusMessage(event.message);
        }
        if (event.type === "chunk" && typeof event.text === "string") {
          setStreamPreview((current) => `${current}${event.text}`);
        }
        if (event.type === "error") {
          setError({
            error:
              typeof event.error === "string"
                ? event.error
                : "Failed to generate analysis.",
            code: typeof event.code === "string" ? event.code : undefined,
          });
        }
        if (event.type === "done") {
          setAnalysis(typeof event.analysis === "string" ? event.analysis : "");
          setActionItems(
            Array.isArray(event.actionItems)
              ? (event.actionItems as MeetingActionItem[])
              : [],
          );
          setAnalysisModel(
            typeof event.modelUsed === "string" ? event.modelUsed : null,
          );
          setVideoUsed(Boolean(event.videoUsed));
          setVideoNotice(
            typeof event.videoNotice === "string" ? event.videoNotice : null,
          );
          setRestoredAt(null);
        }
      });
    } catch {
      setError({
        error: "Unable to reach the analysis API.",
        code: "NETWORK_ERROR",
      });
    } finally {
      setIsAnalyzing(false);
      setStatusMessage(null);
    }
  }, [
    transcript,
    meetingChat,
    selectedMeeting,
    artifactDateKsa,
    artifactDateLabel,
    artifactSource,
    includeVideo,
    recordingFallback,
    lane,
    resetSendState,
  ]);

  const handleSendEmail = useCallback(async (): Promise<void> => {
    const selected = actionItems.filter((item) => item.include);
    if (selected.length === 0) {
      return;
    }

    setIsSendingEmail(true);
    setError(null);

    try {
      const response = await fetch("/api/email", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ actionItems: selected, analysis }),
      });
      const data = (await response.json()) as EmailSendResult & ApiErrorResponse;

      if (!response.ok) {
        setError({
          error: data.error ?? "Failed to send action items email.",
          code: data.code,
        });
        return;
      }

      setEmailSendSuccess(true);
      setEmailRecipients(data.recipients ?? []);
      setActionItems((current) =>
        current.map((item) =>
          item.include ? { ...item, alreadyFiled: true, include: false } : item,
        ),
      );
    } catch {
      setError({
        error: "Unable to reach the email API.",
        code: "NETWORK_ERROR",
      });
    } finally {
      setIsSendingEmail(false);
    }
  }, [actionItems, analysis]);

  const handleSendTeams = useCallback(async (includeIris: boolean): Promise<void> => {
    const selected = actionItems.filter((item) => item.include);
    if (selected.length === 0) {
      return;
    }

    setIsSendingTeams(true);
    setError(null);
    setTeamsSummary(null);

    try {
      const response = await fetch("/api/teams", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ actionItems: selected, includeIris }),
      });
      const data = (await response.json()) as TeamsSendResult &
        ApiErrorResponse & { message?: string };

      if (!response.ok) {
        if (data.code === "TEAMS_SIGNIN_REQUIRED") {
          onTeamsAccountChange({ signedIn: false });
        }
        setError({
          error: data.error ?? "Failed to send Teams chats.",
          code: data.code,
        });
        return;
      }

      const filed = new Set(data.filedKeys ?? []);
      setTeamsSummary(data.message ?? `Messaged ${data.recipients.join(", ")}.`);
      setActionItems((current) =>
        current.map((item) =>
          filed.has(item.filingKey)
            ? { ...item, alreadyFiled: true, include: false }
            : item,
        ),
      );
      if ((data.failed?.length ?? 0) > 0 || (data.unmatched?.length ?? 0) > 0) {
        const unmatched = data.unmatched?.length
          ? `No Teams user matched: ${data.unmatched.join(", ")}.`
          : "";
        const failures = (data.failed ?? [])
          .map((failure) => `${failure.assignee}: ${failure.error}`)
          .join(" ");
        setError({
          error: `${unmatched} ${failures}`.trim(),
          code: "TEAMS_PARTIAL",
        });
      }
    } catch {
      setError({
        error: "Unable to reach the Teams API.",
        code: "NETWORK_ERROR",
      });
    } finally {
      setIsSendingTeams(false);
    }
  }, [actionItems, onTeamsAccountChange]);

  const handleSendPmt = useCallback(async (): Promise<void> => {
    const selected = actionItems.filter((item) => item.include);
    if (selected.length === 0) {
      return;
    }
    const destination = destinations.find((item) => String(item.id) === destinationId);
    if (!destination) {
      setError({
        error: sheets ? "Choose a sheet first." : "Choose a board first.",
        code: "PMT_DESTINATION",
      });
      return;
    }

    setIsSendingPmt(true);
    setError(null);
    setPmtSummary(null);
    setPmtItems([]);

    try {
      const response = await fetch("/api/pmt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          actionItems: selected.map((item) => ({ ...item, lane })),
          destinationId: destination.id,
          destinationName: destination.name,
        }),
      });
      const data = (await response.json()) as PmtCreateResult & ApiErrorResponse;

      if (!response.ok) {
        setError({
          error: data.error ?? "Failed to create PMT tickets.",
          code: data.code,
        });
        return;
      }

      const createdByKey = new Map(
        (data.items ?? []).map((item) => [item.externalKey, item]),
      );
      const nextItems = actionItems.map((item) => {
        const created = createdByKey.get(item.filingKey);
        if (!created) {
          return item;
        }
        return {
          ...item,
          lane,
          pmtUrl: created.url,
          pmtNumber: created.number,
          relatedTo: created.relatedLabel?.trim() || item.relatedTo || "None",
          ...(created.assigneeEmail
            ? {
                assigneeUpn: created.assigneeEmail,
                assigneeDisplayName: created.assigneeDisplayName ?? item.assigneeDisplayName,
              }
            : {}),
        };
      });
      setActionItems(nextItems);
      persistAnalysis(nextItems, analysis);
      setPmtItems(data.items ?? []);
      setPmtSummary(data.message ?? `Tickets sent to ${destination.name}.`);
      const warnings = (data.items ?? [])
        .map((item) => item.warning)
        .filter((warning): warning is string => Boolean(warning));
      if (warnings.length > 0) {
        setError({
          error: warnings.join(" "),
          code: "PMT_PARTIAL",
        });
      }
    } catch {
      setError({
        error: "Unable to reach the PMT API.",
        code: "NETWORK_ERROR",
      });
    } finally {
      setIsSendingPmt(false);
    }
  }, [actionItems, analysis, destinationId, destinations, lane, persistAnalysis, sheets]);

  const handleTeamsSignOut = useCallback(async (): Promise<void> => {
    setTeamsSummary(null);
    try {
      await fetch("/api/auth/teams", { method: "DELETE" });
    } catch {
      // The next status check shows whether the session is still present.
    }
    onTeamsAccountChange({ signedIn: false });
  }, [onTeamsAccountChange]);

  return (
    <>
      <Sidebar
        key={`${lane}-${selectedUserId}`}
        users={users}
        selectedUserId={selectedUserId}
        onUserChange={onUserChange}
        meetings={meetings}
        selectedEventId={selectedMeeting?.eventId ?? null}
        onSelect={(meeting) => {
          void handleSelectMeeting(meeting);
        }}
        onClearSelection={clearMeeting}
        onFetchData={(date) => fetchMeetings(selectedUserId, date)}
        isLoadingUsers={isLoadingUsers}
        isLoading={isLoadingMeetings}
        hasLoadedMeetings={loadedForUserId === selectedUserId}
        showOrganizer={selectedUserId === "all"}
        heading={sheets ? "Meeting for Sheets" : "Meetings"}
        note={
          sheets
            ? "Non-development meetings. Choose the sheet where these rows should go."
            : "Choose the board where these tickets should go."
        }
      />
      <main className="flex min-h-0 min-w-0 flex-1 flex-col overflow-hidden">
        <ErrorBanner
          message={error?.error ?? null}
          code={error?.code}
          onDismiss={() => setError(null)}
        />
        <div className="grid min-h-0 flex-1 overflow-y-auto lg:grid-cols-2 lg:grid-rows-[minmax(0,1fr)] lg:overflow-hidden">
          <TranscriptPanel
            transcript={transcript}
            chat={meetingChat}
            chatStatus={chatStatus}
            chatError={chatError}
            isLoading={isLoadingTranscript}
            meetingSubject={selectedMeeting?.subject ?? null}
            recordingFallback={recordingFallback}
            artifactDateLabel={artifactDateLabel}
            artifactSource={artifactSource}
          />
          <AnalysisPanel
            analysis={analysis}
            actionItems={actionItems}
            onItemsChange={handleItemsChange}
            modelUsed={analysisModel}
            isLoading={isAnalyzing}
            streamPreview={streamPreview}
            statusMessage={statusMessage}
            hasTranscript={Boolean(transcript?.trim() || meetingChat?.trim())}
            videoRequired={recordingFallback}
            onGenerate={() => {
              void handleGenerateAnalysis();
            }}
            onSendEmail={() => {
              void handleSendEmail();
            }}
            onSendTeams={(includeIris) => {
              void handleSendTeams(includeIris);
            }}
            onSendPmt={() => {
              void handleSendPmt();
            }}
            onTeamsSignIn={() => {
              window.location.assign("/api/auth/teams/start");
            }}
            onTeamsSignOut={() => {
              void handleTeamsSignOut();
            }}
            isSendingEmail={isSendingEmail}
            isSendingTeams={isSendingTeams}
            emailSendSuccess={emailSendSuccess}
            emailRecipients={emailRecipients}
            emailRecipient={emailRecipient}
            teamsAccount={teamsAccount}
            teamsSummary={teamsSummary}
            isSendingPmt={isSendingPmt}
            pmtSummary={pmtSummary}
            pmtItems={pmtItems}
            includeVideo={includeVideo}
            onToggleIncludeVideo={setIncludeVideo}
            videoUsed={videoUsed}
            videoNotice={videoNotice}
            restoredAt={restoredAt}
            filingNote={
              sheets
                ? "Rows use the sheet's own columns: title, status, owner, due, and priority."
                : "Tickets use board fields, including type, priority, and dev tag."
            }
            pmtActionLabel={sheets ? "Create Ticket to Sheets" : undefined}
            pmtBusyLabel={sheets ? "Creating sheet rows…" : undefined}
            destinations={destinations}
            destinationId={destinationId}
            onDestinationChange={setDestinationId}
            destinationsLoading={destinationsLoading}
            destinationsError={destinationsError}
            destinationKind={sheets ? "sheet" : "board"}
            directoryUsers={users}
          />
        </div>
      </main>
    </>
  );
}
