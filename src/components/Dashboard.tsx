"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import type {
  ChatFetchStatus,
  EmailSendResult,
  MeetingActionItem,
  MeetingSummary,
  SavedAnalysis,
  PmtCreateResult,
  TeamsAccountStatus,
  TeamsSendResult,
  TenantUser,
} from "@/types/meeting";
import { AnalysisPanel } from "@/components/AnalysisPanel";
import { AgingTicketsPanel } from "@/components/AgingTicketsPanel";
import { PmtAnalyticsPanel } from "@/components/PmtAnalyticsPanel";
import { ErrorBanner } from "@/components/ErrorBanner";
import { Sidebar } from "@/components/Sidebar";
import { TranscriptPanel } from "@/components/TranscriptPanel";
import { formatKsaDate } from "@/lib/ksa-time";
import { buildMeetingKey } from "@/lib/meeting-key";
import { readSseEvents } from "@/lib/sse-client";

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

function formatMeetingDateLabel(isoDate: string): string {
  return formatKsaDate(isoDate);
}

export function Dashboard(): React.ReactNode {
  const [users, setUsers] = useState<TenantUser[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(true);
  const [selectedUserId, setSelectedUserId] = useState("all");
  const [meetings, setMeetings] = useState<MeetingSummary[]>([]);
  const [selectedMeeting, setSelectedMeeting] = useState<MeetingSummary | null>(
    null,
  );
  const [transcript, setTranscript] = useState<string | null>(null);
  const [meetingChat, setMeetingChat] = useState<string | null>(null);
  const [chatStatus, setChatStatus] = useState<ChatFetchStatus | null>(null);
  const [chatError, setChatError] = useState<string | null>(null);
  const [artifactDateKsa, setArtifactDateKsa] = useState<string | null>(null);
  const [artifactDateLabel, setArtifactDateLabel] = useState<string | null>(
    null,
  );
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
  const [emailRecipient, setEmailRecipient] = useState("mom@sindbad.tech");
  const [teamsAccount, setTeamsAccount] = useState<TeamsAccountStatus | null>(null);
  const [isSendingTeams, setIsSendingTeams] = useState(false);
  const [teamsSummary, setTeamsSummary] = useState<string | null>(null);
  const [isSendingPmt, setIsSendingPmt] = useState(false);
  const [pmtSummary, setPmtSummary] = useState<string | null>(null);
  const [pmtItems, setPmtItems] = useState<PmtCreateResult["items"]>([]);
  const [includeVideo, setIncludeVideo] = useState(false);
  const [videoUsed, setVideoUsed] = useState(false);
  const [videoNotice, setVideoNotice] = useState<string | null>(null);
  const [recordingFallback, setRecordingFallback] = useState(false);
  const [restoredAt, setRestoredAt] = useState<string | null>(null);
  const [error, setError] = useState<ApiErrorResponse | null>(null);
  const persistTimer = useRef<number | null>(null);

  const resetEmailState = useCallback((): void => {
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
            meetingKey: buildMeetingKey(selectedMeeting),
            analysis: nextAnalysis,
            actionItems: nextItems,
            modelUsed: analysisModel,
            videoUsed,
            videoNotice,
          }),
        });
      }, 800);
    },
    [analysisModel, selectedMeeting, videoNotice, videoUsed],
  );

  const handleItemsChange = useCallback(
    (items: MeetingActionItem[]): void => {
      setActionItems(items);
      persistAnalysis(items, analysis);
    },
    [analysis, persistAnalysis],
  );

  const meetingsRequest = useRef<AbortController | null>(null);

  const fetchMeetings = useCallback(async (userId: string): Promise<void> => {
    meetingsRequest.current?.abort();
    const controller = new AbortController();
    meetingsRequest.current = controller;

    setIsLoadingMeetings(true);
    setError(null);
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
    resetEmailState();

    try {
      const query = userId === "all" ? "" : `?userId=${encodeURIComponent(userId)}`;
      const response = await fetch(`/api/meetings${query}`, {
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
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
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
  }, [resetEmailState]);

  useEffect(() => {
    let active = true;

    async function loadUsers(): Promise<void> {
      try {
        const usersResponse = await fetch("/api/users");
        const usersData = (await usersResponse.json()) as {
          users?: TenantUser[];
          defaultUserId?: string;
          summaryEmailRecipient?: string;
        } & ApiErrorResponse;

        if (active && usersResponse.ok) {
          const nextUsers = usersData.users ?? [];
          setUsers(nextUsers);
          if (usersData.summaryEmailRecipient) {
            setEmailRecipient(usersData.summaryEmailRecipient);
          }
          const defaultUserId = usersData.defaultUserId;
          const preferredUserId =
            defaultUserId &&
            nextUsers.some((user) => user.userPrincipalName === defaultUserId)
              ? defaultUserId
              : nextUsers[0]?.userPrincipalName;
          if (preferredUserId) {
            setSelectedUserId((current) =>
              current === "all" ? preferredUserId : current,
            );
          }
        }
      } catch {
        // Names stay on the last list; meetings still load only from Fetch Data.
      } finally {
        if (active) {
          setIsLoadingUsers(false);
        }
      }
    }

    async function loadTeamsAccount(): Promise<void> {
      try {
        const response = await fetch("/api/auth/teams");
        const data = (await response.json()) as TeamsAccountStatus & ApiErrorResponse;
        if (active && response.ok) {
          setTeamsAccount(data);
        }
      } catch {
        // Sign-in status is optional until the user sends a chat.
      }
    }

    const params = new URLSearchParams(window.location.search);
    const teamsNotice = params.get("teams");
    if (teamsNotice) {
      if (teamsNotice === "error") {
        setError({
          error: params.get("reason") ?? "Teams sign-in failed.",
          code: "TEAMS_SIGNIN",
        });
      }
      const clean = new URL(window.location.href);
      clean.search = "";
      window.history.replaceState({}, "", `${clean.pathname}${clean.hash}`);
    }

    void loadUsers();
    void loadTeamsAccount();

    return () => {
      active = false;
    };
  }, []);

  const handleUserChange = useCallback((userId: string): void => {
    setSelectedUserId(userId);
    setMeetings([]);
    setLoadedForUserId(null);
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
    resetEmailState();
    setError(null);
  }, [resetEmailState]);

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
      resetEmailState();
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
            `/api/analyses?meetingKey=${encodeURIComponent(buildMeetingKey(meeting))}`,
          ),
        ]);

        const data = (await transcriptResponse.json()) as {
          transcript?: string;
          chat?: string;
          chatStatus?: ChatFetchStatus;
          chatError?: string;
          chatMessageCount?: number;
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
          setActionItems(savedData.analysis.actionItems);
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
    [resetEmailState],
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
    resetEmailState();
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
    resetEmailState,
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
          setTeamsAccount({ signedIn: false });
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
  }, [actionItems]);

  const handleSendPmt = useCallback(async (): Promise<void> => {
    const selected = actionItems.filter((item) => item.include);
    if (selected.length === 0) {
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
        body: JSON.stringify({ actionItems: selected }),
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
      setPmtSummary(data.message ?? "Tickets sent to PMT.");
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
  }, [actionItems, analysis, persistAnalysis]);

  const handleTeamsSignOut = useCallback(async (): Promise<void> => {
    setTeamsSummary(null);
    try {
      await fetch("/api/auth/teams", { method: "DELETE" });
    } catch {
      // The next status check shows whether the session is still present.
    }
    setTeamsAccount({ signedIn: false });
  }, []);

  const handleClearSelection = useCallback((): void => {
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
    resetEmailState();
  }, [resetEmailState]);

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <ErrorBanner
        message={error?.error ?? null}
        code={error?.code}
        onDismiss={() => setError(null)}
      />

      <div className="flex min-h-0 flex-1 flex-col lg:flex-row">
        <Sidebar
          key={selectedUserId}
          users={users}
          selectedUserId={selectedUserId}
          onUserChange={handleUserChange}
          meetings={meetings}
          selectedJoinUrl={selectedMeeting?.joinUrl ?? null}
          onSelect={(meeting) => {
            void handleSelectMeeting(meeting);
          }}
          onClearSelection={handleClearSelection}
          onFetchData={() => fetchMeetings(selectedUserId)}
          isLoadingUsers={isLoadingUsers}
          isLoading={isLoadingMeetings}
          hasLoadedMeetings={loadedForUserId === selectedUserId}
          showOrganizer={selectedUserId === "all"}
        />

        <main className="flex min-h-0 flex-1 flex-col gap-4 p-4 lg:grid lg:grid-cols-2 lg:gap-4 lg:p-6">
          <AgingTicketsPanel
            teamsAccount={teamsAccount}
            onTeamsSignIn={() => {
              window.location.assign("/api/auth/teams/start");
            }}
          />
          <PmtAnalyticsPanel
            teamsAccount={teamsAccount}
            onTeamsSignIn={() => {
              window.location.assign("/api/auth/teams/start");
            }}
          />
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
          />
        </main>
      </div>
    </div>
  );
}
