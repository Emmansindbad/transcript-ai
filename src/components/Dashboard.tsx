"use client";

import { useCallback, useEffect, useState } from "react";
import type { TeamsAccountStatus, TenantUser, TicketLane } from "@/types/meeting";
import { AgingTicketsPanel } from "@/components/AgingTicketsPanel";
import { CodeQcPanel } from "@/components/CodeQcPanel";
import { ErrorBanner } from "@/components/ErrorBanner";
import { MeetingsWorkspace } from "@/components/MeetingsWorkspace";
import { PmtAnalyticsPanel } from "@/components/PmtAnalyticsPanel";

interface ApiErrorResponse {
  error?: string;
  code?: string;
}

type WorkspaceId = TicketLane | "aging" | "analytics" | "code-qc";

export function Dashboard(): React.ReactNode {
  const [users, setUsers] = useState<TenantUser[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(true);
  const [selectedUserId, setSelectedUserId] = useState("all");
  const [emailRecipient, setEmailRecipient] = useState("mom@sindbad.tech");
  const [teamsAccount, setTeamsAccount] = useState<TeamsAccountStatus | null>(null);
  const [error, setError] = useState<ApiErrorResponse | null>(null);
  const [workspace, setWorkspace] = useState<WorkspaceId>("development");

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

        if (!active) {
          return;
        }
        if (!usersResponse.ok) {
          setError({
            error: usersData.error ?? "Could not load people from Microsoft 365.",
            code: usersData.code ?? "USERS_FAILED",
          });
          return;
        }
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
  }, []);

  const workspaces: Array<{ id: WorkspaceId; label: string }> = [
    { id: "development", label: "Meetings" },
    { id: "sheets", label: "Meeting for Sheets" },
    { id: "aging", label: "Aging" },
    { id: "analytics", label: "Analytics" },
    { id: "code-qc", label: "CODE QC" },
  ];

  const meetingWorkspace = workspace === "development" || workspace === "sheets";

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <ErrorBanner
        message={error?.error ?? null}
        code={error?.code}
        onDismiss={() => setError(null)}
      />

      <nav className="flex shrink-0 gap-6 overflow-x-auto border-b border-zinc-200 bg-white px-5">
        {workspaces.map((item) => (
          <button
            key={item.id}
            type="button"
            onClick={() => setWorkspace(item.id)}
            className={`shrink-0 border-b-2 py-3 text-sm whitespace-nowrap transition-colors ${
              workspace === item.id
                ? "border-zinc-900 font-medium text-zinc-900"
                : "border-transparent text-zinc-500 hover:text-zinc-800"
            }`}
          >
            {item.label}
          </button>
        ))}
      </nav>

      <div className="flex min-h-0 flex-1 flex-col lg:flex-row">
        <div className={workspace === "development" ? "contents" : "hidden"}>
          <MeetingsWorkspace
            lane="development"
            users={users}
            isLoadingUsers={isLoadingUsers}
            selectedUserId={selectedUserId}
            onUserChange={handleUserChange}
            emailRecipient={emailRecipient}
            teamsAccount={teamsAccount}
            onTeamsAccountChange={setTeamsAccount}
          />
        </div>
        <div className={workspace === "sheets" ? "contents" : "hidden"}>
          <MeetingsWorkspace
            lane="sheets"
            users={users}
            isLoadingUsers={isLoadingUsers}
            selectedUserId={selectedUserId}
            onUserChange={handleUserChange}
            emailRecipient={emailRecipient}
            teamsAccount={teamsAccount}
            onTeamsAccountChange={setTeamsAccount}
          />
        </div>
        {meetingWorkspace ? null : (
          <main className="min-h-0 flex-1 overflow-y-auto bg-white">
            <div
              className={
                workspace === "aging" ? "mx-auto w-full max-w-3xl px-6 py-8" : "hidden"
              }
            >
              <AgingTicketsPanel
                teamsAccount={teamsAccount}
                onTeamsSignIn={() => {
                  window.location.assign("/api/auth/teams/start");
                }}
              />
            </div>
            <div
              className={
                workspace === "analytics" ? "mx-auto w-full max-w-3xl px-6 py-8" : "hidden"
              }
            >
              <PmtAnalyticsPanel
                teamsAccount={teamsAccount}
                onTeamsSignIn={() => {
                  window.location.assign("/api/auth/teams/start");
                }}
              />
            </div>
            <div
              className={
                workspace === "code-qc" ? "mx-auto w-full max-w-5xl px-6 py-8" : "hidden"
              }
            >
              <CodeQcPanel
                teamsAccount={teamsAccount}
                onTeamsSignIn={() => {
                  window.location.assign("/api/auth/teams/start");
                }}
              />
            </div>
          </main>
        )}
      </div>
    </div>
  );
}
