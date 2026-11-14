export interface TenantUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
}

export type ChatFetchStatus = "ok" | "empty" | "forbidden" | "unavailable";

export type ProbeStatus = "ok" | "denied" | "missing" | "unknown" | "failed";

export interface MeetingListResult {
  meetings: MeetingSummary[];
  transcriptAccessDisabled: boolean;
  fromCache?: boolean;
}

export interface MeetingSummary {
  eventId: string;
  subject: string;
  start: string;
  end: string;
  joinUrl: string;
  userId: string;
  userDisplayName: string;
  /** Azure AD object ID of the meeting organizer (required for Graph onlineMeetings API). */
  organizerId: string;
  /** Recurring series id when this event is an occurrence. */
  seriesMasterId?: string;
  /** Populated when listing meetings that have transcripts. */
  onlineMeetingId?: string;
  latestTranscriptId?: string;
  hasTranscript?: boolean;
  hasRecording?: boolean;
}

export interface TranscriptInfo {
  id: string;
  createdDateTime: string;
}

export type ActionItemType = "Task" | "Bug";
export type ActionItemPriority = "High" | "Medium" | "Low";

export interface ActionItemEvidence {
  timestamp: string;
  speaker: string;
  quote: string;
}

/** Raw action item extracted from Gemini (before meeting metadata is applied). */
export interface ParsedActionItem {
  title: string;
  type: ActionItemType;
  status: string;
  priority: ActionItemPriority;
  pointPerson: string;
  task: string;
  deadline: string;
  devTag: string;
  relatedTo: string;
  description: string;
  evidence: ActionItemEvidence[];
  assigneeUpn?: string;
  assigneeDisplayName?: string;
}

/** Action item with meeting context for display, email, and Microsoft To Do export. */
export interface MeetingActionItem {
  id: string;
  meetingDate: string;
  meetingTitle: string;
  meetingKey: string;
  meetingSeriesId: string;
  ksaDay: string;
  filingKey: string;
  include: boolean;
  alreadyFiled: boolean;
  title: string;
  type: ActionItemType;
  status: string;
  priority: ActionItemPriority;
  pointPerson: string;
  assigneeUpn?: string;
  assigneeDisplayName?: string;
  task: string;
  deadline: string;
  devTag: string;
  relatedTo: string;
  description: string;
  evidence: ActionItemEvidence[];
  pmtUrl?: string;
  pmtNumber?: number;
}

export interface AnalyzeResult {
  analysis: string;
  actionItems: ParsedActionItem[];
  modelUsed?: string;
  videoUsed?: boolean;
  videoNotice?: string;
}

export interface SavedAnalysis {
  meetingKey: string;
  savedAt: string;
  analysis: string;
  actionItems: MeetingActionItem[];
  modelUsed?: string;
  videoUsed?: boolean;
  videoNotice?: string;
}

export interface TodoExportResult {
  created: number;
  listName: string;
  todoUrl: string;
}

export interface EmailSendResult {
  sent: number;
  recipients: string[];
  unmatched: string[];
  message?: string;
}

export interface TeamsAccountStatus {
  signedIn: boolean;
  displayName?: string;
  username?: string;
}

export interface PmtTicketResult {
  externalKey: string;
  id: number;
  number: number;
  url: string;
  assigneeMatched: boolean;
  created: boolean;
  warning?: string;
  relatedLabel?: string;
  /** Fields written onto an existing ticket because they were still empty. */
  filled?: string[];
  assigneeEmail?: string;
  assigneeDisplayName?: string;
}

export interface PmtCreateResult {
  board: { id: number; name: string };
  items: PmtTicketResult[];
  message: string;
}

export interface PmtBoardOption {
  id: number;
  name: string;
}

export interface PmtAgingBucket {
  label: string;
  count: number;
}

export interface PmtStatusCount {
  name: string;
  count: number;
  color: string;
}

export interface PmtHiddenAgingTicket {
  number: number;
  name: string;
  column: string;
  days: number;
  owner: string;
  ownerEmail?: string | null;
  due: string | null;
  url: string;
}

export interface PmtAnalyticsReport {
  board: { id: number; name: string; url: string };
  range: "week" | "month";
  fromLabel: string;
  toLabel: string;
  open: number;
  closed: number;
  overdue: number;
  unassigned: number;
  aging: PmtAgingBucket[];
  statuses: PmtStatusCount[];
  chartTitle: string;
  hiddenAging: PmtHiddenAgingTicket[];
  chartPng: string;
}

export interface PmtAgingPerson {
  name: string;
  email: string;
}

export interface PmtAssigneeAgingTicket {
  number: number;
  name: string;
  column: string;
  days: number;
  url: string;
}

export interface PmtAssigneeAgingProject {
  id: number;
  name: string;
  more: number;
  tickets: PmtAssigneeAgingTicket[];
}

export interface PmtAssigneeAgingReport {
  people: PmtAgingPerson[];
  person: PmtAgingPerson | null;
  projects: PmtAssigneeAgingProject[];
  omittedProjects: number;
}

export interface MeetingChatOption {
  id: string;
  topic: string;
}

export interface TeamsSendFailure {
  assignee: string;
  error: string;
}

export interface TeamsSendResult {
  sent: number;
  recipients: string[];
  unmatched: string[];
  failed: TeamsSendFailure[];
  filedKeys: string[];
  sender: string;
  /** True when Iris was added as a third member of each assignee chat. */
  includedIris: boolean;
}

export interface HealthProbe {
  status: ProbeStatus;
  detail: string;
}

export interface GraphHealthReport {
  token: ProbeStatus;
  targetUser: string;
  summaryEmailRecipient: string;
  /** Application (client) ID from AZURE_CLIENT_ID — must match the Entra app you granted. */
  clientId: string;
  roles: string[];
  missingRoles: string[];
  probes: {
    calendar: HealthProbe;
    transcripts: HealthProbe;
    recordings: HealthProbe;
    chat: HealthProbe;
    mail: HealthProbe;
    todo: HealthProbe;
  };
  hint: string | null;
  hints: string[];
}
