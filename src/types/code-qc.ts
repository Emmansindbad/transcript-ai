export interface CodeQcIssue {
  key: string;
  message: string;
  type: "BUG" | "VULNERABILITY" | "CODE_SMELL" | string;
  severity: "BLOCKER" | "CRITICAL" | "MAJOR" | "MINOR" | "INFO" | string;
  status: string;
  project: string;
  projectName: string;
  component: string;
  line?: number;
  rule?: string;
  url: string;
  alreadyFiled?: boolean;
}

export interface CodeQcProject {
  key: string;
  name: string;
  issueCount: number;
}

export interface CodeQcSnapshot {
  issues: CodeQcIssue[];
  projects: CodeQcProject[];
  boards: Array<{ id: number; name: string }>;
  users: Array<{
    id: string;
    displayName: string;
    userPrincipalName: string;
    mail?: string;
  }>;
  defaultBoardName?: string;
}

export interface CodeQcExistingTicket {
  externalKey: string;
  number: number;
  url: string;
  title: string;
}

export interface CodeQcProposedTicket {
  id: string;
  title: string;
  type: "Task" | "Bug";
  priority: "High" | "Medium" | "Low";
  issueKeys: string[];
  description: string;
  relatedTo: string;
  existing?: CodeQcExistingTicket;
}

export interface CodeQcAnalyzeResult {
  tickets: CodeQcProposedTicket[];
  modelUsed: string;
  message: string;
}

export interface CodeQcCreateTicketInput {
  title: string;
  type: "Task" | "Bug";
  priority: "High" | "Medium" | "Low";
  issueKeys: string[];
}

export interface CodeQcCreateRequest {
  boardName: string;
  issueKeys?: string[];
  tickets?: CodeQcCreateTicketInput[];
  assignAllUserId?: string;
}

export interface CodeQcSendTicket {
  externalKey: string;
  number: number;
  url: string;
  title?: string;
}

export interface CodeQcSendRequest {
  boardName: string;
  groupUserIds: string[];
  note: string;
  tickets: CodeQcSendTicket[];
}

export interface CodeQcTeamsResult {
  sent: boolean;
  topic: string;
  previewCount: number;
  hiddenCount: number;
  warning?: string;
}

export interface CodeQcCreateResult {
  board: { id: number; name: string };
  items: Array<{
    externalKey: string;
    id: number;
    number: number;
    url: string;
    title: string;
    created: boolean;
    warning?: string;
  }>;
  message: string;
  teams?: CodeQcTeamsResult;
}
