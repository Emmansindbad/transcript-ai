import "server-only";

export type RequiredEnvVar =
  | "AZURE_TENANT_ID"
  | "AZURE_CLIENT_ID"
  | "AZURE_CLIENT_SECRET"
  | "GEMINI_API_KEY"
  | "TARGET_USER_ID";

export function getEnvVar(name: RequiredEnvVar): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export function getTargetUserId(): string {
  return getEnvVar("TARGET_USER_ID");
}

export function getSummaryEmailRecipient(): string {
  return process.env.SUMMARY_EMAIL_RECIPIENT?.trim() || "mom@sindbad.tech";
}

export function getPmtBaseUrl(): string {
  const configured = process.env.PMT_BASE_URL?.trim() || "https://pmt.sindbad.tech";
  return configured.replace(/\/+$/, "");
}

export function getPmtApiToken(): string {
  const token = process.env.PMT_API_TOKEN?.trim() ?? "";
  if (!token) {
    throw new Error("Missing required environment variable: PMT_API_TOKEN");
  }
  return token;
}
