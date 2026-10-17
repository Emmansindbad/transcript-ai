import { ConfidentialClientApplication } from "@azure/msal-node";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

function loadEnv() {
  const envPath = resolve(process.cwd(), ".env.local");
  const content = readFileSync(envPath, "utf8");
  for (const line of content.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    process.env[trimmed.slice(0, eq)] = trimmed.slice(eq + 1);
  }
}

loadEnv();

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;
const userId = process.env.TARGET_USER_ID;

const cca = new ConfidentialClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret,
  },
});

const tokenResult = await cca.acquireTokenByClientCredential({
  scopes: ["https://graph.microsoft.com/.default"],
});

if (!tokenResult?.accessToken) {
  console.error("TOKEN_FAIL: Could not acquire access token");
  process.exit(1);
}

console.log("TOKEN_OK");

const tokenPayload = JSON.parse(
  Buffer.from(tokenResult.accessToken.split(".")[1], "base64url").toString(),
);
const roles = tokenPayload.roles ?? [];
console.log("TOKEN_ROLES:", roles.join(", ") || "(none)");

const required = [
  "Calendars.Read",
  "OnlineMeetings.Read.All",
  "OnlineMeetingTranscript.Read.All",
  "OnlineMeetingRecording.Read.All",
  "User.Read.All",
  "Mail.Send",
  "Chat.Read.All",
  "ChatMessage.Read.All",
  "Tasks.ReadWrite.All",
];
for (const permission of required) {
  console.log(
    permission + ":",
    roles.includes(permission) ? "GRANTED" : "MISSING",
  );
}

// Probe mail send capability by checking the mailbox (sendMail requires a POST so we verify access via mailFolders)
const mailResponse = await fetch(
  `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/mailFolders/SentItems`,
  { headers: { Authorization: `Bearer ${tokenResult.accessToken}` } },
);
console.log("MAIL_STATUS:", mailResponse.status);
if (!mailResponse.ok) {
  const mailBody = await mailResponse.text();
  console.log("MAIL_BODY:", mailBody.slice(0, 500));
}

const end = new Date();
const start = new Date();
start.setDate(start.getDate() - 30);

const url =
  `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/calendarView` +
  `?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$top=5`;

const response = await fetch(url, {
  headers: {
    Authorization: `Bearer ${tokenResult.accessToken}`,
    Prefer: 'outlook.timezone="UTC"',
  },
});

const body = await response.text();
console.log("CALENDAR_STATUS:", response.status);
console.log("CALENDAR_BODY:", body.slice(0, 1500));

const rolesResponse = await fetch(
  `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${clientId}'&$select=appRoles,oauth2PermissionGrants,appRoleAssignments`,
  {
    headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
  },
);

if (rolesResponse.ok) {
  const rolesBody = await rolesResponse.json();
  console.log("SP_INFO:", JSON.stringify(rolesBody, null, 2).slice(0, 2000));
} else {
  console.log("SP_INFO_STATUS:", rolesResponse.status);
}
