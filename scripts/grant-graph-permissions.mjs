/**
 * Grants Microsoft Graph APPLICATION permissions + admin consent.
 * Run in your own PowerShell/terminal (not embedded IDE shell):
 *   node scripts/grant-graph-permissions.mjs
 */
import { PublicClientApplication } from "@azure/msal-node";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";

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
const graphAppId = "00000003-0000-0000-c000-000000000000";

const requiredRoles = [
  { id: "798ee544-9d2d-430c-a058-570e29e34338", name: "Calendars.Read" },
  { id: "c1684f21-1984-47fa-9d61-2dc8c296bb70", name: "OnlineMeetings.Read.All" },
  {
    id: "a4a80d8d-d283-4bd8-8504-555ec3870630",
    name: "OnlineMeetingTranscript.Read.All",
  },
  {
    id: "a4a08342-c95d-476b-b943-97e100569c8d",
    name: "OnlineMeetingRecording.Read.All",
  },
  { id: "df021288-bdef-4463-88db-98f22de89214", name: "User.Read.All" },
  { id: "b633e1c5-b582-4048-a93e-9f11b44c7e96", name: "Mail.Send" },
  { id: "6b7d71aa-70aa-4810-a8d9-5d9fb2830017", name: "Chat.Read.All" },
  { id: "b9bb2381-47a4-46cd-aafb-00cb12f68504", name: "ChatMessage.Read.All" },
  { id: "44e666d1-d276-445b-a5fc-8815eeb81d55", name: "Tasks.ReadWrite.All" },
];

const pca = new PublicClientApplication({
  auth: {
    clientId: "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
});

console.log("\nSign in as a Global Admin or Privileged Role Admin.\n");

const tokenResult = await pca.acquireTokenByDeviceCode({
  deviceCodeCallback: (response) => {
    console.log(response.message);
  },
  scopes: [
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
  ],
});

const token = tokenResult.accessToken;
const headers = {
  Authorization: `Bearer ${token}`,
  "Content-Type": "application/json",
};

async function graph(path, options = {}) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: { ...headers, ...(options.headers ?? {}) },
  });
  const text = await response.text();
  const body = text ? JSON.parse(text) : null;
  if (!response.ok) {
    throw new Error(
      `${options.method ?? "GET"} ${path} failed (${response.status}): ${text}`,
    );
  }
  return body;
}

const app = (
  await graph(
    `/applications?$filter=appId eq '${clientId}'&$select=id,displayName,requiredResourceAccess`,
  )
).value[0];

if (!app) {
  throw new Error(`App registration not found for client ID ${clientId}`);
}

console.log(`Found app: ${app.displayName}`);

await graph(`/applications/${app.id}`, {
  method: "PATCH",
  body: JSON.stringify({
    requiredResourceAccess: [
      {
        resourceAppId: graphAppId,
        resourceAccess: requiredRoles.map((role) => ({
          id: role.id,
          type: "Role",
        })),
      },
    ],
  }),
});

console.log("Updated requiredResourceAccess on app registration.");

const graphSp = (
  await graph(
    `/servicePrincipals?$filter=appId eq '${graphAppId}'&$select=id,appRoles`,
  )
).value[0];

const appSp = (
  await graph(
    `/servicePrincipals?$filter=appId eq '${clientId}'&$select=id,displayName`,
  )
).value[0];

if (!appSp) {
  console.log(
    "Service principal not found yet. Creating by requesting a token for the app...",
  );
  await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    },
  );
}

const appSpResolved =
  appSp ??
  (
    await graph(
      `/servicePrincipals?$filter=appId eq '${clientId}'&$select=id,displayName`,
    )
  ).value[0];

const existingAssignments = await graph(
  `/servicePrincipals/${appSpResolved.id}/appRoleAssignments`,
);

for (const role of requiredRoles) {
  const alreadyGranted = existingAssignments.value.some(
    (assignment) => assignment.appRoleId === role.id,
  );

  if (alreadyGranted) {
    console.log(`Already granted: ${role.name}`);
    continue;
  }

  await graph(`/servicePrincipals/${appSpResolved.id}/appRoleAssignments`, {
    method: "POST",
    body: JSON.stringify({
      principalId: appSpResolved.id,
      resourceId: graphSp.id,
      appRoleId: role.id,
    }),
  });

  console.log(`Granted admin consent: ${role.name}`);
}

console.log("\nVerification — run: node scripts/diagnose-graph.mjs");
console.log("Then restart: npm run dev\n");

const rl = readline.createInterface({ input, output });
await rl.question("Press Enter to exit...");
rl.close();
