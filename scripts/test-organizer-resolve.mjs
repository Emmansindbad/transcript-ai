import { ConfidentialClientApplication } from "@azure/msal-node";
import { readFileSync } from "node:fs";

const env = Object.fromEntries(
  readFileSync(".env.local", "utf8")
    .split("\n")
    .filter((line) => line.includes("=") && !line.startsWith("#"))
    .map((line) => {
      const index = line.indexOf("=");
      return [line.slice(0, index), line.slice(index + 1)];
    }),
);

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}`,
    clientSecret: env.AZURE_CLIENT_SECRET,
  },
});

const token = (
  await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  })
).accessToken;

const headers = { Authorization: `Bearer ${token}` };
const joinUrl =
  "https://teams.microsoft.com/l/meetup-join/19%3ameeting_NzA1YzE1OTgtOTNkYS00N2Q0LWFkNWMtZmMxZjk5MDUzNDQ2%40thread.v2/0?context=%7b%22Tid%22%3a%22e60d4ac4-0106-4e11-8bdf-5715d69670b1%22%2c%22Oid%22%3a%2225d32a6a-e4f3-47b3-942c-bafe91941ea8%22%7d";

const organizerId = JSON.parse(
  decodeURIComponent(new URL(joinUrl).searchParams.get("context") ?? "{}"),
).Oid;

const upn = env.TARGET_USER_ID;

const userRes = await fetch(
  `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${upn.replace(/'/g, "''")}'&$select=id,userPrincipalName,displayName`,
  { headers },
);
const userList = await userRes.json();
const user = userList.value?.[0];
console.log("Calendar user:", user?.id, user?.displayName);
console.log("Organizer from URL:", organizerId);
console.log("Same person?", user?.id === organizerId);

for (const uid of [user?.id, organizerId].filter(Boolean)) {
  const filter = encodeURIComponent(`JoinWebUrl eq '${joinUrl.replace(/'/g, "''")}'`);
  const url = `https://graph.microsoft.com/v1.0/users/${uid}/onlineMeetings?$filter=${filter}`;
  const response = await fetch(url, { headers });
  const body = await response.json();
  console.log(
    `Try user ${uid}:`,
    response.status,
    "meetings:",
    body.value?.length ?? 0,
    body.error?.code ?? "ok",
  );

  if (body.value?.[0]?.id) {
    const tr = await fetch(
      `https://graph.microsoft.com/v1.0/users/${uid}/onlineMeetings/${body.value[0].id}/transcripts`,
      { headers },
    );
    const trBody = await tr.json();
    console.log("  transcripts:", tr.status, trBody.value?.length ?? 0);
  }
}
