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

const organizerId = "25d32a6a-e4f3-47b3-942c-bafe91941ea8";
const joinUrl =
  "https://teams.microsoft.com/l/meetup-join/19%3ameeting_NzA1YzE1OTgtOTNkYS00N2Q0LWFkNWMtZmMxZjk5MDUzNDQ2%40thread.v2/0?context=%7b%22Tid%22%3a%22e60d4ac4-0106-4e11-8bdf-5715d69670b1%22%2c%22Oid%22%3a%2225d32a6a-e4f3-47b3-942c-bafe91941ea8%22%7d";

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
const escaped = joinUrl.replace(/'/g, "''");
const omRes = await fetch(
  `https://graph.microsoft.com/v1.0/users/${organizerId}/onlineMeetings?$filter=JoinWebUrl eq '${escaped}'`,
  { headers },
);
const om = await omRes.json();
const meetingId = om.value[0].id;

const trRes = await fetch(
  `https://graph.microsoft.com/v1.0/users/${organizerId}/onlineMeetings/${meetingId}/transcripts`,
  { headers },
);
const tr = await trRes.json();
const transcriptId = tr.value[0].id;

const badRes = await fetch(
  `https://graph.microsoft.com/v1.0/users/${organizerId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`,
  { headers: { ...headers, Accept: "*/*" } },
);
console.log("Accept */*:", badRes.status, (await badRes.json()).error?.message);

const goodRes = await fetch(
  `https://graph.microsoft.com/v1.0/users/${organizerId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content?$format=text/vtt`,
  { headers: { ...headers, Accept: "text/vtt" } },
);
const text = await goodRes.text();
console.log("Accept text/vtt:", goodRes.status, text.slice(0, 80));
