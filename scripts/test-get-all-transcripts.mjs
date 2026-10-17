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

const end = new Date();
const start = new Date();
start.setDate(start.getDate() - 30);

const url =
  `https://graph.microsoft.com/v1.0/users/${organizerId}/onlineMeetings/getAllTranscripts` +
  `(meetingOrganizerUserId='${organizerId}',startDateTime=${start.toISOString()},endDateTime=${end.toISOString()})`;

const response = await fetch(url, {
  headers: { Authorization: `Bearer ${token}` },
});
const body = await response.json();
console.log("status", response.status);
console.log("count", body.value?.length ?? 0);
if (body.value?.[0]) {
  console.log("sample", {
    meetingId: body.value[0].meetingId?.slice(0, 40) + "...",
    createdDateTime: body.value[0].createdDateTime,
  });
}
if (body.error) console.log("error", body.error);
