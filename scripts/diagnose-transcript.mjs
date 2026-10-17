import { ConfidentialClientApplication } from "@azure/msal-node";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

function loadEnv() {
  const envPath = resolve(process.cwd(), ".env.local");
  for (const line of readFileSync(envPath, "utf8").split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    process.env[trimmed.slice(0, eq)] = trimmed.slice(eq + 1);
  }
}

loadEnv();

const userId = process.env.TARGET_USER_ID;
const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
});

const token = (
  await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  })
).accessToken;

const headers = { Authorization: `Bearer ${token}` };

const end = new Date();
const start = new Date();
start.setDate(start.getDate() - 30);

const calUrl =
  `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/calendarView` +
  `?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$top=5&$select=id,subject,isOnlineMeeting,onlineMeeting`;

const calRes = await fetch(calUrl, {
  headers: { ...headers, Prefer: 'outlook.timezone="UTC"' },
});
const cal = await calRes.json();

console.log("CALENDAR_STATUS", calRes.status);
const onlineEvents = (cal.value ?? []).filter((e) => e.isOnlineMeeting && e.onlineMeeting?.joinUrl);

if (onlineEvents.length === 0) {
  console.log("No online meetings in calendar sample");
  process.exit(0);
}

for (const event of onlineEvents.slice(0, 3)) {
  const joinUrl = event.onlineMeeting.joinUrl;
  console.log("\n--- EVENT ---");
  console.log("Subject:", event.subject);
  console.log("Event ID:", event.id);
  console.log("Join URL:", joinUrl);

  const organizerId = (() => {
    try {
      const context = new URL(joinUrl).searchParams.get("context");
      if (!context) return null;
      return JSON.parse(decodeURIComponent(context)).Oid ?? null;
    } catch {
      return null;
    }
  })();

  console.log("Organizer OID from URL:", organizerId);

  for (const testUserId of [userId, organizerId].filter(Boolean)) {
    const escaped = joinUrl.replace(/'/g, "''");
    const omUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(testUserId)}/onlineMeetings` +
      `?$filter=JoinWebUrl eq '${escaped}'`;

    const omRes = await fetch(omUrl, { headers });
    const om = await omRes.json();
    console.log(`ONLINE_MEETINGS as ${testUserId}:`, omRes.status, om.value?.length ?? 0, om.error?.code ?? "ok");
    if (om.value?.[0]?.id) {
      const meetingId = om.value[0].id;
      const trUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(testUserId)}/onlineMeetings/${meetingId}/transcripts`;
      const trRes = await fetch(trUrl, { headers });
      const tr = await trRes.json();
      console.log("TRANSCRIPTS:", trRes.status, tr.value?.length ?? 0, tr.error?.code ?? "ok");
    }
  }

  // skip old block below
  continue;
  const escaped = joinUrl.replace(/'/g, "''");
  const omUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/onlineMeetings` +
    `?$filter=JoinWebUrl eq '${escaped}'`;

  const omRes = await fetch(omUrl, { headers });
  const om = await omRes.json();
  console.log("ONLINE_MEETINGS_FILTER_STATUS", omRes.status);
  console.log("ONLINE_MEETINGS_COUNT", om.value?.length ?? 0);
  if (om.error) console.log("OM_ERROR", om.error);

  if (om.value?.[0]?.id) {
    const meetingId = om.value[0].id;
    const trUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/onlineMeetings/${meetingId}/transcripts`;
    const trRes = await fetch(trUrl, { headers });
    const tr = await trRes.json();
    console.log("TRANSCRIPTS_STATUS", trRes.status);
    console.log("TRANSCRIPT_COUNT", tr.value?.length ?? 0);
    if (tr.error) console.log("TR_ERROR", tr.error);
  }

  const eventUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/events/${event.id}?$select=onlineMeeting,subject`;
  const evRes = await fetch(eventUrl, { headers });
  const ev = await evRes.json();
  console.log("EVENT_DETAIL_STATUS", evRes.status);
  console.log("EVENT_ONLINE_MEETING", JSON.stringify(ev.onlineMeeting, null, 2));
}
