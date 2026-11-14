# M365 Transcript AI Analyzer

Internal SindbadTech dashboard that reads Microsoft Teams meetings from Graph, then uses Gemini to produce ticket-shaped action items (Task vs Bug), a meeting summary, and optional email / Microsoft To Do export.

Times in the UI and ticket day-matching use **Arabia Standard Time (KSA, Asia/Riyadh)**.

## What it does

1. Lists recent calendar Teams meetings (last 30 days), with **Transcript** / **Recording** badges.
2. Loads spoken transcript (VTT with timestamps), same-day meeting chat, and optional recording.
3. Analyzes with Gemini **3.8 Flash**, then **3.6 Flash** if 3.8 is busy. Set `GEMINI_MODEL=gemini-3.1-pro-preview` to try Pro first.
4. Lets you **edit tickets** before sending. Already-filed work is skipped by default.
5. Emails `SUMMARY_EMAIL_RECIPIENT` (default `mom@sindbad.tech`) and/or creates Microsoft To Do tasks.
6. Saves analyses under `data/analyzer-store.json` so refresh does not lose tickets.

## Run locally

```bash
npm install
copy .env.example .env.local
npm run dev
```

Open http://localhost:3000. Click **Fetch Data**, pick a meeting, then **Generate Analysis**.

## Azure / Graph setup

App-only **client credentials**. Never put secrets in `NEXT_PUBLIC_*`.

Required application permissions (admin consent):

- `Calendars.Read`
- `OnlineMeetings.Read.All`
- `OnlineMeetingTranscript.Read.All`
- `OnlineMeetingRecording.Read.All`
- `User.Read.All`
- `Mail.Send`
- `Chat.Read.All`
- `ChatMessage.Read.All`
- `Tasks.ReadWrite.All`

Grant them with:

```powershell
.\scripts\grant-graph-permissions.ps1
```

Also required:

1. An **Application Access Policy** that allows this app to read online meetings for `TARGET_USER_ID` (and organizers you care about).
2. Teams tenant setting **Graph access to transcripts** on:

```powershell
.\scripts\enable-graph-transcript-access.ps1
```

`TARGET_USER_ID` is the mailbox used as the email sender and the primary calendar identity.

## Gemini

Set `GEMINI_API_KEY` from [Google AI Studio](https://aistudio.google.com/apikey). Optional `GEMINI_MODEL` caps the fallback list.

Recordings larger than **150 MB** are skipped so the Node process does not load the whole file into memory. Extra recorded sessions that are not on the calendar (same Teams series, different day) still appear in Fetch Data.

## Health

The header **Graph health** chips probe calendar, transcripts, recordings, chat, mail, and To Do. Expand them when Fetch Data looks empty.
