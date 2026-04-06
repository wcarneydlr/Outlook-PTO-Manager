# PTO Return Assistant — Outlook Add-in

**Version:** 1.0.0  
**Author:** Jeremy Holen / DLR Group  
**Platform:** Outlook Desktop for Windows (Microsoft 365)

---

## What It Does

A two-feature Outlook add-in to help you return from PTO efficiently.

### Feature 1 — Email Triage
- Pick a custom date range (your PTO window)
- Emails are auto-grouped by configurable rules:
  - **Keywords** — e.g. "Renewal", "Quote", "Purchase Order" → Software Renewals group
  - **Project number pattern** — `##-#####-##` (e.g. `26-45123-01`) → DLR Projects group
  - **Team email addresses** — Andy Wallace, Devin Brown, Brian Myers → My Team group
- Each group is collapsible; unread emails are bolded
- Action-needed emails (containing `?`) are tagged
- CC/FYI emails are flagged
- Daily summary: "N emails · X groups · Y with questions · Z CC/FYI"

### Feature 2 — Calendar Sync to Personal Email
- Scans your work calendar for events matching your criteria:
  - Has a Teams/Zoom/Meet/WebEx link
  - Matches Outlook categories (e.g. "Client", "Travel")
  - Falls within a date range
  - Title contains keywords
- Each criteria has its own AND/OR logic toggle
- Matching events are listed with a **"Sync to Personal"** button
- Clicking it opens a new meeting invite pre-filled with:
  - Title, time, location
  - Attendee names (in body — not as actual invitees, so your work colleagues don't see anything)
  - Conference join link
  - Original body preview
- You review and click **Send** — the invite goes to your personal Hotmail/Outlook account only

---

## Setup Instructions

### Step 1 — Host the Files

Upload all files to your GitHub repo (`wcarneydlr.github.io/Carneyasada` or a new repo).

Files to upload:
```
manifest.xml
taskpane.html
commands.html
assets/icon-16.png
assets/icon-32.png
assets/icon-64.png
assets/icon-80.png
```

Enable GitHub Pages on the repo (Settings → Pages → main branch).

### Step 2 — Update manifest.xml

Replace all 6 occurrences of:
```
https://YOUR_HOSTING_URL
```
with your actual GitHub Pages URL, e.g.:
```
https://wcarneydlr.github.io/Carneyasada
```

### Step 3 — Add Icons

Create simple PNG icons in an `assets/` folder. These appear as the ribbon button icon in Outlook.

| File | Size |
|------|------|
| icon-16.png | 16×16 |
| icon-32.png | 32×32 |
| icon-64.png | 64×64 |
| icon-80.png | 80×80 |

You can use any simple icon — a plane ✈ or calendar icon works well.

### Step 4 — Sideload into Outlook

1. Create folder: `C:\OfficeAddins\PTOAddin\`
2. Copy `manifest.xml` into it
3. In Outlook: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
4. Add `C:\OfficeAddins\PTOAddin\` as a catalog, check "Show in Menu"
5. Restart Outlook
6. Open any email → **Home tab** → find "PTO Assistant" button in ribbon

---

## Configuring Your Groups (Settings Tab)

Open the add-in → **⚙ Settings** tab:

- **Keywords** — one per line per group; case-insensitive substring match against email subject
- **Pattern** — regex pattern; the default DLR project pattern is `\d{2}-\d{5}-\d{2}`
- **Team Only** — toggle on to match only emails from addresses listed in "My Team"
- **My Team** — add your teammates' full email addresses

Click **Save Settings** — settings persist across sessions via localStorage.

---

## Notes & Limitations

- **REST API permissions** — the add-in requests `ReadWriteMailbox` to access email and calendar data. On first load in a new Outlook profile you may see a permissions prompt.
- **Demo mode** — if the REST token is unavailable (e.g. during sideload testing without full M365 auth), the add-in shows sample data so you can see the layout.
- **Calendar sync privacy** — the meeting invite is sent only to your personal email. Your work colleagues and the original organizer see nothing. The invite is created fresh (not forwarded), so there's no "forwarded by" trail.
- **No auto-archive, no auto-reply, no AI reordering** — consistent with the "Organize, Don't Prioritize" design philosophy from the original spec.
