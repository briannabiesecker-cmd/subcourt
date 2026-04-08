# Rally — MWF Tennis League Sub Manager

A web app for managing MWF Tennis League match scheduling, sub requests, and player availability.

**Live app:** https://briannabiesecker-cmd.github.io/subcourt/tennis-sub-manager.html
**GitHub repo:** https://github.com/briannabiesecker-cmd/subcourt
**Google Sheet:** https://docs.google.com/spreadsheets/d/1GLWl0a6lRgHsrpG5sZ3S8LtY7HJUGJplNCiPUHIuyIw

---

## Architecture

| Component | Where it lives |
|---|---|
| Frontend (HTML/JS) | GitHub Pages — `tennis-sub-manager.html` |
| Backend | Google Apps Script web app (bound to the Sheet) |
| Data | Google Sheet — 6 tabs (see below) |

The frontend calls the Apps Script via JSONP GET requests (no server, no CORS issues). All data lives in the Google Sheet.

### Google Sheet tabs

| Tab | Purpose |
|---|---|
| **Players** | Player roster with skill ratings |
| **SubRequests** | Incoming sub requests |
| **Volunteers** | Volunteer availability submissions |
| **Availability** | Monthly availability submissions from players |
| **MatchGroups** | Published monthly schedules |
| **Config** | App settings and availability window state |

---

## What the app does

### End-user tabs
- **Home** — Overview and quick links
- **Request Sub** — Submit a sub request for a match; look up past requests by name
- **Volunteer** — Submit dates you're available to sub; look up past submissions by name
- **Availability** — Submit your available match dates for the upcoming month
- **Schedule** — View the published monthly schedule; type your name to filter to just your match dates

### Admin tab
Requires OTP login via email. Three sub-tabs:

| Sub-tab | Purpose |
|---|---|
| **Scheduler** | Manage the availability window + generate and publish the monthly schedule |
| **Player Profiles** | View and edit player skill ratings |
| **Dispatch** | Manual sub request management (auto-dispatch currently disabled) |

---

## Admin: Monthly Scheduling Workflow

This is the main recurring task — done once per month.

### Step 1 — Open the availability window

1. Log in to the **Admin** tab
2. Go to **Scheduler** (loads automatically)
3. Set the **Close date** — the last day players can submit their availability
4. Click **Open Window & Notify Players**
   - The window opens today
   - All players on the roster receive an email notification
   - The status badge turns green: **Window is OPEN**

> **Automatic reminders are sent:**
> - T-2 days before close: reminder email to players who haven't submitted yet
> - T-1 day before close: final reminder to players who still haven't submitted
> - The window closes automatically on the close date (no action needed)

---

### Step 2 — Generate the schedule

Once the window closes (or whenever you're ready):

1. Go to **Admin → Scheduler**
2. The **Generate Schedule** card shows how many players submitted and for which month
3. Click **Generate**
   - The scheduler runs a local search optimization (~10–30 seconds)
   - Groups players into sets of 4 (one sit-out if not divisible by 4)
   - Balances skill levels across groups
   - Minimizes repeat pairings across the month
   - Assigns one **Captain** per group (~25% of each player's scheduled dates)
4. The **Schedule Preview** table appears showing all dates and groups

**Reading the table:**

| Column | What it shows |
|---|---|
| Date | Match date |
| Group | Group A, B, C… |
| Captain | Player with [C] badge — responsible for coordinating that session |
| Players | Remaining 3 players in the group |
| Sit Out | Player sitting out that date (if roster isn't divisible by 4) |

---

### Step 3 — Review and export (optional)

Before publishing, you can:

- **Export to Excel** — Pivot-format spreadsheet: players as rows, dates as columns, group letter in each cell. Captains shown as `A [C]`. Formatted with alternating row shading and frozen header.
- **Export to PDF** — Print-ready version of the schedule table

You can re-run **Generate** as many times as you like — nothing is saved until you click **Publish**.

---

### Step 4 — Publish

Click **Publish** to save the schedule to the Google Sheet.

- Existing rows for that month are cleared first (safe to re-publish)
- Each group row is written to the **MatchGroups** tab with the captain in the P1 position
- Players can immediately see the schedule on the **Schedule** tab of the app

---

## Admin: Player Profiles

Use the **Player Profiles** tab to view and update skill ratings.

- Ratings are used internally by the scheduler to balance groups by skill level
- Ratings are never shown to end users

---

## Email notifications summary

| Trigger | Who receives it |
|---|---|
| Admin clicks **Open Window** | All players on the roster |
| T-2 days before close date | Players who haven't submitted yet |
| T-1 day before close date | Players who haven't submitted yet |

> **Note:** `EMAIL_ENABLED` in `SubCourt-AppScript.js` must be set to `true` for emails to send. It defaults to `false` for testing. Admin OTP login emails always send regardless of this flag.

---

## Deployment

### After any Apps Script change

1. Open the Google Sheet → **Extensions → Apps Script**
2. Paste in the updated `SubCourt-AppScript.js`
3. Click **Deploy → Manage deployments → Edit** (pencil icon on the current deployment)
4. Bump the version (select "New version" from the dropdown)
5. Click **Deploy** — the Web App URL stays the same

### After any HTML change

Push to `main` on GitHub. GitHub Pages redeploys automatically within ~1 minute.

### One-time trigger setup (new account or fresh script)

Run `setupTriggers()` once from the Apps Script editor to install:
- Daily check that auto-closes the availability window when the close date passes
- Daily T-2 / T-1 reminder emails
- Monthly cleanup of old availability records
- Config tab onEdit watcher

---

## Key constants (Apps Script)

```javascript
const SHEET_ID     = '1GLWl0a6lRgHsrpG5sZ3S8LtY7HJUGJplNCiPUHIuyIw';
const EMAIL_ENABLED = false; // set to true in production
```

## Key constant (HTML)

```javascript
const SCRIPT_URL = 'https://script.google.com/macros/s/<deployment-id>/exec';
```

See `RUNBOOK-migration.md` for full instructions on transferring the app to a new Google account or GitHub repo.
