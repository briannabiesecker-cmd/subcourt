# Rally — MWF Tennis League Sub Manager

Web app for managing MWF Tennis League match scheduling, sub requests, and player availability. Single prod instance, deployed to GitHub Pages + Google Apps Script.

> Prior to July 2026 there was a separate test/dev instance (`SubCourt-AppScript.js`, `rally-tennis-test.html`, `rally-tennis-dev.html`). Those files were deleted in commit `a7a7428` and never recreated — prod is now the only environment. `deploy.sh` still has a `dev` target referencing the old files; it will fail until/unless that environment is rebuilt.

## Architecture

| Component | Where |
|---|---|
| Frontend | `rally-tennis-prod.html` → GitHub Pages |
| Backend | `SubCourt-AppScript-PROD.js` → Google Apps Script |
| Data | Google Sheets |

Frontend calls Apps Script via JSONP GET (no CORS, no server). All data lives in the Sheet.

## Deploying

**Apps Script (use clasp, not manual paste):**
```bash
bash deploy.sh prod    # pushes to prod Apps Script (auto-substitutes sheet ID)
```
This already deploys to the fixed `PROD_DEPLOYMENT_ID`, which activates it on the web app URL — no manual version bump needed.

**Frontend:** Push to `main` on GitHub. GitHub Pages redeploys in ~1 minute.

## Workflow

1. Iterate directly on `SubCourt-AppScript-PROD.js` / `rally-tennis-prod.html`
2. Run `bash deploy.sh prod`
3. Push to `main` for the frontend

## Sheet structure

**Players sheet:**
| Col | Field |
|---|---|
| A | Name |
| B | Email |
| C | Rating (computed average) |
| D | No8am (boolean — exclude from 8:00 AM slots) |
| E | isAdmin (boolean) |
| F–J | Coordinator rating columns (5 slots, header = coordinator email) |

**Other tabs:** SubRequests, Volunteers, Availability, MatchGroups, Config

**Config tab key cells:**
- B16 = avail window open date
- B17 = avail window close date
- B18 = avail window active flag
- B20–B25 = scheduler weights/iterations/restarts
- B27 = email enabled flag

## Audience

**Non-technical seniors** (tennis league coordinators). Design principles:
- Visual hierarchy guides the user, not instruction text
- Plain language, no jargon ("Responses due by", not "Close date")
- Big tap targets, clear button labels
- Clear loading states — never show stale text or both states at once

## Key conventions

- **JSONP for all API calls** — `apiGetWithParams()` and `apiPost()` in the HTML
- **Single endpoint pattern** — combine related data into one call (e.g. `getSchedulerDashboard`) to minimize round-trips. Apps Script cold start is ~7s.
- **Batch sheet writes** — use `setValues()` not loops of `setValue()`
- **Cache scheduler/players data** in JS module-level vars; don't refetch on every tab switch
- **Captain is always P1** in the schedule output
- Write commits with `Co-Authored-By: Claude` trailer

## Known constraints

- JSONP URL length ~8KB max — chunk large publishes (one slot per request)
- Google Sheets auto-converts date strings to Date objects on write — handle both forms on read
- Disabled buttons don't fire `title` tooltip — wrap in `<span title="...">` instead
- Apps Script `getValue()` returns `''` for empty cells, not `null`/`undefined`
