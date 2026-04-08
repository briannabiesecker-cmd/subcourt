# Rally — MTC Tennis Sub Manager

Web app for managing MTC tennis team match scheduling, sub requests, and player availability. Two instances (test/prod), shared codebase, deployed to GitHub Pages + Google Apps Script.

## Architecture

| Component | Where |
|---|---|
| Frontend (test) | `rally-tennis-test.html` → GitHub Pages |
| Frontend (prod) | `rally-tennis-prod.html` → GitHub Pages |
| Backend | `SubCourt-AppScript.js` → Google Apps Script (one per env) |
| Data | Google Sheets (one per env) |

Frontend calls Apps Script via JSONP GET (no CORS, no server). All data lives in the Sheet.

## Test vs Prod differences (DO NOT port these between files)

- Favicon (clay tennis ball in test, default emoji in prod)
- Sticky test banner + lime header in test
- `Request Sub` / `Volunteer` buttons enabled in test, disabled in prod
- Different `SCRIPT_URL` constants (different Apps Script deployments)
- Test sheet ID vs prod sheet ID (handled by `deploy.sh`)

Everything else should be identical between the two HTML files.

## Deploying

**Apps Script (use clasp, not manual paste):**
```bash
bash deploy.sh test    # pushes to test Apps Script
bash deploy.sh prod    # pushes to prod Apps Script (auto-substitutes sheet ID)
```
After clasp push, **bump the deployment version** in the Apps Script editor (Deploy → Manage deployments → Edit → New version → Deploy) — clasp updates code but doesn't activate it on the web app URL.

**Frontend:** Push to `main` on GitHub. GitHub Pages redeploys in ~1 minute.

## Workflow

1. Iterate on `rally-tennis-test.html` + `SubCourt-AppScript.js`
2. Test on the test instance
3. When ready, package an "update set" → port functional changes to `rally-tennis-prod.html`
4. Run `bash deploy.sh prod` and bump prod deployment version

Verify diffs after porting:
```bash
diff rally-tennis-prod.html rally-tennis-test.html
```
Only test-specific items should remain (favicon, banner, disabled buttons, SCRIPT_URL).

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
