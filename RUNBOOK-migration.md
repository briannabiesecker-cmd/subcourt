# SubCourt — Migration Runbook
*Google Sheet & Account Ownership Transfer*

---

## Overview of moving parts

| Component | Where it lives | What changes on migration |
|---|---|---|
| Google Sheet (data) | Google Drive | File ownership / sharing |
| Apps Script (backend) | Bound to the Sheet | Script ID + Web App URL change |
| GitHub Pages (frontend) | `briannabiesecker-cmd/subcourt` repo | `SCRIPT_URL` constant in HTML |
| Time-based triggers | Google account that ran `setupTriggers()` | Must re-run under new account |
| Config (matching params) | Config tab in the Sheet | Carries over with the Sheet |

---

## Step-by-step migration

### 1. Transfer the Google Sheet

1. Open the Google Sheet
2. **File → Share → Transfer ownership** to the new Google account
3. The new owner accepts ownership from their Google account
4. Confirm the new owner can open the sheet and all tabs are intact:
   - Players, SubRequests, Volunteers, Config, DispatchLog

### 2. Copy the Apps Script project

The Apps Script is bound to the Sheet — it follows the Sheet on transfer. Verify:

1. Open the Sheet → **Extensions → Apps Script**
2. Confirm the new owner sees the full script (all functions)
3. Note the **Script ID** from: Project Settings → IDs (top of left sidebar)

> The Script ID is embedded in the Web App URL. If the project was re-created (not transferred), the Script ID will be different and the Web App URL will change — see Step 4.

### 3. Deploy the Apps Script under the new account

The Web App must be deployed (or re-deployed) by the new owner so it runs under their identity:

1. Open Apps Script editor
2. Click **Deploy → New deployment** (or **Manage deployments → Edit**)
3. Set:
   - Type: **Web app**
   - Execute as: **Me** (new owner's account)
   - Who has access: **Anyone**
4. Copy the new **Web App URL** — format: `https://script.google.com/macros/s/<SCRIPT_ID>/exec`

> If the Script ID did NOT change (ownership transferred, not re-created), the URL may stay the same. Confirm by testing: `<web-app-url>?action=getConfig&callback=test`

### 4. Update `SCRIPT_URL` in the frontend HTML

Open `tennis-sub-manager.html` and find line ~962:

```javascript
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxr-kM_2if-fdkXFu_le83WYsX3MtEVQAh1tR7I9zAO2TKHBtjgVYe_AHIlvpQicqGHsg/exec';
```

Replace the URL with the new Web App URL from Step 3, then commit and push:

```bash
git add tennis-sub-manager.html
git commit -m "Update SCRIPT_URL to new Apps Script deployment"
git push origin main
```

GitHub Pages will redeploy automatically within ~1 minute.

### 5. Update `SHEET_ID` in the Apps Script (if Sheet was re-created)

If the Google Sheet was **copied** (not transferred), it has a new Sheet ID. Open Apps Script and find line ~6:

```javascript
const SHEET_ID = '1X2oM9GwH206qzFHBqi-oIPZ2wPdn43tIfcMRtvOw9wM';
```

Replace with the new Sheet ID (found in the Sheet's URL: `spreadsheets/d/<SHEET_ID>/edit`), then redeploy.

> If the Sheet was transferred (not copied), the Sheet ID stays the same — skip this step.

### 6. Re-register the auto-dispatch triggers

Triggers are tied to the Google account that created them, not the Sheet. The new owner must register their own:

1. Open Apps Script editor (logged in as new owner)
2. In the top toolbar, select function: `setupTriggers`
3. Click **Run**
4. Grant permissions when prompted
5. Verify under **Triggers** (alarm clock icon, left sidebar): you should see two triggers:
   - `runAutoDispatch` — Time-driven, daily at the hour set in Config B14
   - `onConfigEdit` — From spreadsheet, On edit

> The old owner's triggers will become orphaned and stop firing once they lose editor access. They can be manually deleted from the old account's [Apps Script triggers dashboard](https://script.google.com/home/triggers).

### 7. Transfer the GitHub repo (optional)

If the GitHub account changes:

1. Go to `github.com/briannabiesecker-cmd/subcourt` → Settings → Danger Zone → **Transfer**
2. The new owner accepts the transfer
3. Update the local git remote:
   ```bash
   git remote set-url origin https://github.com/<new-org>/subcourt.git
   ```
4. Re-enable GitHub Pages under the new account: Settings → Pages → Source → `main` branch → `/` (root)

---

## Quick-reference: key IDs

### TEST
| Item | Value |
|---|---|
| Google Sheet ID | `1GLWl0a6lRgHsrpG5sZ3S8LtY7HJUGJplNCiPUHIuyIw` |
| Web App URL | `https://script.google.com/macros/s/AKfycby6CoAz-3PE5p5r1Rhi3P9eva8cHu5IzpEe-WqepPJLVN9WMNtBWmDp1Aya_MkFwQO8/exec` |
| GitHub Pages URL | `https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-test.html` |

### PROD
| Item | Value |
|---|---|
| Google Sheet ID | `1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o` |
| Web App URL | `https://script.google.com/macros/s/AKfycbyMaR0EQGmjTrVvjwSqYseZVGPZIaXI0-axH_miCCDgZaQKc4vOVkf4sjD9IDA4Q0Yxnw/exec` |
| GitHub Pages URL | `https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html` |

### Shared
| Item | Value |
|---|---|
| GitHub repo | `briannabiesecker-cmd/subcourt` |

---

## Smoke test after migration

Run these checks to confirm everything is wired up:

- [ ] Open the GitHub Pages URL — app loads without errors
- [ ] Board tab loads player/request data from the Sheet
- [ ] Volunteer tab: submit a test volunteer, confirm it appears in the Volunteers sheet tab
- [ ] Sub Request tab: submit a test request, confirm email arrives
- [ ] Dispatch tab: status banner shows correct enabled/disabled state from Config
- [ ] Config tab: change `autoDispatchEnabled` to FALSE, reload Dispatch tab — banner goes grey
- [ ] Apps Script Triggers panel shows both `runAutoDispatch` and `onConfigEdit`
