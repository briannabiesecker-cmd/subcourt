# Rally Tennis — Developer Notes

## Key IDs

| Environment | Sheet ID |
|---|---|
| DEV | `1VjFuq63KLEgZpYvCVi2bJrWEgMxDP6hXygYwjDpUmRE` |
| PROD | `1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o` |

| Environment | Apps Script Deployment ID |
|---|---|
| DEV | `AKfycbxS8vYTuuuxsjbVoLS0Mup8VYiCj0t95N6dq7cCKIimnwfLW4or5qBoGFHGbVZIT597Ug` |

| Environment | Apps Script Project ID |
|---|---|
| DEV | `1eSjMqsbLquKowkSPe1O-KIkG0YXP2lAojDqKYm_3tVOBSgOBZYJH8Iki` |

## URLs

| Environment | App URL |
|---|---|
| DEV | https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-dev.html |
| PROD | https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html |

## Local File Locations

| File | Purpose |
|---|---|
| `C:\Users\marob\subcourt\SubCourt-AppScript.js` | DEV Apps Script (working copy) |
| `C:\Users\marob\subcourt\SubCourt-AppScript-PROD.js` | PROD Apps Script (snapshot of live production) |
| `C:\Users\marob\subcourt\rally-tennis-dev.html` | DEV frontend |
| `C:\Users\marob\subcourt\rally-tennis-prod.html` | PROD frontend |
| `C:\Users\marob\subcourt-dev-script\` | clasp folder for pushing to DEV Apps Script |

---

## After Every Claude Code Session — Verify SHEET_ID

Claude Code may accidentally change the SHEET_ID. Always check after a session:

1. Open `SubCourt-AppScript.js` in VS Code
2. Press **Ctrl+G** and go to line 6
3. Confirm it shows the DEV Sheet ID:
   ```javascript
   const SHEET_ID = '1VjFuq63KLEgZpYvCVi2bJrWEgMxDP6hXygYwjDpUmRE';
   ```
4. If it's wrong, fix it, save, and push via clasp (see below)

---

## Pushing Apps Script Changes to DEV

After editing `SubCourt-AppScript.js` in VS Code:

```
cd C:\Users\marob\subcourt-dev-script
copy C:\Users\marob\subcourt\SubCourt-AppScript.js C:\Users\marob\subcourt-dev-script\SubCourt-AppScript.js
clasp push --force
clasp deploy --deploymentId AKfycbxS8vYTuuuxsjbVoLS0Mup8VYiCj0t95N6dq7cCKIimnwfLW4or5qBoGFHGbVZIT597Ug --description "DEV"
```

---

## Pushing HTML Changes to DEV

After editing `rally-tennis-dev.html` in VS Code:

1. Go to **Source Control** in VS Code (branch icon in left sidebar)
2. Type a commit message describing the change
3. Click **Commit**
4. Click **Sync Changes**
5. Wait 1-2 minutes for GitHub Pages to refresh
6. Verify at: https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-dev.html

---

## Deploying DEV to PROD

### Step 1 — Compare and review changes
In Claude Code, type:
> "Compare SubCourt-AppScript.js and SubCourt-AppScript-PROD.js and tell me what changed"

and/or:
> "Compare rally-tennis-dev.html and rally-tennis-prod.html and tell me what changed"

### Step 2 — Deploy Apps Script to PROD (if changed)
1. Open `SubCourt-AppScript.js` in VS Code
2. Change line 6 SHEET_ID to the PROD value:
   ```javascript
   const SHEET_ID = '1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o';
   ```
3. Copy all the code
4. Open the PROD Apps Script in the browser:
   https://script.google.com/home
5. Open **PROD Rally Tennis Team Manager** project
6. Select all and paste the updated code
7. Confirm SHEET_ID is the PROD value
8. Click **Deploy → Manage Deployments → pencil icon → New version → Deploy**
9. Change line 6 back to the DEV SHEET_ID in VS Code and save

### Step 3 — Deploy HTML to PROD (if changed)
1. Open `rally-tennis-prod.html` in VS Code
2. Apply the same changes you made in `rally-tennis-dev.html`
3. Save with **Ctrl+S**
4. Go to Source Control → commit message → **Commit** → **Sync Changes**
5. Verify at: https://briannabiesecker-cmd.github.io/subcourt/rally-tennis-prod.html

### Step 4 — Update the PROD snapshot
After deploying, update the PROD snapshot so future comparisons are accurate:
```
copy C:\Users\marob\subcourt\SubCourt-AppScript.js C:\Users\marob\subcourt\SubCourt-AppScript-PROD.js
```
Then open `SubCourt-AppScript-PROD.js` and change line 6 to the PROD SHEET_ID:
```javascript
const SHEET_ID = '1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o';
```
Save, commit and sync to GitHub.

---

## Starting a Claude Code Session

Always begin a Claude Code session with this instruction to prevent changes to PROD files:

> "In this session only make changes to rally-tennis-dev.html and SubCourt-AppScript.js. Do not touch rally-tennis-prod.html, rally-tennis-test.html, or SubCourt-AppScript-PROD.js"
