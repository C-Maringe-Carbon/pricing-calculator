# Carbon Group — AI Committee

Standalone HTML app + SPFx web part for the AI Committee SharePoint site.

---

## What's included

| File | Purpose |
|---|---|
| `carbon_group_committee.html` | Standalone app — works offline, has Export button |
| `src/webparts/aiCommittee/CarbonAiCommitteeWebPart.ts` | SPFx web part (iframe loader) |
| `sharepoint-lists/site-script.json` | List schema — 1 list, 13 columns |
| `sharepoint-lists/Seed-Lists.ps1` | Creates the list and loads all 112 tasks in one command |
| `sharepoint-lists/AI_Committee_Tasks.csv` | All 112 tasks (CSV fallback for manual import) |
| `sharepoint-lists/AI_Committee_Projects.csv` | 16 projects / task groups |
| `sharepoint-lists/AI_Committee_Pillars.csv` | 4 pillars with descriptions |
| `sharepoint-lists/AI_Committee_Agenda.csv` | 8 standard monthly agenda items |

---

## Deployment order

### 1. Install dependencies
```bash
npm install
```

### 2. Build & package the .sppkg
```bash
npm run package
```
Produces: `sharepoint/solution/carbon-group-committee.sppkg`

### 3. Deploy the .sppkg to the App Catalog
1. Go to SharePoint Admin Center → **More features → Apps → App Catalog**
2. Upload `sharepoint/solution/carbon-group-committee.sppkg`
3. Click **Deploy** (trust the solution)

### 4. Upload the HTML to Site Assets
1. Go to your target SharePoint site (e.g. `SystemCommittee`)
2. Open **Site Contents → Site Assets**
3. Upload `carbon_group_committee.html`

### 5. Seed the SharePoint list
Requires [PnP PowerShell](https://pnp.github.io/powershell/). Install once if needed:
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```
Then run:
```powershell
.\sharepoint-lists\Seed-Lists.ps1 -SiteUrl "https://carbongroup.sharepoint.com/sites/SystemCommittee"
```
This will:
- Create the `AI_Committee` list with all 13 columns
- Load all 112 tasks, correctly ordered by pillar and project

### 6. Add the web part to a page
1. Edit any page on the site
2. Add web part → search **AI Committee**
3. The app loads from `SiteAssets/carbon_group_committee.html`

### 7. Pin as a Teams tab (optional)
In Teams → navigate to the SharePoint page → **⋯ → Add tab → SharePoint**

---

## List schema — AI_Committee

| Column | Type | Notes |
|---|---|---|
| Title | Single line | Task description |
| Pillar | Single line | Dynamics CRM / Process & System Automation / Integration with Government Agencies / Technical Know-How |
| PillarID | Single line | p1 / p2 / p3 / p4 |
| PillarOrder | Number | 1–4, controls sort order |
| PillarChampion | Single line | Person responsible for the pillar |
| Project | Single line | Task group name |
| Phase | Single line | Month 1 — Foundation / Month 2 — Core Engine / etc. |
| ProjectOrder | Number | 1–16, controls sort order within pillars |
| TaskOrder | Number | Sort order within each project |
| Done | Yes/No | Completion state |
| CompletedDate | Date and Time | Set when Done is ticked |
| Owner | Single line | Person responsible for this specific task |
| Notes | Multiple lines | Any additional context |

---

## How it works

- The web part renders an iframe pointing to `SiteAssets/carbon_group_committee.html`
- Because the HTML is served from the same SharePoint tenant, the session cookie is shared — no extra login
- The HTML is self-contained and fully editable in-browser via the **Enable editing** toggle
- The **Export HTML** button saves the current state (including any edits) as a new `.html` file

---

## Document owner
Nathan (Carbon Group) · Questions: raise with Nathan before proceeding
