# Deployment Guide — Carbon Group AI Committee

## Prerequisites
- Node.js 18.x (`node --version` should show v18.x)
- Run once: `npm install -g gulp-cli`

## 1. Install dependencies
```bash
npm install
```

## 2. Build & package the .sppkg
```bash
npm run package
```
This produces: `sharepoint/solution/carbon-group-committee.sppkg`

## 3. Upload the HTML file to SharePoint
1. Go to: `https://carbongroup.sharepoint.com/sites/CarbonMarketingAssets`
2. Open **Site Contents → Site Assets**
3. Upload `carbon_group_committee.html`

## 4. Deploy the .sppkg to the App Catalog
1. Go to your SharePoint Admin Center → **More features → Apps → App Catalog**
   - If no App Catalog exists yet, create one scoped to the CarbonMarketingAssets site
2. Upload `sharepoint/solution/carbon-group-committee.sppkg`
3. Click **Deploy** (trust the solution)

## 5. Add the web part to a page
1. Open or create a page on `CarbonMarketingAssets`
2. Edit the page → add web part → search **AI Committee**
3. The page loads automatically from `SiteAssets/carbon_group_committee.html`

## How it works
- The web part renders an iframe pointing to the HTML file in Site Assets
- Because the HTML is served from the same SharePoint tenant, the browser session
  cookie is shared — no extra login needed
- `SP_SITE_URL` is set to `https://carbongroup.sharepoint.com/sites/CarbonMarketingAssets`
