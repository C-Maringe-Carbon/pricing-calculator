# Deployment Guide — Pricing Calculator

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
This produces: `sharepoint/solution/pricing-calculator.sppkg`

## 3. Upload the HTML file to SharePoint
1. Go to: `https://carbongroup.sharepoint.com/sites/CarbonMarketingAssets`
2. Open **Site Contents → Site Assets**
3. Upload `pricing-calculator (2).html`, rename it to **`pricing-calculator.html`**

## 4. Deploy the .sppkg to the App Catalog
1. Go to your SharePoint Admin Center → **More features → Apps → App Catalog**
   - If no App Catalog exists yet, create one scoped to the CarbonMarketingAssets site
2. Upload `sharepoint/solution/pricing-calculator.sppkg`
3. Click **Deploy** (trust the solution)

## 5. Add the web part to a page
1. Open or create a page on `CarbonMarketingAssets`
2. Edit the page → add web part → search **Pricing Calculator**
3. The calculator loads automatically from `SiteAssets/pricing-calculator.html`

## How it works
- The web part renders an iframe pointing to the HTML file in Site Assets
- Because the HTML is served from the same SharePoint tenant, the browser session cookie
  is shared — no extra login needed
- The SharePoint REST API call to the `PricingServices` list works automatically
- `SP_SITE_URL` is set to `https://carbongroup.sharepoint.com/sites/CarbonMarketingAssets`

## SharePoint list schema (PricingServices)
| Column          | Type                    | Notes                        |
|-----------------|-------------------------|------------------------------|
| Title           | Single line of text     | Section: "Year End", "Bookkeeping", etc. |
| ServiceType     | Single line of text     | Service display name         |
| Unit            | Single line of text     | e.g. "Per BAS"               |
| GoodRate        | Number                  | Base price                   |
| BadRate         | Number                  | Tidy-up price                |
| AnnualMultiplier| Number                  | Leave blank for 1x           |
| Variants        | Multiple lines of text  | JSON: `[{"label":"X","price":900}]` |
