# STR Revenue Monitor

A one-stop **pacing and revenue** dashboard for short-term rentals managed with **Hospitable** and **PriceLabs**. Use it to see how you’re tracking vs target, spot gaps, and get simple recommendations.

## What it does

- **Pacing** — Revenue and occupancy vs your target for the current month/quarter/year. “On pace” / “Behind” status.
- **Revenue metrics** — Revenue, ADR, occupancy %, RevPAR, nights booked for the selected period.
- **Future unbooked periods (market pacing)** — For each week with unbooked nights: **market current occupancy** (how full the market is today for those dates), **market final expected occupancy** (based on last year), and **remaining potential** (how much more demand is likely). From that, the app suggests a **price percentile** (e.g. “price at 75–100%” when final occupancy is high).
- **Configurable pacing rules** — Columns: **Final Occ. %**, **Target Pctl %**, **Lower Pctl %**, **Wks to Lower**, **Lowest Pctl%**, **Wks to Lowest**. When LY final ≥ Final Occ., that row applies. Ramp **Target → Lower** by **Wks to Lower** weeks out, then **Lower → Lowest** by stay; **Wks to Lowest** is when the ramp starts from Target. **Wks to Lower** = 0 (or invalid) → one-step Target → Lowest. **Wks to Lowest** = 0 → always Target. **0%** Final Occ. = catch-all.
- **Gap analysis** — Unbooked dates in the next 90 days and **pricing change recommendations** (e.g. “Week of 3/17: market final 72% → price at 50–75 percentile”).

## Quick start

1. Open `index.html` in a browser (double-click or drag into Chrome/Edge).
2. Click **Use sample data** to see the dashboard with demo reservations.
3. Set a **Revenue target** and click **Save target** to see pacing vs that target.
4. Switch **Pacing period** (month / quarter / year) to change the comparison window.
5. Click **Pacing settings** to edit pacing rows. **Save as defaults** stores the current table as what **Reset to defaults** will reload (separate from **Save settings**, which applies the live dashboard config).

## Your own data (CSV)

Export reservations or calendar from Hospitable (or any tool) as CSV. The app looks for columns like:

- **Check-in** — `check_in`, `check-in`, `arrival`, `start`
- **Check-out** — `check_out`, `check-out`, `departure`, `end`
- **Revenue** (optional) — `revenue`, `payout`, `total`, `amount`, `earnings`

Use **Upload calendar / revenue CSV** to load the file. Pacing and metrics will use this data.

## Connecting Hospitable & PriceLabs (later)

### Hospitable API

- **Docs:** [developer.hospitable.com](https://developer.hospitable.com/docs/public-api-docs/d862b3ee512e6-introduction) and [help.hospitable.com – API](https://help.hospitable.com/en/articles/4607709-integrate-with-our-api).
- **Auth:** Personal Access Token (PAT). In Hospitable: **Settings → Integrations → API access → Add new** (read access is enough for this monitor).
- **Use for:** Reservations (check-in, check-out, payout), calendar, property list. You can replace the CSV flow with a small backend or serverless function that calls the API and returns the same reservation shape (`checkIn`, `checkOut`, `revenue`).

### PriceLabs

- **APIs:** Dynamic Pricing API (for PMS integration) and Revenue Estimator API. Access often requires contacting [support@pricelabs.co](mailto:support@pricelabs.co) and domain/plan setup.
- **Use for:** **Market occupancy** (current and final expected by period) and **price percentiles**. Right now the app uses **sample market data** for “current” and “final expected” occupancy by week. When PriceLabs API or exports provide this data, plug it in so the “Future unbooked periods” table and pricing recommendations use real market pacing and percentile suggestions.

### Making it “live”

To turn this into a live one-stop shop:

1. Add a small **backend** (e.g. Node, Python, or serverless) that:
   - Calls Hospitable API with your PAT and fetches reservations.
   - Optionally calls PriceLabs if you have access, for targets or recommendations.
2. Have the backend return JSON in the same format the app expects (list of `{ checkIn, checkOut, revenue }`).
3. In the app, replace the CSV upload / sample data with a “Refresh” button that fetches from your backend.

The dashboard logic (pacing, metrics, gaps, recommendations) can stay as-is; only the data source changes.

## Files

- `index.html` — Dashboard layout and sections.
- `styles.css` — Layout and theme.
- `app.js` — Pacing, metrics, future-period market pacing, configurable pacing rows, unbooked dates, pricing recommendations, CSV parsing, sample data.
- `README.md` — This file.

## Pacing in plain English

- You set a **revenue target** for the period (e.g. $4,500 for the month).
- The app compares **actual revenue so far** to **how much of the period has elapsed**.
  - Example: 10 days into a 30-day month → you’ve had 33% of the period. To be on pace, you’d want at least ~33% of your target (e.g. $1,485 if target is $4,500).
- **On pace** = actual revenue is at or above that implied target. **Behind** = below; the recommendations suggest checking PriceLabs and minimum stays to improve pacing.

**Market-based pricing:** For future unbooked dates, the app compares the market’s **current** occupancy (how booked the market is today for that period) to the market’s **final expected** occupancy (e.g. from last year). The gap is “remaining potential” demand. Based on your **Pacing settings** (threshold rows), it suggests a **target price percentile** for each period. Later you can replace the sample market data with PriceLabs (or another source) for live market occupancy and percentile data.
