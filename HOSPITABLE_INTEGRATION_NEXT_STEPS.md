# Hospitable Integration — Next Steps

## Goal
Add an option to pull **real-time reservations** from **Hospitable** instead of only using **“Use sample data”**.

## Key Constraint
Hospitable API authentication uses a **Personal Access Token (PAT)**, so the browser should **not** call Hospitable directly. Use a **backend proxy**.

## Recommended Implementation (safe + simple)
1. Add a small backend proxy (Node/Express or serverless).
2. The browser calls the backend endpoint (no token in the browser).
3. The backend calls Hospitable using:
   - `PAT` stored in environment variables
   - optional `property UUID(s)` filtering
4. Backend returns reservations in the same format the app already expects:
   - `[{ checkIn: 'YYYY-MM-DD', checkOut: 'YYYY-MM-DD', revenue: number }, ...]`
5. Client updates pacing by reusing existing render logic.

## UI / UX Change
Add a new button/option next to **“Use sample data”**:
- `Fetch from Hospitable`

## Backend Contract (proposed)
`GET /api/hospitable/reservations?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD`
→ returns JSON array of reservation objects:
`[{ checkIn, checkOut, revenue }]`

`endDate` is **capped at today + 1 year** (Hospitable does not allow pricing/reservations queries too far in the future).

## Questions to Confirm Before Implementing
1. Should this run **locally only**, or do you want it **hosted**?
2. Do you already have a **Hospitable PAT** and the **property UUID(s)**?
3. Which revenue field should be used?
   - host payout vs total payout vs another Hospitable field
4. What date range(s) should fetch cover?
   - only the **next 90 days** (for “future unbooked periods”)
   - and/or also include reservations needed for **current month/quarter/year pacing**

## How to run locally (recommended)

1. Copy `.env.example` → `.env` and set `HOSPITABLE_ACCESS_TOKEN` (and optionally `HOSPITABLE_PROPERTY_IDS`).
2. From the project folder: `npm install` then `npm start`.
3. **Open the app in the browser at:** `http://localhost:3001/`  
   (Same origin as the API — avoids `file://` pages failing to fetch `http://localhost:3001`.)
4. Click **Fetch from Hospitable**.

## If fetch fails

- **“HOSPITABLE_ACCESS_TOKEN not configured”** — create `.env` in the project root (not only `.env.example`).
- **401 / invalid token** — regenerate the PAT in Hospitable (**Apps → API access**) and ensure **read** access for reservations.
- **Opening `index.html` by double-click** — use `http://localhost:3001/` instead so the page and API share the same origin.

