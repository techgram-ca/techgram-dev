# Delivery-Cost Automation

Upload a dispatch export (CSV/XLSX), calculate a delivery cost per row from
pharmacy-specific pricing rules, and download one output sheet per pharmacy.

## Privacy constraint

Uploaded delivery/customer data is **never persisted** — no database, disk,
logs, or cache. The file is decoded, parsed, priced, and turned into
per-pharmacy output files entirely in memory for the duration of one request
(`api/delivery-costs.js`). The only thing stored in Supabase is pharmacy
pricing config (the `pharmacies` table) — no customer data.

## Pieces

| Path | What |
| --- | --- |
| `supabase/migrations/20260703000000_create_pharmacies.sql` | `pharmacies` pricing-config table (new; touches nothing else) |
| `api/pharmacies.js` | CRUD for pharmacy config (plain-JS Vercel function, Supabase service key) |
| `api/delivery-costs.js` | In-memory upload processing / pricing / per-pharmacy file generation |
| `admin/pharmacies/` | Admin UI: list/create/edit the 12 pharmacies, structured `city_rates` editor |
| `admin/delivery-costs/` | Admin UI: upload, on-screen summary, per-pharmacy downloads |

## Setup

1. **Run the migration** against your Supabase project (only adds `pharmacies`):
   - Supabase CLI: `supabase db push`, or
   - paste `supabase/migrations/20260703000000_create_pharmacies.sql` into the SQL editor.
2. **Install deps**: `npm install` (adds `xlsx` for CSV/XLSX parsing + writing).
3. **Environment variables** (in addition to the existing `SUPABASE_URL` / `SUPABASE_SECRET_KEY`):
   - No Google Maps key is needed right now — distance-based pricing is
     disabled. Pricing uses the pharmacy's city flat rates only.
   - `ADMIN_ACCESS_KEY` — **required to create pharmacies.** The pharmacy form
     has an *Admin passcode* field that must equal this value; `POST
     /api/pharmacies` is rejected with 403 if the key is not configured and 401
     if the passcode does not match. It also gates the other admin routes via
     the `x-admin-key` header (read/edit/delete). When unset, reads/edits stay
     open but insertion is blocked.
4. **Seed the 12 pharmacies** via `/admin/pharmacies`.

## Pricing logic (per row)

1. Only `Task_Type = Delivery` rows are priced (for both `Completed` and
   `Failed` statuses). `Pick-up` rows are kept in the sheet with a blank `Cost`.
2. Resolve the delivery city by parsing `Customer_Address`. Normalized
   (lowercase/trim) lookup in the pharmacy's `city_rates` → flat rate.
3. If the city has no flat rate (city not resolved, or not a key in
   `city_rates`), `Cost` is set to the literal **`Need to Calculate`** — never
   0 — and the row is counted under "Need to Calculate" in the summary.
   Distance-based pricing is disabled for now; the sheet's `Distance(KM)`
   column is not used, and no external APIs are called.

## Row inclusion

- **`Completed` and `Failed` rows are kept.** For both statuses, Delivery rows
  are priced with the same city flat-rate logic and Pick-up rows stay blank —
  so failed pick-ups and failed deliveries appear in the per-pharmacy sheets
  (each row's `Task_Status` column shows whether it was Completed or Failed).
- Every other status (`Cancelled`, `Unassigned`, `Assigned`, blank, …) is
  discarded. The summary reports the discarded total and a per-status
  breakdown (`summary.discarded.byStatus`).

## Assumptions made (change points)

These were defaulted because the spec left them open; each is easy to change:

- **Output columns**: the single `OUTPUT_COLUMNS` array at the top of
  `api/delivery-costs.js` (`Merge_ID, Task_ID, Order_ID, Task_Type, Agent_ID,
  Agent_Name, Pick_up_From, Pharmacy_Address, Customer_Name, Customer_Address,
  Customer_Phone, Complete_Before, Completion_Time, Day, Task_Status` + appended
  `Cost`). `Pick_up_From` (pharmacy name) and `Pharmacy_Address` are sourced
  from the matched pharmacy record in the DB (by `Order_ID`), not the raw sheet
  text — this removes name/address discrepancies. `Day` is computed (weekday
  parsed from `Completion_Time`, e.g. `Tuesday`). Each sheet ends with a
  `TOTAL` row summing `Cost` (excluding `Need to Calculate`). Edit
  `OUTPUT_COLUMNS` / `PHARMACY_SOURCED` / `COMPUTED` to change the layout.
- **Unknown/blank status** (spec §4.3 said to ask): defaulted to *discard +
  surface in summary*. To instead apply the agent check, adjust the `else`
  branch in `processRows`.
- **Admin auth**: this repo had no auth. Added an optional shared-secret gate
  (`ADMIN_ACCESS_KEY`) that is **off unless the env var is set**. Set it (or add
  Vercel deployment protection on `/admin` and `/api/pharmacies`) before
  exposing this publicly — it handles customer PII.

## Output

`POST /api/delivery-costs` with `{ fileBase64, format }` returns
`{ summary, files }`. `files[i]` = `{ filename, mimeType, base64, rows,
deliveries, pickups, needsCalc }`, one per pharmacy, named
`<file_name>.<csv|xlsx>` from the pharmacy config. The summary reports
total/kept/discarded (with reasons), unmatched `Order_ID`s, and
`Need to Calculate` rows — nothing is silently dropped.

`summary.perPharmacy[i]` also carries the invoice data used by the UI:
`breakdown` (`[{ rate, count, subtotal }]` by rate), `cityBreakdown`
(`[{ city, count, rate, subtotal }]` — deliveries per city), `total` (sum of
priced deliveries, excluding `Need to Calculate`), `needsCalc`, and `final`
(`true` when there are no `Need to Calculate` rows). The upload page renders a
green/red status flag per file and an Invoice Summary section (per-pharmacy
`rate × count = subtotal`, deliveries-by-city, `Need to Calculate` count,
pharmacy total, final flag, and a grand total).
