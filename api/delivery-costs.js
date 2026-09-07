import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Delivery-cost automation — processing endpoint.
//
// HARD CONSTRAINT: no uploaded delivery / customer data is ever persisted.
// The uploaded file is decoded, parsed, priced and turned into per-pharmacy
// output files entirely in memory, for the duration of this single request.
// Nothing is written to Supabase, disk, logs, or any cache. The only Supabase
// read here is the (non-customer) pharmacy pricing configuration.
// ---------------------------------------------------------------------------

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SECRET_KEY
);

// The uploaded file arrives base64-encoded in the JSON body. Vercel's default
// serverless request-body limit (~4.5 MB) accommodates the base64 of a typical
// dispatch export; very large exports may need to be split.

// --- Output sheet columns -------------------------------------------------
// The full column set. `Cost` is appended automatically and must not be listed
// here. Change this array to change every output sheet.
const OUTPUT_COLUMNS = [
  "Merge_ID",
  "Task_ID",
  "Order_ID",
  "Task_Type",
  "Agent_ID",
  "Agent_Name",
  "Pick_up_From",       // canonical pharmacy name, sourced from DB by Order_ID
  "Pharmacy_Address",   // canonical pharmacy address, sourced from DB by Order_ID
  "Customer_Name",
  "Customer_Address",
  "Customer_Phone",
  "Complete_Before",
  "Completion_Time",
  "Day",                // computed: day of week from Completion_Time
  "Task_Status",
];
const COST_COLUMN = "Cost";

// Columns whose values come from the matched pharmacy record (DB), not the
// uploaded sheet — this removes name/address discrepancies in the raw export.
const PHARMACY_SOURCED = {
  Pick_up_From: (p) => p.name,
  Pharmacy_Address: (p) => p.address,
};

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

// Parse a dispatch timestamp like "30 Jun 2026 07:08:00 pm" into a Date.
// Returns null for blank/"-"/unparseable values.
export function parseCompletionTime(v) {
  if (v === undefined || v === null) return null;
  const s = String(v).trim();
  if (!s || s === "-") return null;
  const m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?)?/i);
  if (m) {
    const month = MONTHS[m[2].slice(0, 3).toLowerCase()];
    if (month !== undefined) {
      let hour = m[4] ? parseInt(m[4], 10) : 0;
      const ap = m[7] && m[7].toLowerCase();
      if (ap === "pm" && hour < 12) hour += 12;
      if (ap === "am" && hour === 12) hour = 0;
      return new Date(parseInt(m[3], 10), month, parseInt(m[1], 10), hour,
        m[5] ? parseInt(m[5], 10) : 0, m[6] ? parseInt(m[6], 10) : 0);
    }
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Columns computed from the uploaded row (not copied verbatim).
const COMPUTED = {
  Day: (row) => {
    const d = parseCompletionTime(row.Completion_Time);
    return d ? WEEKDAYS[d.getDay()] : "";
  },
};

// Placeholder written to the Cost cell for Delivery rows whose city has no flat
// rate in the pharmacy's city_rates. (Distance-based calculation is disabled
// for now — these are meant to be computed later.)
const NEEDS_CALC = "Need to Calculate";

// Task_Status values that are kept in the output. Both Completed and Failed
// pick-ups/deliveries are included; Failed deliveries are priced with the same
// city flat-rate logic as Completed deliveries. Every other status
// (Cancelled, Unassigned, Assigned, blank, …) is discarded.
const KEPT_STATUSES = new Set(["completed", "failed"]);

function parseNumber(v) {
  if (v === undefined || v === null || String(v).trim() === "" || String(v).trim() === "-") return null;
  const n = Number(String(v).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

export function parseOrderId(v) {
  const n = parseNumber(v);
  return n === null ? null : Math.trunc(n);
}

// --- Pharmacy config validation (manual; Zod may replace this later) ------
function normalizeCityRates(raw) {
  const out = {};
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    for (const [k, v] of Object.entries(raw)) {
      const key = String(k).trim().toLowerCase();
      const rate = typeof v === "string" ? Number(v) : v;
      if (key && typeof rate === "number" && Number.isFinite(rate)) out[key] = rate;
    }
  }
  return out;
}

// --- Pharmacy matching (section 3) ----------------------------------------
export function matchPharmacy(pharmacies, orderId) {
  if (orderId !== null) {
    const match = pharmacies.find((p) => p.order_id === orderId);
    if (match) return { pharmacy: match, source: "order_id" };
  }
  return { pharmacy: null, source: "unmatched" };
}

// --- City resolution ------------------------------------------------------
const PROVINCE_PATTERN =
  /^(Ontario|ON|Quebec|QC|British Columbia|BC|Alberta|AB|Manitoba|MB|Saskatchewan|SK|Nova Scotia|NS|New Brunswick|NB|Newfoundland(?: and Labrador)?|NL|Prince Edward Island|PE|Northwest Territories|NT|Yukon|YT|Nunavut|NU)\b/i;

// Upper-tier regional municipalities / counties that dispatch addresses often
// insert between the city and the province, e.g.
// "1616 Haig Boulevard, Mississauga, Peel, Ontario, Canada". When one of these
// is the token before the province, the real city is one token further back.
const REGION_TOKENS = new Set([
  "peel", "york", "halton", "durham", "niagara", "waterloo",
  "simcoe", "muskoka", "golden horseshoe",
]);

function isRegionToken(t) {
  const s = t.trim().toLowerCase();
  return REGION_TOKENS.has(s) || /\bregion\b/.test(s) || /\bcounty\b/.test(s);
}

export function parseCityFromAddress(address) {
  if (!address) return null;
  const tokens = String(address).split(",").map((t) => t.trim()).filter(Boolean);
  const provinceIndex = tokens.findIndex((t) => PROVINCE_PATTERN.test(t));
  if (provinceIndex <= 0) return null;
  // Walk back from the province, skipping regional-municipality tokens, to reach
  // the actual city.
  let i = provinceIndex - 1;
  while (i >= 0 && isRegionToken(tokens[i])) i--;
  const candidate = tokens[i];
  if (!candidate || /^\d+\s/.test(candidate)) return null; // guards against picking up a street
  return candidate;
}

// --- Cost calculation -----------------------------------------------------
// Distance-based pricing is disabled for now: we only apply the pharmacy's
// city flat rates. A Delivery row whose city has no flat rate gets the
// `Need to Calculate` placeholder (to be priced later), never 0.
// Returns { cost, resolved, city, rate, reason? }.
function calculateCost(row, pharmacy) {
  const city = parseCityFromAddress(row.Customer_Address);
  const key = city ? city.trim().toLowerCase() : null;
  if (key) {
    const rate = pharmacy.city_rates[key];
    if (rate !== undefined)
      return { cost: round2(rate), resolved: true, city: key, rate: round2(rate) };
  }
  return {
    cost: NEEDS_CALC,
    resolved: false,
    city: key,
    rate: null,
    reason: city ? `no flat rate for "${city}"` : "city not resolved from address",
  };
}

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function projectRow(row, cost, pharmacy) {
  const out = {};
  for (const col of OUTPUT_COLUMNS) {
    out[col] = PHARMACY_SOURCED[col] ? PHARMACY_SOURCED[col](pharmacy)
      : COMPUTED[col] ? COMPUTED[col](row)
      : row[col] ?? "";
  }
  out[COST_COLUMN] = cost === null || cost === undefined ? "" : cost;
  return out;
}

// --- Simple admin gate (activates only when ADMIN_ACCESS_KEY is set) -------
function adminAllowed(req) {
  const required = process.env.ADMIN_ACCESS_KEY;
  if (!required) return true; // not configured → non-breaking (flagged in UI)
  const provided = req.headers["x-admin-key"] || req.query.admin_key;
  return provided === required;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }
  if (!adminAllowed(req)) return res.status(401).json({ error: "Unauthorized" });

  try {
    const { fileBase64, format } = req.body || {};
    if (!fileBase64) return res.status(400).json({ error: "No file provided" });
    const fmt = format === "xlsx" ? "xlsx" : "csv";

    // Load pharmacy pricing config (active only). No customer data involved.
    const { data: rawPharmacies, error: dbErr } = await supabase
      .from("pharmacies")
      .select("*")
      .eq("is_active", true);
    if (dbErr) return res.status(500).json({ error: "Failed to load pharmacy config" });

    const pharmacies = (rawPharmacies || []).map((p) => ({
      ...p,
      order_id: Number(p.order_id),
      city_rates: normalizeCityRates(p.city_rates),
    }));

    // Parse the uploaded file in memory.
    const buf = Buffer.from(fileBase64, "base64");
    let rows;
    try {
      const wb = XLSX.read(buf, { type: "buffer" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    } catch {
      return res.status(400).json({ error: "Could not parse file. Expected CSV or XLSX." });
    }

    const { buckets, summary } = processRows(rows, pharmacies);
    const files = buildFiles(buckets, summary, fmt);

    return res.status(200).json({ summary, files });
  } catch (err) {
    console.error("delivery-costs handler error:", err);
    return res.status(500).json({ error: "Internal server error" });
  }
}

// Apply inclusion rules, match by Order_ID, and price Delivery rows via the
// pharmacy's city flat rates. Pure and synchronous. Returns { buckets, summary }.
export function processRows(rows, pharmacies) {
  const buckets = new Map(); // pharmacyId -> { pharmacy, rows, cityStats }
  const summary = {
    totalRows: rows.length,
    kept: 0,
    // Completed and Failed rows are kept; everything else is discarded, tallied by status.
    discarded: { total: 0, byStatus: {} },
    unmatched: { count: 0, orderIds: {} },
    needsCalculation: { count: 0, sample: [] },
    perPharmacy: [],
    warnings: [],
  };

  for (const row of rows) {
    const statusKey = String(row.Task_Status ?? "").trim().toLowerCase();

    // --- Row inclusion: keep Completed and Failed rows; discard everything else. ---
    if (!KEPT_STATUSES.has(statusKey)) {
      summary.discarded.total++;
      const label = statusKey || "(blank)";
      summary.discarded.byStatus[label] = (summary.discarded.byStatus[label] || 0) + 1;
      continue;
    }

    // --- Match to pharmacy by Order_ID only ---
    const orderId = parseOrderId(row.Order_ID);
    const { pharmacy } = matchPharmacy(pharmacies, orderId);
    if (!pharmacy) {
      summary.unmatched.count++;
      const k = orderId === null ? "(blank)" : String(orderId);
      summary.unmatched.orderIds[k] = (summary.unmatched.orderIds[k] || 0) + 1;
      continue;
    }

    summary.kept++;
    if (!buckets.has(pharmacy.id))
      buckets.set(pharmacy.id, { pharmacy, rows: [], cityStats: new Map() });
    const bucket = buckets.get(pharmacy.id);

    // Delivery rows (Completed or Failed) are priced with the city flat-rate
    // logic; Pick-up rows stay blank regardless of status.
    let cost = null;
    const isDelivery = String(row.Task_Type ?? "").trim().toLowerCase() === "delivery";
    if (isDelivery) {
      const result = calculateCost(row, pharmacy);
      cost = result.cost;

      // Per-city delivery stats for the invoice.
      const ck = result.city || "(unresolved)";
      const cs = bucket.cityStats.get(ck) || { city: ck, count: 0, rate: result.rate };
      cs.count++;
      if (cs.rate == null && result.rate != null) cs.rate = result.rate;
      bucket.cityStats.set(ck, cs);

      if (!result.resolved) {
        summary.needsCalculation.count++;
        if (summary.needsCalculation.sample.length < 10) {
          summary.needsCalculation.sample.push({
            pharmacy: pharmacy.name,
            order_id: pharmacy.order_id,
            address: row.Customer_Address,
            reason: result.reason,
          });
        }
      }
    }

    bucket.rows.push(projectRow(row, cost, pharmacy));
  }

  return { buckets, summary };
}

// Label column for the trailing TOTAL row (the data column just before Cost).
const TOTAL_LABEL_COLUMN = OUTPUT_COLUMNS[OUTPUT_COLUMNS.length - 1];

// Build one output file per pharmacy, entirely in memory. Mutates summary
// with the per-pharmacy breakdown + invoice. Returns
// [{ filename, mimeType, base64, … }].
export function buildFiles(buckets, summary, fmt) {
  const files = [];
  for (const { pharmacy, rows: outRows, cityStats } of buckets.values()) {
    // --- Invoice breakdown: group priced deliveries by rate (rate × count) ---
    const groups = new Map(); // rate -> count
    let needsCalc = 0;
    for (const r of outRows) {
      const c = r[COST_COLUMN];
      if (c === NEEDS_CALC) needsCalc++;
      else if (typeof c === "number") groups.set(c, (groups.get(c) || 0) + 1);
    }
    const breakdown = [...groups.entries()]
      .map(([rate, count]) => ({ rate, count, subtotal: round2(rate * count) }))
      .sort((a, b) => a.rate - b.rate);
    const total = round2(breakdown.reduce((s, b) => s + b.subtotal, 0));

    // --- Per-city delivery stats (count, and rate/subtotal when priced) ---
    const cityBreakdown = [...(cityStats ? cityStats.values() : [])]
      .map((s) => ({
        city: s.city,
        count: s.count,
        rate: s.rate ?? null,
        subtotal: s.rate != null ? round2(s.rate * s.count) : null,
      }))
      .sort((a, b) => b.count - a.count);

    const deliveries = outRows.filter(
      (r) => String(r.Task_Type).trim().toLowerCase() === "delivery"
    ).length;

    // --- Append a TOTAL row (sum of Cost, excluding "Need to Calculate") ---
    const totalRow = {};
    for (const col of OUTPUT_COLUMNS) totalRow[col] = "";
    totalRow[TOTAL_LABEL_COLUMN] = "TOTAL";
    totalRow[COST_COLUMN] = total;
    const sheetRows = [...outRows, totalRow];

    const ws = XLSX.utils.json_to_sheet(sheetRows, {
      header: [...OUTPUT_COLUMNS, COST_COLUMN],
    });
    let base64, mimeType;
    if (fmt === "xlsx") {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Deliveries");
      base64 = XLSX.write(wb, { type: "base64", bookType: "xlsx" });
      mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    } else {
      const csv = XLSX.utils.sheet_to_csv(ws);
      base64 = Buffer.from(csv, "utf-8").toString("base64");
      mimeType = "text/csv";
    }

    files.push({
      filename: `${pharmacy.file_name}.${fmt}`,
      mimeType,
      base64,
      rows: outRows.length,
      deliveries,
      pickups: outRows.length - deliveries,
      needsCalc,
    });

    summary.perPharmacy.push({
      name: pharmacy.name,
      order_id: pharmacy.order_id,
      file_name: `${pharmacy.file_name}.${fmt}`,
      rows: outRows.length,
      deliveries,
      pickups: outRows.length - deliveries,
      needsCalc,
      // final = fully priced: no Need to Calculate rows.
      final: needsCalc === 0,
      breakdown,                // [{ rate, count, subtotal }] by rate
      cityBreakdown,            // [{ city, count, rate, subtotal }] by city
      total,                    // sum of priced deliveries (excludes Need to Calculate)
    });
  }

  summary.perPharmacy.sort((a, b) => a.name.localeCompare(b.name));
  files.sort((a, b) => a.filename.localeCompare(b.filename));
  return files;
}
