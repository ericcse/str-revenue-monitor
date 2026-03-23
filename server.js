require('dotenv').config();

const path = require('path');
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const qs = require('qs');

const app = express();
const port = process.env.PORT || 3001;

const HOSP_BASE_URL = process.env.HOSPITABLE_BASE_URL || 'https://public.api.hospitable.com/v2';
const HOSP_TOKEN = process.env.HOSPITABLE_ACCESS_TOKEN;
const HOSP_PROPERTY_IDS = (process.env.HOSPITABLE_PROPERTY_IDS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

if (!HOSP_TOKEN) {
  console.warn(
    '[WARN] HOSPITABLE_ACCESS_TOKEN is not set. /api/hospitable/* endpoints will return 500 until you configure .env.'
  );
}

app.use(cors());
app.use(express.json());

// Serve the dashboard from the same origin as the API (avoids file:// + fetch to localhost issues)
app.use(express.static(path.join(__dirname)));

// Simple health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, source: 'str-revenue-monitor-backend' });
});

/** YYYY-MM-DD: today + 1 year (Hospitable rejects queries too far in the future) */
function oneYearFromTodayYMD() {
  const d = new Date();
  d.setFullYear(d.getFullYear() + 1);
  return d.toISOString().slice(0, 10);
}

function minYmd(a, b) {
  return String(a) <= String(b) ? String(a) : String(b);
}

/** JSON:API or flat property resource → Hospitable property UUID */
function extractPropertyIdFromResource(item) {
  if (!item) return null;
  if (typeof item === 'string') return item.trim() || null;
  const id = item.id ?? item.uuid ?? item.property_id;
  if (id != null && String(id).length) return String(id).trim();
  return null;
}

// List properties helper (optional, for you to discover property UUIDs)
app.get('/api/hospitable/properties', async (_req, res) => {
  if (!HOSP_TOKEN) {
    return res.status(500).json({ error: 'HOSPITABLE_ACCESS_TOKEN not configured' });
  }

  try {
    const response = await axios.get(`${HOSP_BASE_URL}/properties`, {
      headers: {
        Authorization: `Bearer ${HOSP_TOKEN}`,
      },
      params: {
        per_page: 100,
        page: 1,
      },
    });

    res.json(response.data);
  } catch (err) {
    console.error('Error fetching properties from Hospitable:', err.response?.data || err.message);
    res.status(500).json({
      error: 'Failed to fetch properties from Hospitable',
      status: err.response?.status,
      details: err.response?.data || err.message,
    });
  }
});

/**
 * Merge JSON:API `attributes` with top-level fields on the resource.
 * Some Hospitable payloads put stay dates only on the resource object; others only under `attributes`.
 * Previously we only spread `attributes`, which dropped top-level `arrival_date` / `departure_date` and mapped 0 rows.
 */
function flattenHospitableReservation(r) {
  if (!r || typeof r !== 'object') return {};
  const attrs =
    r.attributes && typeof r.attributes === 'object' && !Array.isArray(r.attributes)
      ? r.attributes
      : {};
  const { attributes: _omit, relationships, id, type, ...restTop } = r;
  return {
    ...restTop,
    ...attrs,
    id: id ?? attrs.id,
    type: type ?? attrs.type,
    relationships: relationships ?? attrs.relationships,
  };
}

function reservationStayCancelled(row) {
  const cat = (
    row.reservation_status?.current?.category ||
    row.reservation_status?.category ||
    row.status ||
    ''
  )
    .toString()
    .toLowerCase()
    .trim();
  if (!cat) return false;
  // Only drop clearly terminal states (avoid substring false positives on "accepted", etc.)
  if (
    cat === 'cancelled' ||
    cat === 'canceled' ||
    cat === 'not accepted' ||
    cat === 'not_accepted' ||
    cat.startsWith('cancelled_') ||
    cat.startsWith('canceled_')
  ) {
    return true;
  }
  if (cat === 'declined' || cat.startsWith('declined_')) return true;
  return false;
}

/** First usable calendar day as YYYY-MM-DD (prefer date-only; else slice ISO). */
function firstDateYMD(...candidates) {
  for (const v of candidates) {
    if (v == null) continue;
    if (typeof v === 'number' && Number.isFinite(v)) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const mo = String(d.getMonth() + 1).padStart(2, '0');
        const da = String(d.getDate()).padStart(2, '0');
        return `${y}-${mo}-${da}`;
      }
      continue;
    }
    const s = String(v).trim();
    if (!s) continue;
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = /^(\d{4}-\d{2}-\d{2})/.exec(s);
    if (m) return m[1];
  }
  return null;
}

/** Add whole calendar days to YYYY-MM-DD (checkout = arrival + nights for Hospitable-style stays). */
function addDaysToYMD(ymd, daysToAdd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || '').trim());
  if (!m || !Number.isFinite(daysToAdd)) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
  d.setDate(d.getDate() + daysToAdd);
  const y = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  return `${y}-${mo}-${da}`;
}

/** Parse money-like inputs from Hospitable payloads ("$1,234.56", numbers, etc). */
function toMoneyNumber(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/[$€£,\s]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function pickPath(obj, path) {
  let cur = obj;
  for (const key of path) {
    if (!cur || typeof cur !== 'object') return undefined;
    cur = cur[key];
  }
  return cur;
}

/**
 * Hospitable financial nodes often look like:
 * { amount: 87500, formatted: "$875.00", ... } where amount is in minor units (cents).
 */
function toMoneyFromFinancialNode(node) {
  if (node == null) return 0;
  if (typeof node === 'number' || typeof node === 'string') return toMoneyNumber(node);
  if (typeof node !== 'object') return 0;
  if (node.formatted != null) {
    const byFmt = toMoneyNumber(node.formatted);
    if (byFmt) return byFmt;
  }
  if (node.amount != null && Number.isFinite(Number(node.amount))) {
    return Number(node.amount) / 100;
  }
  return 0;
}

// Normalize revenue + stay dates for the frontend.
function mapReservationToMonitorShape(r) {
  const row = flattenHospitableReservation(r);
  if (reservationStayCancelled(row)) return null;

  // Financials may live on resource or under included[] — keep best-effort.
  const financials = row.financials || row.financial || {};

  // Flat fallback fields (legacy / alternative payload shapes)
  const accommodationFlat = toMoneyNumber(
    financials.accommodationAmount ??
      financials.accommodation ??
      financials.accommodation_amount ??
      row.accommodationAmount ??
      row.accommodation ??
      row.accommodation_amount
  );
  const hostPayoutFlat = toMoneyNumber(
    financials.hostPayoutAmount ??
      financials.hostAmount ??
      financials.host_payout ??
      financials.host_amount ??
      row.hostPayoutAmount ??
      row.hostAmount ??
      row.host_payout ??
      row.host_amount
  );
  const totalPayoutFlat = toMoneyNumber(
    financials.totalPayoutAmount ??
      financials.totalAmount ??
      financials.total ??
      financials.total_amount ??
      row.totalPayoutAmount ??
      row.totalAmount ??
      row.total ??
      row.total_amount
  );

  // Nested Hospitable v2 shape (financials.guest.*, financials.host.*)
  const accommodationNested =
    toMoneyFromFinancialNode(pickPath(financials, ['host', 'accommodation'])) ||
    toMoneyFromFinancialNode(pickPath(financials, ['guest', 'accommodation']));
  const hostPayoutNested = toMoneyFromFinancialNode(
    pickPath(financials, ['host', 'total_payout'])
  );
  const totalPayoutNested = toMoneyFromFinancialNode(
    pickPath(financials, ['guest', 'total_price'])
  );

  const accommodation = accommodationFlat || accommodationNested || 0;
  const hostPayout = hostPayoutFlat || hostPayoutNested || 0;
  const totalPayout = totalPayoutFlat || totalPayoutNested || 0;

  // Hospitable: arrival_date / departure_date are the canonical stay nights (checkout day exclusive).
  // Do NOT prefer check_in before arrival_date — timestamps can extend the computed stay incorrectly.
  const checkIn = firstDateYMD(
    row.arrival_date,
    row.arrivalDate,
    row.arrival,
    row.check_in,
    row.checkIn,
    row.checkInDate,
    row.start_date,
    row.startDate,
    row.starts_at,
    row.startsAt
  );
  let checkOut = firstDateYMD(
    row.departure_date,
    row.departureDate,
    row.departure,
    row.check_out,
    row.checkOut,
    row.checkOutDate,
    row.end_date,
    row.endDate,
    row.ends_at,
    row.endsAt
  );

  const nightsNum = Number(row.nights);
  if (checkIn && (!checkOut || checkOut <= checkIn) && Number.isFinite(nightsNum) && nightsNum > 0) {
    checkOut = addDaysToYMD(checkIn, nightsNum);
  }

  if (!checkIn || !checkOut || checkOut <= checkIn) return null;

  return {
    checkIn,
    checkOut,
    revenue: accommodation || hostPayout || totalPayout || 0,
    accommodation,
    hostPayout,
    totalPayout,
    hospitableId: row.id || r.id,
    raw: r,
  };
}

/** Night `nightYmd` is booked iff checkIn <= night < checkOut (YYYY-MM-DD lexicographic ok). */
function reservationCoversNight(checkIn, checkOut, nightYmd) {
  return checkIn && checkOut && nightYmd >= checkIn && nightYmd < checkOut;
}

async function fetchReservationsForDateQuery(propertyIds, startDate, endDate, dateQuery) {
  const perPage = 100;
  const maxPages = 500;
  const batch = [];
  let page = 1;
  let hasMore = true;

  while (hasMore && page <= maxPages) {
    const response = await axios.get(`${HOSP_BASE_URL}/reservations`, {
      headers: { Authorization: `Bearer ${HOSP_TOKEN}` },
      params: {
        per_page: perPage,
        page,
        start_date: startDate,
        end_date: endDate,
        include: 'financials',
        date_query: dateQuery,
        properties: propertyIds,
      },
      paramsSerializer: (p) => qs.stringify(p, { arrayFormat: 'brackets', encodeValuesOnly: true }),
    });

    const data = Array.isArray(response.data?.data)
      ? response.data.data
      : Array.isArray(response.data)
      ? response.data
      : [];

    batch.push(...data);

    const meta = response.data?.meta || {};
    const totalPages =
      meta.last_page ||
      meta.lastPage ||
      meta.totalPages ||
      meta.total_pages ||
      (data.length < perPage ? page : page + 1);

    page += 1;
    hasMore = page <= totalPages && data.length > 0;
  }

  return batch;
}

/** Same reservation can match check-in filter or check-out filter; merge so long stays aren’t dropped. */
function mergeReservationRecordsById(lists) {
  const map = new Map();
  for (const list of lists) {
    for (const item of list) {
      const flat = flattenHospitableReservation(item);
      const id = flat.id != null && String(flat.id).trim() !== '' ? String(flat.id) : null;
      const key = id || `__anon_${map.size}`;
      map.set(key, item);
    }
  }
  return [...map.values()];
}

/**
 * Paginate GET /properties and collect unique property UUIDs (matches Hospitable v2 client).
 */
async function fetchAllPropertyIdsFromApi() {
  const all = [];
  let page = 1;
  const perPage = 100;

  for (;;) {
    const response = await axios.get(`${HOSP_BASE_URL}/properties`, {
      headers: { Authorization: `Bearer ${HOSP_TOKEN}` },
      params: { per_page: perPage, page },
    });

    const body = response.data || {};
    const items = Array.isArray(body.data) ? body.data : [];
    if (!items.length) break;

    for (const item of items) {
      const id = extractPropertyIdFromResource(item);
      if (id) all.push(id);
    }

    const meta = body.meta || {};
    const lastPage = meta.last_page ?? meta.lastPage;
    if (lastPage != null && page >= lastPage) break;
    if (items.length < perPage) break;

    page += 1;
    if (page > 100) break;
  }

  return [...new Set(all)];
}

/**
 * Resolves property UUIDs: optional query/env list is validated against GET /properties.
 * Wrong IDs in .env (e.g. placeholders) are rejected with a clear error.
 */
async function resolvePropertyIds(req) {
  const fromQuery =
    req.query.properties &&
    String(req.query.properties)
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);

  const requested = (fromQuery && fromQuery.length ? fromQuery : HOSP_PROPERTY_IDS) || [];

  const validFromApi = await fetchAllPropertyIdsFromApi();
  const validSet = new Set(validFromApi);

  if (!validFromApi.length) {
    return { ids: [], validFromApi: [], error: 'Could not load any properties from Hospitable (empty list or parse error).' };
  }

  if (requested.length) {
    const filtered = requested.filter((id) => validSet.has(id));
    if (!filtered.length) {
      return {
        ids: [],
        validFromApi,
        error:
          'None of the configured property IDs exist on your Hospitable account. ' +
          'Use each `id` from GET /api/hospitable/properties → `data` (or clear HOSPITABLE_PROPERTY_IDS to use all properties).',
      };
    }
    return { ids: filtered, validFromApi };
  }

  return { ids: validFromApi, validFromApi };
}

// Fetch reservations from Hospitable and map into the app’s shape.
app.get('/api/hospitable/reservations', async (req, res) => {
  if (!HOSP_TOKEN) {
    return res.status(500).json({ error: 'HOSPITABLE_ACCESS_TOKEN not configured' });
  }

  const startDate = req.query.startDate || req.query.start_date || '1970-01-01';
  const maxEnd = oneYearFromTodayYMD();
  let endDate = req.query.endDate || req.query.end_date || maxEnd;
  endDate = minYmd(endDate, maxEnd);

  try {
    const resolved = await resolvePropertyIds(req);
    if (resolved.error) {
      return res.status(400).json({
        error: resolved.error,
        validPropertyIdsSample: (resolved.validFromApi || []).slice(0, 10),
      });
    }
    const propertyIds = resolved.ids;
    if (!propertyIds.length) {
      return res.status(400).json({
        error: 'No property IDs available',
        hint:
          'Set HOSPITABLE_PROPERTY_IDS in .env (comma-separated UUIDs from GET /api/hospitable/properties), or leave it empty to use all properties.',
      });
    }

    let checkinList = [];
    let checkoutList = [];
    let checkoutQueryError = null;
    try {
      checkinList = await fetchReservationsForDateQuery(propertyIds, startDate, endDate, 'checkin');
    } catch (e) {
      console.error('Hospitable checkin reservations query failed:', e.response?.data || e.message);
      throw e;
    }
    try {
      checkoutList = await fetchReservationsForDateQuery(propertyIds, startDate, endDate, 'checkout');
    } catch (e) {
      checkoutQueryError = e.response?.data || e.message || String(e);
      console.warn('Hospitable checkout reservations query failed (using check-in results only):', checkoutQueryError);
    }

    const all = mergeReservationRecordsById([checkinList, checkoutList]);

    const mapped = all.map(mapReservationToMonitorShape).filter((r) => r && r.checkIn && r.checkOut);

    const skippedMapping = all.length - mapped.length;

    const explainDate = req.query.explainDate || req.query.explain_date;
    let explain;
    if (explainDate && /^\d{4}-\d{2}-\d{2}$/.test(String(explainDate).trim())) {
      const night = String(explainDate).trim().slice(0, 10);
      explain = {
        night,
        reservationsMarkingNightBooked: mapped
          .filter((r) => reservationCoversNight(r.checkIn, r.checkOut, night))
          .map((r) => {
            const flat = flattenHospitableReservation(r.raw);
            return {
              hospitableId: r.hospitableId,
              checkIn: r.checkIn,
              checkOut: r.checkOut,
              arrival_date: flat.arrival_date,
              departure_date: flat.departure_date,
              statusCategory: flat.reservation_status?.current?.category || flat.status || null,
            };
          }),
      };
    }

    res.json({
      startDate,
      endDate,
      count: mapped.length,
      reservations: mapped.map((r) => ({
        checkIn: r.checkIn,
        checkOut: r.checkOut,
        revenue: r.revenue,
      })),
      diagnostics: {
        rawReservationObjects: all.length,
        returnedReservations: mapped.length,
        droppedByMapping: skippedMapping,
        checkinQueryRows: checkinList.length,
        checkoutQueryRows: checkoutList.length,
        checkoutQueryFailed: Boolean(checkoutQueryError),
        checkoutQueryError: checkoutQueryError || undefined,
        sampleMergedKeys:
          all.length > 0
            ? Object.keys(flattenHospitableReservation(all[0]))
                .filter((k) => k.length < 80)
                .slice(0, 60)
            : [],
      },
      ...(explain ? { explain } : {}),
    });
  } catch (err) {
    console.error('Error fetching reservations from Hospitable:', err.response?.data || err.message);
    res.status(500).json({
      error: 'Failed to fetch reservations from Hospitable',
      status: err.response?.status,
      details: err.response?.data || err.message,
      hint:
        !HOSP_TOKEN
          ? 'Create a .env file with HOSPITABLE_ACCESS_TOKEN (copy from .env.example).'
          : 'Check token scopes (reservations read) and that the API base URL is correct.',
    });
  }
});

app.listen(port, () => {
  console.log(`STR Revenue Monitor backend listening on http://localhost:${port}`);
});

