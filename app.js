/**
 * STR Revenue Monitor — Pacing & revenue management
 * Works with CSV upload or sample data. Ready for Hospitable/PriceLabs API later.
 */

const STORAGE_KEYS = {
  revenueTarget: 'str-monitor-revenue-target',
  period: 'str-monitor-period',
  reservations: 'str-monitor-reservations',
  pacingTiers: 'str-monitor-pacing-tiers',
  unbookedHorizonDays: 'str-monitor-unbooked-horizon-days',
  marketPacing: 'str-monitor-market-pacing',
};

/** Allowed unbooked lookahead windows (≈ calendar months). Default 4 months. */
const UNBOOKED_HORIZON_OPTIONS = [30, 120, 180, 270, 365];
const DEFAULT_UNBOOKED_HORIZON_DAYS = 120;

/**
 * How far ahead we scan when building Fri–Thu unbooked rows for the Future table.
 * Must cover max dropdown (12 months) so a period that starts inside the selected window
 * but ends after the cutoff still shows as one full row (not truncated at lastYmd).
 */
const UNBOOKED_PERIOD_TABLE_LOOKAHEAD_DAYS = 370;

let futurePacingChartInstance = null;
let futurePacingChartResizeObserver = null;
let chartDataLabelsPluginRegistered = false;

function ensureChartDataLabelsRegistered() {
  if (chartDataLabelsPluginRegistered || typeof Chart === 'undefined') return;
  const plug = typeof ChartDataLabels !== 'undefined' ? ChartDataLabels : globalThis.ChartDataLabels;
  if (!plug) return;
  try {
    Chart.register(plug);
    chartDataLabelsPluginRegistered = true;
  } catch {
    chartDataLabelsPluginRegistered = true;
  }
}

// Local backend for Hospitable proxy (run with: npm start).
// When you open the app at http://localhost:3001/ use same-origin (empty string).
// When opening index.html as file://, use full URL (browsers often block fetch to localhost otherwise).
function getApiBase() {
  if (typeof window === 'undefined') return 'http://localhost:3001';
  const { protocol, hostname, port } = window.location;
  if (protocol === 'file:' || protocol === 'null') return 'http://localhost:3001';
  if (hostname === 'localhost' && String(port) === '3001') return '';
  return 'http://localhost:3001';
}

// Default: final occupancy ≥75% → 75-100 pct, 50-75% → 50-75 pct, <50% → 25-50 pct
const DEFAULT_PACING_TIERS = [
  { id: 'high', label: 'High demand', minFinalOccupancy: 75, maxFinalOccupancy: 100, minPercentile: 75, maxPercentile: 100 },
  { id: 'mid', label: 'Mid demand', minFinalOccupancy: 50, maxFinalOccupancy: 75, minPercentile: 50, maxPercentile: 75 },
  { id: 'low', label: 'Low demand', minFinalOccupancy: 0, maxFinalOccupancy: 50, minPercentile: 25, maxPercentile: 50 },
];

// --- Period helpers ---
function getPeriodBounds(periodKey) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);

  if (periodKey === 'month') {
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
    end.setMonth(end.getMonth() + 1);
    end.setDate(0);
    end.setHours(23, 59, 59, 999);
  } else if (periodKey === 'quarter') {
    const q = Math.floor(now.getMonth() / 3) + 1;
    start.setMonth((q - 1) * 3, 1);
    start.setHours(0, 0, 0, 0);
    end.setMonth(q * 3, 0);
    end.setHours(23, 59, 59, 999);
  } else {
    start.setMonth(0, 1);
    start.setHours(0, 0, 0, 0);
    end.setMonth(11, 31);
    end.setHours(23, 59, 59, 999);
  }
  return { start, end };
}

function daysInPeriod(start, end) {
  return Math.round((end - start) / (24 * 60 * 60 * 1000)) + 1;
}

function isInPeriod(date, start, end) {
  const d = new Date(date);
  d.setHours(12, 0, 0, 0);
  const s = new Date(start);
  s.setHours(0, 0, 0, 0);
  const e = new Date(end);
  e.setHours(23, 59, 59, 999);
  return d >= s && d <= e;
}

// --- Reservations: array of { checkIn, checkOut, revenue } ---
function parseReservations(data) {
  if (Array.isArray(data)) return data;
  const raw = typeof data === 'string' ? data : JSON.stringify(data);
  try {
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

function nightsInPeriod(reservation, periodStart, periodEnd) {
  const cin = new Date(reservation.checkIn);
  const cout = new Date(reservation.checkOut);
  const start = new Date(Math.max(cin.getTime(), periodStart.getTime()));
  const end = new Date(Math.min(cout.getTime(), periodEnd.getTime()));
  if (start >= end) return 0;
  return Math.round((end - start) / (24 * 60 * 60 * 1000));
}

function revenueInPeriod(reservation, periodStart, periodEnd) {
  const nights = nightsInPeriod(reservation, periodStart, periodEnd);
  const totalNights = Math.round(
    (new Date(reservation.checkOut) - new Date(reservation.checkIn)) / (24 * 60 * 60 * 1000)
  );
  if (totalNights <= 0) return 0;
  const rev = Number(reservation.revenue) || 0;
  return (rev / totalNights) * nights;
}

// --- Metrics ---
function computePeriodMetrics(reservations, periodStart, periodEnd) {
  const totalDays = daysInPeriod(periodStart, periodEnd);
  let revenue = 0;
  let nights = 0;
  for (const r of reservations) {
    revenue += revenueInPeriod(r, periodStart, periodEnd);
    nights += nightsInPeriod(r, periodStart, periodEnd);
  }
  const occupancy = totalDays > 0 ? (nights / totalDays) * 100 : 0;
  const adr = nights > 0 ? revenue / nights : 0;
  const revpar = totalDays > 0 ? revenue / totalDays : 0;
  return { revenue, nights, totalDays, occupancy, adr, revpar };
}

// --- Pacing ---
function computePacing(reservations, periodKey, targetRevenue) {
  const { start, end } = getPeriodBounds(periodKey);
  const now = new Date();
  const elapsed = Math.max(
    0,
    Math.round((Math.min(now, end) - start) / (24 * 60 * 60 * 1000)) + 1
  );
  const totalDays = daysInPeriod(start, end);
  const expectedProgress = totalDays > 0 ? elapsed / totalDays : 0;

  const { revenue, nights, totalDays: periodDays } = computePeriodMetrics(
    reservations,
    start,
    end
  );
  const revenuePace = targetRevenue > 0 ? (revenue / targetRevenue) * 100 : null;
  const occupancyPace = periodDays > 0 ? (nights / periodDays) * 100 : 0;
  const targetOccupancyPct = 70; // optional: could be user setting
  const occupancyVsTarget =
    targetOccupancyPct > 0 ? (occupancyPace / targetOccupancyPct) * 100 : null;

  let status = '—';
  let statusClass = '';
  let action = 'Set a revenue target to see pacing.';

  if (targetRevenue != null && targetRevenue > 0 && revenuePace != null) {
    const onPace = revenuePace >= expectedProgress * 100 * 0.95;
    const behind = revenuePace < expectedProgress * 100 * 0.8;
    if (onPace) {
      status = 'On pace';
      statusClass = 'on-track';
      action = `Revenue at ${revenuePace.toFixed(0)}% of target with ${(expectedProgress * 100).toFixed(0)}% of period elapsed.`;
    } else if (behind) {
      status = 'Behind pace';
      statusClass = 'at-risk';
      action = `Revenue at ${revenuePace.toFixed(0)}% of target. Consider adjusting PriceLabs or minimum stays.`;
    } else {
      status = 'Slightly behind';
      statusClass = 'behind';
      action = `Revenue at ${revenuePace.toFixed(0)}% of target.`;
    }
  }

  return {
    revenue,
    revenuePace,
    occupancyPace,
    occupancyVsTarget,
    elapsed,
    totalDays,
    expectedProgress: expectedProgress * 100,
    status,
    statusClass,
    action,
  };
}

/** Calendar YYYY-MM-DD in the browser's local timezone (avoids UTC off-by-one from toISOString). */
function toYMDLocal(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** Parse YYYY-MM-DD as a local calendar date (noon) so night iteration is stable across DST. */
function parseYMDLocal(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || '').trim());
  if (!m) return new Date(ymd);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
}

/**
 * First and last calendar dates (YYYY-MM-DD) included in the unbooked window.
 * Exactly `daysAhead` nights: from "today" through today + (daysAhead - 1).
 */
function getUnbookedWindowYMDBounds(fromDate, daysAhead) {
  if (!Number.isFinite(daysAhead) || daysAhead < 1) return { firstYmd: null, lastYmd: null };
  const start = new Date(fromDate);
  start.setHours(12, 0, 0, 0);
  const firstYmd = toYMDLocal(start);
  const last = new Date(start);
  last.setDate(last.getDate() + daysAhead - 1);
  const lastYmd = toYMDLocal(last);
  return { firstYmd, lastYmd };
}

// --- Unbooked dates: exactly `daysAhead` calendar nights in [firstYmd, lastYmd] (inclusive) ---
function getUnbookedDates(reservations, fromDate, daysAhead = DEFAULT_UNBOOKED_HORIZON_DAYS) {
  if (!Number.isFinite(daysAhead) || daysAhead < 1) return [];

  const start = new Date(fromDate);
  start.setHours(12, 0, 0, 0);
  const endExclusive = new Date(start);
  endExclusive.setDate(endExclusive.getDate() + daysAhead);

  const booked = new Set();
  for (const r of reservations) {
    if (!r.checkIn || !r.checkOut) continue;
    const cin = parseYMDLocal(r.checkIn);
    const cout = parseYMDLocal(r.checkOut);
    if (isNaN(cin.getTime()) || isNaN(cout.getTime()) || cout <= cin) continue;
    for (let d = new Date(cin); d < cout; d.setDate(d.getDate() + 1)) {
      booked.add(toYMDLocal(d));
    }
  }

  const out = [];
  const cur = new Date(start);
  while (cur < endExclusive) {
    const key = toYMDLocal(cur);
    if (!booked.has(key)) out.push(key);
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

/** True if a Fri–Thu period intersects the user's selected unbooked window [firstYmd, lastYmd]. */
function periodOverlapsUnbookedWindow(p, firstYmd, lastYmd) {
  if (!p || !p.periodStart || !p.periodEnd || !firstYmd || !lastYmd) return false;
  return p.periodStart <= lastYmd && p.periodEnd >= firstYmd;
}

// --- Unbooked periods by consecutive nights (for market pacing) ---
function getUnbookedPeriods(reservations, fromDate, daysAhead = DEFAULT_UNBOOKED_HORIZON_DAYS) {
  const unbooked = getUnbookedDates(reservations, fromDate, daysAhead);
  if (unbooked.length === 0) return [];

  const sorted = [...unbooked].sort(); // YYYY-MM-DD ISO strings sort lexicographically
  const out = [];

  let periodStart = sorted[0];
  let periodEnd = sorted[0];
  let nights = 1;

  for (let i = 1; i < sorted.length; i++) {
    const prev = new Date(periodEnd + 'T12:00:00');
    const cur = new Date(sorted[i] + 'T12:00:00');
    const diffDays = Math.round((cur.getTime() - prev.getTime()) / (24 * 60 * 60 * 1000));

    if (diffDays === 1) {
      // Extend the current consecutive streak.
      periodEnd = sorted[i];
      nights++;
      continue;
    }

    // Split each consecutive streak into Fri-Thu chunks so periods never cross.
    out.push(...splitPeriodIntoFriThuChunks(periodStart, periodEnd));

    periodStart = sorted[i];
    periodEnd = sorted[i];
    nights = 1;
  }

  out.push(...splitPeriodIntoFriThuChunks(periodStart, periodEnd));
  return out;
}

// --- Pacing tier: given final expected occupancy, return which tier (percentile range) ---
function getPricingTier(finalExpectedOccupancy, tiers) {
  const pct = Number(finalExpectedOccupancy);
  if (isNaN(pct)) return null;
  for (let i = 0; i < tiers.length; i++) {
    const t = tiers[i];
    const min = Number(t.minFinalOccupancy);
    const max = Number(t.maxFinalOccupancy);
    const inRange = max >= 100 ? (pct >= min && pct <= 100) : (pct >= min && pct < max);
    if (inRange) return { ...t, tierClass: t.id };
  }
  return tiers.length ? { ...tiers[tiers.length - 1], tierClass: tiers[tiers.length - 1].id } : null;
}

/**
 * Implied percentile (0–100) of Final Price vs market ladder P25–P90: piecewise linear between
 * (P25,25), (P50,50), (P75,75), (P90,90); below P25 / above P90 uses the same slopes as the adjacent segment.
 */
function deriveFinalPricePercentile(finalPrice, p25, p50, p75, p90) {
  const f = Number(finalPrice);
  const q25 = Number(p25);
  const q50 = Number(p50);
  const q75 = Number(p75);
  const q90 = Number(p90);
  if (!Number.isFinite(f) || ![q25, q50, q75, q90].every(Number.isFinite) || q25 <= 0) return null;
  if (q25 > q50 || q50 > q75 || q75 > q90) return null;
  const eps = 1e-6;
  let pct;
  if (f <= q50) {
    const d = q50 - q25;
    if (d < eps) return Math.round(Math.max(0, Math.min(100, (25 * f) / q25)));
    pct = 25 + ((f - q25) / d) * 25;
  } else if (f <= q75) {
    const d = q75 - q50;
    if (d < eps) return 50;
    pct = 50 + ((f - q50) / d) * 25;
  } else if (f <= q90) {
    const d = q90 - q75;
    if (d < eps) return 75;
    pct = 75 + ((f - q75) / d) * 15;
  } else {
    const d = q90 - q75;
    if (d < eps) pct = 90;
    else pct = 90 + ((f - q90) / d) * 15;
  }
  return Math.round(Math.max(0, Math.min(100, pct)));
}

// --- Sample market occupancy by week: current, LY same-date pace (LYT), LY final realized — Fri–Thu weeks. ---
function getSampleMarketOccupancy(fromDate, weeksAhead = 14) {
  const d = new Date(fromDate);
  d.setHours(0, 0, 0, 0);
  const day = d.getDay();
  const friday = new Date(d);
  // Fri-Thu weeks: start on Friday (Fri=5), end on Thursday.
  // day 0(Sun) => go back 2 days to Friday, day 4(Thu) => go back 6 days to Friday.
  const daysSinceFriday = (day - 5 + 7) % 7;
  friday.setDate(d.getDate() - daysSinceFriday);
  const out = [];
  const denom = Math.max(1, weeksAhead - 1);
  for (let w = 0; w < weeksAhead; w++) {
    const weekStart = new Date(friday);
    weekStart.setDate(friday.getDate() + w * 7);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6); // Thursday
    const key = formatLocalNoonToYMD(weekStart); // Friday key (local calendar, not UTC)
    const lastYearSameWeek = (52 + w * 2 + (weekStart.getMonth() % 3) * 5) % 100;
    const finalExpected = Math.min(92, Math.max(28, lastYearSameWeek + 35));
    const progress = w / denom; // 0 = nearest Fri–Thu block, 1 = farthest in window
    // LYT: market booking level for this stay week on the same calendar date last year (pace snapshot).
    const lytFactor = 0.18 + 0.72 * Math.pow(progress, 1.12);
    const lastYearTodayOccupancy = Math.round(
      Math.min(finalExpected - 1, Math.max(5, finalExpected * lytFactor))
    );
    const currentRaw = finalExpected * (0.17 + progress * 0.62 + ((weekStart.getMonth() % 5) * 0.015));
    const currentMarketOccupancy = Math.round(
      Math.min(finalExpected * 0.96, Math.max(3, currentRaw))
    );
    const priceAnchor = 178 + ((w * 13 + weekStart.getMonth() * 7 + lastYearSameWeek) % 88);
    const finalPrice = Math.round(priceAnchor * 1.08);
    const priceP25 = Math.round(priceAnchor * 0.8);
    const priceP50 = Math.round(priceAnchor * 0.93);
    const priceP75 = Math.round(priceAnchor * 1.05);
    const priceP90 = Math.round(priceAnchor * 1.17);
    out.push({
      periodStart: key,
      periodEnd: formatLocalNoonToYMD(weekEnd),
      currentMarketOccupancy,
      lastYearTodayOccupancy,
      finalExpectedOccupancy: Math.round(finalExpected),
      finalPrice,
      priceP25,
      priceP50,
      priceP75,
      priceP90,
    });
  }
  return out;
}

function getMarketForPeriod(marketData, periodStart) {
  if (!marketData || !marketData.length) return null;
  const key = getFriThuWeekStartKey(periodStart);
  return marketData.find((m) => m.periodStart === key) || null;
}

function parseYMDToLocalNoon(dateStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr).trim());
  if (!m) return null;
  const yyyy = Number(m[1]);
  const mm = Number(m[2]) - 1;
  const dd = Number(m[3]);
  return new Date(yyyy, mm, dd, 12, 0, 0, 0);
}

function formatLocalNoonToYMD(d) {
  if (!d || isNaN(d.getTime())) return '';
  return toYMDLocal(d);
}

function getFriThuWeekStartKey(dateStr) {
  // Fri-Thu week: start on Friday (Fri=5) and end on Thursday.
  const d = parseYMDToLocalNoon(dateStr);
  if (!d) return dateStr;
  const day = d.getDay();
  const daysSinceFriday = (day - 5 + 7) % 7;
  d.setDate(d.getDate() - daysSinceFriday);
  d.setHours(12, 0, 0, 0);
  return formatLocalNoonToYMD(d);
}

function addDaysYMD(dateStr, days) {
  const d = parseYMDToLocalNoon(dateStr);
  if (!d) return dateStr;
  d.setDate(d.getDate() + days);
  return formatLocalNoonToYMD(d);
}

function diffDaysInclusive(fromDateStr, toDateStr) {
  const a = parseYMDToLocalNoon(fromDateStr);
  const b = parseYMDToLocalNoon(toDateStr);
  if (!a || !b) return 0;
  const ms = b.getTime() - a.getTime();
  return Math.round(ms / (24 * 60 * 60 * 1000)) + 1;
}

function splitPeriodIntoFriThuChunks(periodStart, periodEnd) {
  // Split a consecutive unbooked streak into Fri-Thu sub-periods.
  // - If it starts before Friday, first chunk ends on Thursday and next starts Friday.
  // - If it ends after Thursday, last chunk ends on Thursday and next starts Friday.
  // Each returned chunk will always fall within a single Fri-Thu week.
  if (!periodStart || !periodEnd || periodEnd < periodStart) return [];

  const chunks = [];
  let cur = periodStart;

  // Loop forward until we cover the full streak.
  while (cur <= periodEnd) {
    const weekStartKey = getFriThuWeekStartKey(cur); // Friday key for this cur date
    const weekStart = parseYMDToLocalNoon(weekStartKey);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6); // Thursday
    const weekEndKey = formatLocalNoonToYMD(weekEnd);

    const segEnd = weekEndKey < periodEnd ? weekEndKey : periodEnd;
    chunks.push({
      periodStart: cur,
      periodEnd: segEnd,
      unbookedNights: diffDaysInclusive(cur, segEnd),
    });

    cur = addDaysYMD(segEnd, 1);
  }

  return chunks;
}

// --- Recommendations (includes pricing tier suggestions from market pacing) ---
function getRecommendations(
  reservations,
  periodKey,
  targetRevenue,
  unbooked,
  pacingTiers,
  futurePeriodsWithMarket,
  unbookedHorizonDays
) {
  const recs = [];
  const { start, end } = getPeriodBounds(periodKey);
  const { revenue, occupancy } = computePeriodMetrics(reservations, start, end);
  const pacing = computePacing(reservations, periodKey, targetRevenue);

  if (targetRevenue > 0 && pacing.revenuePace != null && pacing.revenuePace < pacing.expectedProgress * 0.9) {
    recs.push({
      type: 'warning',
      text: `Revenue is behind pace (${pacing.revenuePace.toFixed(0)}% of target). Review PriceLabs rules or minimum stay for key dates.`,
    });
  }
  if (futurePeriodsWithMarket && futurePeriodsWithMarket.length > 0 && pacingTiers && pacingTiers.length > 0) {
    futurePeriodsWithMarket.slice(0, 5).forEach((fp) => {
      const tier = fp.tier;
      if (!tier) return;
      recs.push({
        type: 'action',
        text: `Week of ${fp.periodStart}: Market final occupancy ${fp.finalExpected}% → price at ${tier.minPercentile}-${tier.maxPercentile} percentile (${tier.label}). ${fp.remainingPotential > 30 ? 'Good remaining demand—consider holding rate.' : 'Limited remaining demand—consider promotions.'}`,
      });
    });
  }
  if (unbooked.length > 14) {
    recs.push({
      type: 'action',
      text: `${unbooked.length} unbooked nights in the selected window (${unbookedHorizonDays} days). Use the Future unbooked periods table to set price percentiles by week.`,
    });
  }
  const todayNoon = new Date();
  todayNoon.setHours(12, 0, 0, 0);
  const nextTwoWeeks = unbooked.filter((d) => {
    const day = parseYMDLocal(d);
    const diff = (day.getTime() - todayNoon.getTime()) / (24 * 60 * 60 * 1000);
    return diff >= -1 && diff <= 14;
  });
  if (nextTwoWeeks.length > 5) {
    recs.push({
      type: 'action',
      text: `${nextTwoWeeks.length} nights unbooked in the next 14 days. Consider a last-minute discount or minimum stay of 1.`,
    });
  }
  if (occupancy < 50 && reservations.length > 0) {
    recs.push({
      type: 'action',
      text: `Occupancy this period is ${occupancy.toFixed(0)}%. Review ADR and length-of-stay settings in PriceLabs.`,
    });
  }
  if (recs.length === 0 && reservations.length > 0) {
    recs.push({ type: 'ok', text: 'Pacing and occupancy look reasonable. Keep monitoring unbooked windows and market pacing tiers.' });
  }
  return recs;
}

// --- Sample data ---
function getSampleReservations() {
  // Generate non-overlapping booked/unbooked streaks across the next 90 days.
  // Goal: unbooked streak lengths should average around ~5 nights (std ~3),
  // and we should avoid too many 1-night unbooked gaps.
  const today = new Date();
  today.setHours(12, 0, 0, 0); // noon reduces DST off-by-one in date slicing

  const horizonDays = 90;
  const baseRevenuePerNight = 180;

  // Seeded PRNG so "Use sample data" feels random but is stable per click.
  function mulberry32(a) {
    return function () {
      let t = (a += 0x6d2b79f5);
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const rng = mulberry32((Date.now() >>> 0) ^ ((today.getDate() * 100000 + today.getMonth() * 1000 + today.getFullYear()) >>> 0));
  const randInt = (minInclusive, maxExclusive) => Math.floor(rng() * (maxExclusive - minInclusive)) + minInclusive;

  // Box-Muller transform for normal-ish values from uniform RNG.
  function randNormal(mean, stdDev) {
    // Avoid log(0)
    const u1 = Math.max(1e-9, rng());
    const u2 = Math.max(1e-9, rng());
    const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
    return mean + z0 * stdDev;
  }

  // Build night-level "booked" boolean array: booked=true means that night is booked.
  const booked = new Array(horizonDays).fill(false);

  // Decide which streak starts first; either is fine.
  let state = rng() < 0.5 ? 'booked' : 'unbooked';
  let pos = 0;

  while (pos < horizonDays) {
    if (state === 'booked') {
      // Booked streak lengths don't need to match a specific distribution.
      // Use longer blocks to avoid creating tiny unbooked gaps.
      const len = Math.min(horizonDays - pos, randInt(3, 10)); // 3..9 nights
      for (let d = pos; d < pos + len; d++) booked[d] = true;
      pos += len;
      state = 'unbooked';
    } else {
      // Unbooked streak lengths target: mean ~5, std ~3.
      // We also downweight/reject 1-night streaks so you don't see too many.
      let x = randNormal(5, 3);
      let len;
      if (x < 1.5) {
        // Most low-tail events become 2-night gaps; occasionally allow 1-night.
        len = rng() < 0.15 ? 1 : 2;
      } else {
        len = Math.round(x);
        len = Math.max(2, len); // keep it from collapsing into many 1-night gaps
      }

      // Cap extremely long gaps.
      len = Math.min(len, 20);
      len = Math.min(horizonDays - pos, len);

      pos += len;
      state = 'booked';
    }
  }

  // Convert booked nights into consecutive blocks => reservation objects.
  const out = [];
  let i = 0;
  while (i < horizonDays) {
    if (!booked[i]) {
      i++;
      continue;
    }

    let j = i;
    while (j < horizonDays && booked[j]) j++;

    const len = j - i;
    if (len > 0) {
      const start = new Date(today);
      start.setDate(today.getDate() + i);

      const end = new Date(today);
      end.setDate(today.getDate() + j);

      // Add variation so revenue/ADR doesn't look perfectly uniform.
      const perNight = baseRevenuePerNight * (0.85 + rng() * 0.35);
      const rev = perNight * len;

      out.push({
        checkIn: formatLocalNoonToYMD(start),
        checkOut: formatLocalNoonToYMD(end),
        revenue: Math.round(rev),
      });
    }

    i = j;
  }

  // Safety: keep only reservations that end in the future.
  const now = new Date();
  return out.filter((r) => new Date(r.checkOut) > now);
}

// --- CSV line split (handles quoted fields) ---
function splitCSVLine(line) {
  const result = [];
  let i = 0;
  let field = '';
  let inQuotes = false;
  while (i < line.length) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += c;
      i++;
      continue;
    }
    if (c === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (c === ',') {
      result.push(field.trim());
      field = '';
      i++;
      continue;
    }
    field += c;
    i++;
  }
  result.push(field.trim());
  return result.map((f) => f.replace(/^"|"$/g, ''));
}

/** CSV text → array of row arrays (including header). */
function csvTextToTableRows(text) {
  const lines = text.trim().split(/\r?\n/).filter((l) => l.length);
  if (!lines.length) return [];
  return lines.map((l) => splitCSVLine(l));
}

/**
 * Excel serial day → YYYY-MM-DD in local calendar (matches Excel’s date column in the UI).
 * Avoids UTC toISOString(), which shifts the calendar day in many timezones.
 */
function excelSerialDateToYMD(serial) {
  const n = Math.floor(Number(serial));
  if (!Number.isFinite(n) || n < 1) return null;
  const d = new Date(1900, 0, 1, 12, 0, 0, 0);
  d.setDate(d.getDate() + n - 1);
  if (isNaN(d.getTime())) return null;
  return toYMDLocal(d);
}

/** Normalize one cell from SheetJS / CSV row for header + data matching. */
function stringifyCellForImport(c) {
  if (c == null || c === '') return '';
  if (Object.prototype.toString.call(c) === '[object Date]' && !isNaN(c.getTime())) {
    return formatLocalNoonToYMD(c);
  }
  if (typeof c === 'number' && Number.isFinite(c)) {
    const whole = Math.floor(Math.abs(c));
    if (whole >= 35000 && whole <= 65000) {
      const ymd = excelSerialDateToYMD(c);
      if (ymd) return ymd;
    }
  }
  return String(c).trim();
}

function isSpreadsheetFile(file) {
  if (!file || !file.name) return false;
  const n = file.name.toLowerCase();
  if (n.endsWith('.xlsx') || n.endsWith('.xls')) return true;
  const t = (file.type || '').toLowerCase();
  return (
    t === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    t === 'application/vnd.ms-excel' ||
    t === 'application/vnd.ms-excel.sheet.macroenabled.12'
  );
}

/**
 * First sheet → row matrix of strings (for same pipeline as CSV).
 * Requires global `XLSX` from SheetJS (see index.html).
 */
function parseWorkbookArrayBufferToRows(arrayBuffer) {
  if (typeof XLSX === 'undefined') {
    return {
      error:
        'Excel support needs the SheetJS script. Check your network, or save as CSV and try again.',
    };
  }
  try {
    const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    const name = wb.SheetNames[0];
    if (!name) return { error: 'Workbook has no sheets.' };
    const sheet = wb.Sheets[name];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
    const rows = raw.map((row) => {
      const arr = Array.isArray(row) ? row : [];
      return arr.map((c) => stringifyCellForImport(c));
    });
    return { rows };
  } catch (err) {
    return { error: err.message || String(err) };
  }
}

function normalizeCsvHeader(h) {
  return String(h)
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_');
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function parsePercentLike(cell) {
  if (cell == null || String(cell).trim() === '') return null;
  const n = parseFloat(String(cell).replace(/%/g, '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function parseMoneyLike(cell) {
  if (cell == null || String(cell).trim() === '') return null;
  const n = parseFloat(String(cell).replace(/[$€£,\s]/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

/** First date in cell → YYYY-MM-DD, or null */
function parseFlexibleDateYMD(cell) {
  const s = String(cell).trim().replace(/^["']|["']$/g, '');
  if (!s) return null;
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (m) {
    const mm = String(m[1]).padStart(2, '0');
    const dd = String(m[2]).padStart(2, '0');
    return `${m[3]}-${mm}-${dd}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatLocalNoonToYMD(d);
  return null;
}

function findFirstColumnIndex(headersNorm, exactList, substringHints) {
  for (const e of exactList) {
    const idx = headersNorm.findIndex((h) => h === e || h.replace(/_/g, '') === e.replace(/_/g, ''));
    if (idx >= 0) return idx;
  }
  if (substringHints && substringHints.length) {
    for (const hint of substringHints) {
      const idx = headersNorm.findIndex((h) => h.includes(hint));
      if (idx >= 0) return idx;
    }
  }
  return -1;
}

/**
 * Detect market pacing CSV vs booking CSV from headers (normalized).
 * Returns shape incl. optional price columns: finalPriceIdx, priceP25Idx, priceP50Idx, priceP75Idx, priceP90Idx
 */
function detectCsvShape(headersNorm) {
  const weekIdx = findFirstColumnIndex(
    headersNorm,
    [
      'date',
      'report_date',
      'as_of_date',
      'period_start',
      'week_start',
      'weekstart',
      'week_of',
      'fri_start',
      'friday_start',
      'stay_week_start',
      'market_week_start',
      'anchor_date',
      'week_begin',
      'periodstart',
    ],
    ['week_start', 'period_start', 'market_week']
  );

  const currentIdx = findFirstColumnIndex(
    headersNorm,
    [
      'market_occ',
      'market_occupancy',
      'mkt_occ',
      'market_occ_current',
      'current_market_occ',
      'occ_market',
      'market_current_occ',
      'occ_current',
      'market_current',
    ],
    ['market_occ', 'mkt_occ']
  );

  const lytIdx = findFirstColumnIndex(
    headersNorm,
    [
      'market_occ_lyt',
      'lyt',
      'last_year_today',
      'market_occ_last_year_today',
      'market_occupancy_(last_year_today)',
      'market_occupancy_last_year_today',
      'occ_lyt',
      'market_lyt',
      'ly_today',
    ],
    ['lyt', 'last_year_today']
  );

  const lyFinalIdx = findFirstColumnIndex(
    headersNorm,
    [
      'market_occ_ly_final',
      'ly_final',
      'market_final_occ',
      'final_market_occ',
      'final_expected_occupancy',
      'occ_ly_final',
      'market_occ_final',
      'market_occupancy_(last_year_final)',
      'market_occupancy_last_year_final',
      'ly_realized',
    ],
    ['ly_final', 'final_market', 'last_year_final']
  );

  const remainingIdx = findFirstColumnIndex(
    headersNorm,
    ['remaining_potential', 'rem_potential', 'booking_potential', 'potential_remaining'],
    ['remaining_potential']
  );

  const periodEndIdx = findFirstColumnIndex(
    headersNorm,
    ['period_end', 'week_end', 'weekend', 'thu_end', 'thursday_end'],
    ['period_end', 'week_end']
  );

  const finalPriceIdx = findFirstColumnIndex(
    headersNorm,
    [
      'final_price',
      'price_final',
      'recommended_price',
      'suggested_price',
      'target_price',
      'dynamic_price',
      'market_final_price',
      'price_recommendation',
      'final_listing_price',
    ],
    ['final_price', 'recommended_price', 'target_price']
  );

  const priceP25Idx = findFirstColumnIndex(
    headersNorm,
    [
      'market_25th_percentile_price',
      'price_p25',
      'p25_price',
      'pctl_25',
      'percentile_25',
      '25th_percentile',
      '25th_percentile_price',
      'price_25th',
      'q1_price',
    ],
    ['market_25th_percentile', 'p25_price', 'percentile_25', '25th_percentile']
  );

  const priceP50Idx = findFirstColumnIndex(
    headersNorm,
    [
      'market_50th_percentile_price',
      'price_p50',
      'p50_price',
      'pctl_50',
      'percentile_50',
      '50th_percentile',
      '50th_percentile_price',
      'median_price',
      'price_median',
    ],
    ['market_50th_percentile', '50th_percentile', 'p50_price', 'percentile_50', 'median_price']
  );

  const priceP75Idx = findFirstColumnIndex(
    headersNorm,
    [
      'market_75th_percentile_price',
      'price_p75',
      'p75_price',
      'pctl_75',
      'percentile_75',
      '75th_percentile',
      '75th_percentile_price',
      'q3_price',
    ],
    ['market_75th_percentile', 'p75_price', 'percentile_75', '75th_percentile']
  );

  const priceP90Idx = findFirstColumnIndex(
    headersNorm,
    [
      'market_90th_percentile_price',
      'price_p90',
      'p90_price',
      'pctl_90',
      'percentile_90',
      '90th_percentile',
      '90th_percentile_price',
    ],
    ['market_90th_percentile', 'p90_price', 'percentile_90', '90th_percentile']
  );

  const checkInIdx = headersNorm.findIndex((h) => /check.?in|^checkin$|arrival|^start$/.test(h));
  const checkOutIdx = headersNorm.findIndex((h) => /check.?out|^checkout$|departure|^end$/.test(h));
  const revIdx = headersNorm.findIndex((h) =>
    /revenue|payout|total|amount|earn|gross/.test(h)
  );

  const hasWeek = weekIdx >= 0;
  const hasAnyMarketMetric =
    currentIdx >= 0 ||
    lytIdx >= 0 ||
    lyFinalIdx >= 0 ||
    finalPriceIdx >= 0 ||
    priceP25Idx >= 0 ||
    priceP50Idx >= 0 ||
    priceP75Idx >= 0 ||
    priceP90Idx >= 0;
  const hasRes = checkInIdx >= 0 && checkOutIdx >= 0;

  let kind = 'unknown';
  if (hasWeek && hasAnyMarketMetric) kind = 'market';
  else if (hasRes) kind = 'reservations';

  return {
    kind,
    weekIdx,
    currentIdx,
    lytIdx,
    lyFinalIdx,
    remainingIdx,
    periodEndIdx,
    finalPriceIdx,
    priceP25Idx,
    priceP50Idx,
    priceP75Idx,
    priceP90Idx,
    checkInIdx,
    checkOutIdx,
    revIdx,
  };
}

function parseMarketMetricsFromRows(headersOriginal, dataRows, shape) {
  const rows = [];
  const usedIdx = new Set(
    [
      shape.weekIdx,
      shape.currentIdx,
      shape.lytIdx,
      shape.lyFinalIdx,
      shape.remainingIdx,
      shape.periodEndIdx,
      shape.finalPriceIdx,
      shape.priceP25Idx,
      shape.priceP50Idx,
      shape.priceP75Idx,
      shape.priceP90Idx,
    ].filter((i) => i >= 0)
  );

  const nCol = headersOriginal.length;

  for (const raw of dataRows) {
    const cells = [];
    for (let i = 0; i < nCol; i++) {
      cells[i] = raw[i] != null ? String(raw[i]).trim() : '';
    }
    if (cells.every((c) => !c)) continue;

    const weekRaw = shape.weekIdx >= 0 ? cells[shape.weekIdx] : '';
    const ymd = parseFlexibleDateYMD(weekRaw);
    if (!ymd) continue;
    const periodStart = getFriThuWeekStartKey(ymd);
    if (!periodStart) continue;

    const currentMarketOccupancy =
      shape.currentIdx >= 0 ? parsePercentLike(cells[shape.currentIdx]) : null;
    const lastYearTodayOccupancy = shape.lytIdx >= 0 ? parsePercentLike(cells[shape.lytIdx]) : null;
    const finalExpectedOccupancy = shape.lyFinalIdx >= 0 ? parsePercentLike(cells[shape.lyFinalIdx]) : null;
    let remainingPotential =
      shape.remainingIdx >= 0 ? parsePercentLike(cells[shape.remainingIdx]) : null;

    if (
      remainingPotential == null &&
      finalExpectedOccupancy != null &&
      currentMarketOccupancy != null
    ) {
      remainingPotential = Math.max(0, finalExpectedOccupancy - currentMarketOccupancy);
    }

    const finalPrice = shape.finalPriceIdx >= 0 ? parseMoneyLike(cells[shape.finalPriceIdx]) : null;
    const priceP25 = shape.priceP25Idx >= 0 ? parseMoneyLike(cells[shape.priceP25Idx]) : null;
    const priceP50 = shape.priceP50Idx >= 0 ? parseMoneyLike(cells[shape.priceP50Idx]) : null;
    const priceP75 = shape.priceP75Idx >= 0 ? parseMoneyLike(cells[shape.priceP75Idx]) : null;
    const priceP90 = shape.priceP90Idx >= 0 ? parseMoneyLike(cells[shape.priceP90Idx]) : null;

    let periodEnd = null;
    if (shape.periodEndIdx >= 0) {
      periodEnd = parseFlexibleDateYMD(cells[shape.periodEndIdx]);
    }

    const extras = {};
    for (let ci = 0; ci < headersOriginal.length; ci++) {
      if (usedIdx.has(ci)) continue;
      const key = String(headersOriginal[ci] || `col_${ci}`).trim();
      if (key && cells[ci] !== undefined && String(cells[ci]).trim() !== '') {
        extras[key] = cells[ci].trim();
      }
    }

    rows.push({
      sourceDateYmd: ymd,
      periodStart,
      periodEnd,
      currentMarketOccupancy,
      lastYearTodayOccupancy,
      finalExpectedOccupancy,
      remainingPotential,
      finalPrice,
      priceP25,
      priceP50,
      priceP75,
      priceP90,
      extras,
    });
  }

  if (rows.length === 0) {
    return { rows: [], error: 'No valid rows: need a recognizable date in the week/period column.' };
  }

  return { rows };
}

/**
 * Index uploaded rows by report/anchor date (sourceDateYmd). Multiple rows for the same day are kept
 * so they can be collapsed before averaging across an unbooked date range.
 */
function buildMarketPacingBySourceDay(rows) {
  const map = new Map();
  for (const r of rows) {
    const d = r.sourceDateYmd;
    if (!d) continue;
    if (!map.has(d)) map.set(d, []);
    map.get(d).push(r);
  }
  return map;
}

function averageMarketMetric(candidates, pick) {
  const vals = [];
  for (const r of candidates) {
    const v = pick(r);
    if (v != null && Number.isFinite(Number(v))) vals.push(Number(v));
  }
  if (!vals.length) return null;
  return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length);
}

function averagedMarketRowFromRows(rows, { sourceDateYmd = null, periodStart = null, periodEnd = null } = {}) {
  return {
    sourceDateYmd,
    periodStart: periodStart ?? rows[0]?.periodStart ?? null,
    periodEnd,
    currentMarketOccupancy: averageMarketMetric(rows, (r) => r.currentMarketOccupancy),
    lastYearTodayOccupancy: averageMarketMetric(rows, (r) => r.lastYearTodayOccupancy),
    finalExpectedOccupancy: averageMarketMetric(rows, (r) => r.finalExpectedOccupancy),
    remainingPotential: averageMarketMetric(rows, (r) => r.remainingPotential),
    finalPrice: averageMarketMetric(rows, (r) => r.finalPrice),
    priceP25: averageMarketMetric(rows, (r) => r.priceP25),
    priceP50: averageMarketMetric(rows, (r) => r.priceP50),
    priceP75: averageMarketMetric(rows, (r) => r.priceP75),
    priceP90: averageMarketMetric(rows, (r) => r.priceP90),
    extras: {},
  };
}

/**
 * Mean of market metrics over each calendar day in [periodStartYmd, periodEndYmd] that has upload data.
 * Single day → that row (or mean if duplicate dates in file). Multiple days → mean across those daily values.
 */
function aggregateUploadedMarketForUnbookedPeriod(dayMap, periodStartYmd, periodEndYmd) {
  if (!dayMap || !periodStartYmd || !periodEndYmd || periodEndYmd < periodStartYmd) return null;
  const perDay = [];
  let cur = periodStartYmd;
  while (cur <= periodEndYmd) {
    const sameDayRows = dayMap.get(cur);
    if (sameDayRows && sameDayRows.length) {
      if (sameDayRows.length === 1) perDay.push(sameDayRows[0]);
      else perDay.push(averagedMarketRowFromRows(sameDayRows, { sourceDateYmd: cur }));
    }
    cur = addDaysYMD(cur, 1);
  }
  if (!perDay.length) return null;
  if (perDay.length === 1) return perDay[0];
  return averagedMarketRowFromRows(perDay, { periodStart: periodStartYmd, periodEnd: periodEndYmd });
}

function loadMarketPacingBundle() {
  const raw = localStorage.getItem(STORAGE_KEYS.marketPacing);
  if (!raw) return null;
  try {
    const o = JSON.parse(raw);
    if (!o || !Array.isArray(o.rows)) return null;
    return {
      rows: o.rows,
      meta: o.meta || {},
    };
  } catch {
    return null;
  }
}

function saveMarketPacingBundle(rows, meta) {
  localStorage.setItem(STORAGE_KEYS.marketPacing, JSON.stringify({ rows, meta }));
}

function clearMarketPacingBundle() {
  localStorage.removeItem(STORAGE_KEYS.marketPacing);
}

function getMarketPacingDayMapFromStorage() {
  const b = loadMarketPacingBundle();
  if (!b || !b.rows.length) return null;
  return buildMarketPacingBySourceDay(b.rows);
}

function describeMarketColumnMap(headersOriginal, shape) {
  const cell = (field, idx) => ({
    field,
    header: idx < 0 ? '—' : headersOriginal[idx] || `Column ${idx + 1}`,
  });
  return [
    cell('Week anchor (Fri–Thu)', shape.weekIdx),
    cell('Market occ. (current)', shape.currentIdx),
    cell('Market occ. LYT', shape.lytIdx),
    cell('Market occ. LY final', shape.lyFinalIdx),
    cell('Remaining potential', shape.remainingIdx),
    cell('Period end (optional)', shape.periodEndIdx),
    cell('Final price', shape.finalPriceIdx),
    cell('Price 25th %ile', shape.priceP25Idx),
    cell('Price 50th %ile', shape.priceP50Idx),
    cell('Price 75th %ile', shape.priceP75Idx),
    cell('Price 90th %ile', shape.priceP90Idx),
  ];
}

function renderMarketCsvPanel() {
  const statusEl = document.getElementById('marketCsvStatus');
  const mapWrap = document.getElementById('marketColumnMap');
  const mapBody = document.getElementById('marketColumnMapBody');
  const extraEl = document.getElementById('marketExtraCols');
  const clearBtn = document.getElementById('clearMarketCsv');
  const dropzone = document.getElementById('marketCsvDropzone');
  if (!statusEl || !mapWrap || !mapBody) return;

  const bundle = loadMarketPacingBundle();
  if (!bundle || !bundle.rows.length) {
    statusEl.textContent = '';
    statusEl.className = 'market-csv-status';
    mapWrap.hidden = true;
    mapBody.innerHTML = '';
    if (extraEl) {
      extraEl.hidden = true;
      extraEl.textContent = '';
    }
    if (clearBtn) clearBtn.hidden = true;
    if (dropzone) dropzone.classList.remove('csv-dropzone--has-file');
    return;
  }

  const meta = bundle.meta || {};
  statusEl.className = 'market-csv-status market-csv-status--ok';
  statusEl.innerHTML = `Using <strong>${escapeHtml(meta.fileName || 'saved market CSV')}</strong> — ${bundle.rows.length} week row(s). The future unbooked table uses this file instead of sample market data.`;

  const columnMap = Array.isArray(meta.columnMap) ? meta.columnMap : [];
  mapBody.innerHTML = columnMap
    .map(
      (row) =>
        `<tr><td>${escapeHtml(row.field)}</td><td><code>${escapeHtml(row.header)}</code></td></tr>`
    )
    .join('');
  mapWrap.hidden = false;

  if (extraEl) {
    const extras = meta.extraHeaders;
    if (extras && extras.length) {
      extraEl.hidden = false;
      extraEl.textContent = `Other columns in file (stored per row): ${extras.join(', ')}`;
    } else {
      extraEl.hidden = true;
      extraEl.textContent = '';
    }
  }

  if (clearBtn) clearBtn.hidden = false;
  if (dropzone) dropzone.classList.add('csv-dropzone--has-file');
}

// --- Reservations: table rows (header + data) like CSV / first Excel sheet ---
function parseReservationsFromTableRows(tableRows) {
  if (!tableRows || tableRows.length < 2) return [];
  const headerCells = tableRows[0].map((c) => String(c ?? '').trim());
  const header = headerCells.map(normalizeCsvHeader);
  const checkInIdx = header.findIndex((h) => /check.?in|^checkin$|arrival|^start$/.test(h));
  const checkOutIdx = header.findIndex((h) => /check.?out|^checkout$|departure|^end$/.test(h));
  const revIdx = header.findIndex((h) => /revenue|payout|total|amount|earn/.test(h));
  if (checkInIdx < 0 || checkOutIdx < 0) return [];

  const nCol = headerCells.length;
  const reservations = [];
  for (let i = 1; i < tableRows.length; i++) {
    const raw = tableRows[i] || [];
    const row = [];
    for (let j = 0; j < nCol; j++) {
      row[j] = raw[j] != null ? String(raw[j]).trim() : '';
    }
    if (row.every((c) => !c)) continue;
    const checkIn = row[checkInIdx];
    const checkOut = row[checkOutIdx];
    if (!checkIn || !checkOut) continue;
    const revenue = revIdx >= 0 ? parseFloat(String(row[revIdx]).replace(/[^0-9.-]/g, '')) || 0 : 0;
    const cinYmd = parseFlexibleDateYMD(checkIn) || String(checkIn).slice(0, 10);
    const coutYmd = parseFlexibleDateYMD(checkOut) || String(checkOut).slice(0, 10);
    reservations.push({
      checkIn: cinYmd,
      checkOut: coutYmd,
      revenue,
    });
  }
  return reservations;
}

// --- CSV: expect columns like check_in, check_out, revenue (or check-in, total_payout, etc.) ---
function parseCSV(text) {
  return parseReservationsFromTableRows(csvTextToTableRows(text));
}

// --- Pacing tiers load/save ---
function loadPacingTiers() {
  const raw = localStorage.getItem(STORAGE_KEYS.pacingTiers);
  if (!raw) return JSON.parse(JSON.stringify(DEFAULT_PACING_TIERS));
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) && parsed.length ? parsed : JSON.parse(JSON.stringify(DEFAULT_PACING_TIERS));
  } catch {
    return JSON.parse(JSON.stringify(DEFAULT_PACING_TIERS));
  }
}

function savePacingTiers(tiers) {
  localStorage.setItem(STORAGE_KEYS.pacingTiers, JSON.stringify(tiers));
}

function buildFuturePeriodsWithMarket(reservations, pacingTiers, horizonDays) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const { firstYmd, lastYmd } = getUnbookedWindowYMDBounds(today, horizonDays);
  if (!firstYmd || !lastYmd) return [];

  // Full streaks/chunks up to UNBOOKED_PERIOD_TABLE_LOOKAHEAD_DAYS, then keep rows that
  // touch the selected window — show entire period (do not trim end at lastYmd).
  const allPeriods = getUnbookedPeriods(reservations || [], today, UNBOOKED_PERIOD_TABLE_LOOKAHEAD_DAYS);
  const marketWeeks = Math.min(55, Math.ceil(UNBOOKED_PERIOD_TABLE_LOOKAHEAD_DAYS / 7) + 2);
  const sampleMarket = getSampleMarketOccupancy(today, marketWeeks);
  const uploadedDayMap = getMarketPacingDayMapFromStorage();
  const out = [];
  for (const p of allPeriods) {
    if (!periodOverlapsUnbookedWindow(p, firstYmd, lastYmd)) continue;
    const market = uploadedDayMap
      ? aggregateUploadedMarketForUnbookedPeriod(uploadedDayMap, p.periodStart, p.periodEnd)
      : getMarketForPeriod(sampleMarket, p.periodStart);
    const finalExpected = market ? market.finalExpectedOccupancy : null;
    const current = market ? market.currentMarketOccupancy : null;
    const lastYearTodayOccupancy = market ? market.lastYearTodayOccupancy : null;
    let remainingPotential = null;
    if (market) {
      if (market.remainingPotential != null && Number.isFinite(Number(market.remainingPotential))) {
        remainingPotential = Math.max(0, Number(market.remainingPotential));
      } else if (finalExpected != null && current != null) {
        remainingPotential = Math.max(0, finalExpected - current);
      }
    }
    let remainingPotentialLY = null;
    if (finalExpected != null && lastYearTodayOccupancy != null) {
      remainingPotentialLY = Math.max(0, finalExpected - lastYearTodayOccupancy);
    }
    const tier = finalExpected != null && pacingTiers && pacingTiers.length ? getPricingTier(finalExpected, pacingTiers) : null;
    const finalPrice = market ? market.finalPrice : null;
    const priceP25 = market ? market.priceP25 : null;
    const priceP50 = market ? market.priceP50 : null;
    const priceP75 = market ? market.priceP75 : null;
    const priceP90 = market ? market.priceP90 : null;
    const finalPricePercentile = deriveFinalPricePercentile(finalPrice, priceP25, priceP50, priceP75, priceP90);
    out.push({
      ...p,
      currentMarketOccupancy: current,
      lastYearTodayOccupancy,
      finalExpectedOccupancy: finalExpected,
      finalExpected: finalExpected,
      remainingPotential,
      remainingPotentialLY,
      finalPrice,
      finalPricePercentile,
      priceP25,
      priceP50,
      priceP75,
      priceP90,
      tier,
    });
  }
  return out;
}

function destroyFuturePacingCharts() {
  if (futurePacingChartResizeObserver) {
    futurePacingChartResizeObserver.disconnect();
    futurePacingChartResizeObserver = null;
  }
  if (futurePacingChartInstance) {
    futurePacingChartInstance.destroy();
    futurePacingChartInstance = null;
  }
}

function futurePacingChartLabels(fps) {
  return fps.map((fp) =>
    fp.periodStart === fp.periodEnd
      ? fp.periodStart.slice(5)
      : `${fp.periodStart.slice(5)}→${fp.periodEnd.slice(5)}`
  );
}

function futurePacingStackedBarData(fps, keyA, keyB) {
  return {
    a: fps.map((fp) => {
      const v = fp[keyA];
      return v != null && Number.isFinite(Number(v)) ? Number(v) : 0;
    }),
    b: fps.map((fp) => {
      const v = fp[keyB];
      return v != null && Number.isFinite(Number(v)) ? Number(v) : 0;
    }),
  };
}

function renderFuturePacingCharts(futurePeriodsWithMarket) {
  const wrap = document.getElementById('futurePacingCharts');
  const canvas = document.getElementById('futurePacingChartCombined');
  destroyFuturePacingCharts();
  if (!wrap || !canvas || typeof Chart === 'undefined') {
    if (wrap) wrap.hidden = true;
    return;
  }
  if (!futurePeriodsWithMarket.length) {
    wrap.hidden = true;
    return;
  }
  wrap.hidden = false;

  ensureChartDataLabelsRegistered();

  const fps = futurePeriodsWithMarket;
  const labels = futurePacingChartLabels(fps);
  const cur = futurePacingStackedBarData(fps, 'currentMarketOccupancy', 'remainingPotential');
  const ly = futurePacingStackedBarData(fps, 'lastYearTodayOccupancy', 'remainingPotentialLY');

  const gridColor = 'rgba(255, 255, 255, 0.06)';
  const tickColor = '#9aa5b4';
  const gridLineMajorOnly = (ctx) => {
    const v = ctx.tick?.value;
    if (!Number.isFinite(v)) return 'transparent';
    const r = Math.round(Number(v));
    return r === 30 || r === 50 || r === 70 || r === 90 ? gridColor : 'transparent';
  };

  const tooltipLabel = (ctx, key) => {
    const raw = fps[ctx.dataIndex][key];
    const name = ctx.dataset.label || '';
    if (raw == null || !Number.isFinite(Number(raw))) return `${name}: —`;
    return `${name}: ${Number(raw)}%`;
  };

  const remOverFinalLabel = (remKey) => ({
    display(ctx) {
      const fp = fps[ctx.dataIndex];
      if (!fp) return false;
      const r = fp[remKey];
      const f = fp.finalExpectedOccupancy;
      return (
        r != null &&
        f != null &&
        Number.isFinite(Number(r)) &&
        Number.isFinite(Number(f)) &&
        Number(r) > 0
      );
    },
    formatter(_, ctx) {
      const fp = fps[ctx.dataIndex];
      const r = Math.round(Number(fp[remKey]));
      const f = Math.round(Number(fp.finalExpectedOccupancy));
      return `${r}/${f}`;
    },
    color: 'rgba(15, 18, 22, 0.96)',
    font: { weight: '700', size: 10, family: "'JetBrains Mono', monospace" },
    anchor: 'center',
    align: 'center',
    clip: true,
  });

  futurePacingChartInstance = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Current · market occ.',
          stack: 'current',
          data: cur.a,
          _fieldKey: 'currentMarketOccupancy',
          backgroundColor: 'rgba(118, 124, 136, 0.88)',
          borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
        },
        {
          label: 'Current · remaining potential',
          stack: 'current',
          data: cur.b,
          _fieldKey: 'remainingPotential',
          backgroundColor: 'rgba(63, 210, 95, 0.72)',
          borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
          datalabels: remOverFinalLabel('remainingPotential'),
        },
        {
          label: 'Last Year · market occ. LYT',
          stack: 'lastYear',
          data: ly.a,
          _fieldKey: 'lastYearTodayOccupancy',
          backgroundColor: 'rgba(108, 112, 122, 0.88)',
          borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 4, bottomRight: 4 },
        },
        {
          label: 'Last Year · remaining potential LY',
          stack: 'lastYear',
          data: ly.b,
          _fieldKey: 'remainingPotentialLY',
          backgroundColor: 'rgba(180, 130, 255, 0.72)',
          borderRadius: { topLeft: 4, topRight: 4, bottomLeft: 0, bottomRight: 0 },
          datalabels: remOverFinalLabel('remainingPotentialLY'),
        },
        {
          type: 'line',
          label: 'Fin %',
          yAxisID: 'y1',
          order: 100,
          data: fps.map((fp) => {
            const v = fp.finalPricePercentile;
            return v != null && Number.isFinite(Number(v)) ? Number(v) : null;
          }),
          _fieldKey: 'finalPricePercentile',
          borderColor: 'rgba(255, 214, 140, 1)',
          backgroundColor: 'rgba(229, 192, 123, 0.08)',
          borderWidth: 2.5,
          tension: 0.22,
          pointRadius: 4,
          pointHoverRadius: 6,
          pointBackgroundColor: 'rgba(255, 224, 160, 1)',
          pointBorderColor: 'rgba(20, 24, 30, 1)',
          pointBorderWidth: 2,
          spanGaps: false,
          fill: false,
          datalabels: {
            display(ctx) {
              const fp = fps[ctx.dataIndex];
              const v = fp?.finalPricePercentile;
              return v != null && Number.isFinite(Number(v));
            },
            formatter(_, ctx) {
              return `${Math.round(Number(fps[ctx.dataIndex].finalPricePercentile))}%`;
            },
            color: 'rgba(255, 230, 175, 1)',
            backgroundColor: 'rgba(18, 22, 28, 0.94)',
            borderColor: 'rgba(229, 192, 123, 0.85)',
            borderWidth: 1,
            borderRadius: 6,
            padding: { top: 4, right: 6, bottom: 4, left: 6 },
            font: { weight: '700', size: 10, family: "'JetBrains Mono', monospace" },
            anchor: 'center',
            align: 'top',
            offset: 18,
            clip: false,
          },
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { top: 24, right: 4 },
      },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        datalabels: {
          display: false,
        },
        legend: {
          position: 'bottom',
          labels: { color: tickColor, boxWidth: 12, padding: 10, font: { size: 11 } },
        },
        tooltip: {
          callbacks: {
            label(ctx) {
              const key = ctx.dataset._fieldKey;
              return key ? tooltipLabel(ctx, key) : `${ctx.dataset.label || ''}`;
            },
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          ticks: { color: tickColor, maxRotation: 50, minRotation: 25, autoSkip: true, maxTicksLimit: 24 },
          grid: { display: false },
        },
        y: {
          id: 'y',
          stacked: true,
          beginAtZero: true,
          min: 0,
          max: 100,
          title: {
            display: true,
            text: 'Occupancy',
            color: tickColor,
            font: { size: 11, weight: '600' },
          },
          ticks: {
            color: tickColor,
            stepSize: 10,
            callback: (v) => (Number.isFinite(v) ? `${v}%` : v),
          },
          grid: { color: gridLineMajorOnly },
        },
        y1: {
          position: 'right',
          beginAtZero: true,
          min: 0,
          max: 100,
          title: {
            display: true,
            text: 'Price',
            color: 'rgba(229, 192, 123, 0.95)',
            font: { size: 11, weight: '600' },
          },
          grid: { drawOnChartArea: false },
          ticks: {
            color: 'rgba(229, 192, 123, 0.95)',
            stepSize: 10,
            callback: (v) => (Number.isFinite(v) ? `${v}%` : v),
          },
        },
      },
    },
  });

  const chartScrollEl = document.querySelector('#futurePacingCharts .future-pacing-chart-scroll');
  const chartCanvasWrap = document.getElementById('futurePacingChartCanvasWrap');
  const pxPerCategory = 94;
  const layoutFuturePacingChartWidth = () => {
    if (!chartScrollEl || !chartCanvasWrap || !futurePacingChartInstance) return;
    const n = Math.max(labels.length, 1);
    const hostW = chartScrollEl.clientWidth || 640;
    const dataMinW = n * pxPerCategory;
    const w = Math.max(dataMinW, hostW);
    chartCanvasWrap.style.minWidth = `${w}px`;
    chartCanvasWrap.style.width = `${w}px`;
    futurePacingChartInstance.resize();
  };
  layoutFuturePacingChartWidth();
  if (typeof ResizeObserver !== 'undefined' && chartScrollEl) {
    futurePacingChartResizeObserver = new ResizeObserver(() => layoutFuturePacingChartWidth());
    futurePacingChartResizeObserver.observe(chartScrollEl);
  }
}

function formatDayOfWeekAbbrev(dateStr) {
  // dateStr is expected as YYYY-MM-DD
  if (!dateStr) return 'N/A';

  // Be strict about YYYY-MM-DD parsing to avoid browser locale quirks.
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr).trim());
  let d;
  if (m) {
    const yyyy = Number(m[1]);
    const mm = Number(m[2]) - 1;
    const dd = Number(m[3]);
    // Use noon to minimize DST edge issues.
    d = new Date(yyyy, mm, dd, 12, 0, 0, 0);
  } else {
    d = new Date(`${dateStr}T12:00:00`);
  }

  const dow = d instanceof Date && !isNaN(d.getTime()) ? d.getDay() : NaN; // 0=Sun ... 6=Sat

  // Match the requested example: "Thur" instead of "Thu".
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thur', 'Fri', 'Sat'];
  return !isNaN(dow) ? DOW[dow] || 'N/A' : 'N/A';
}

function formatDowRange(fromDateStr, toDateStr) {
  const start = formatDayOfWeekAbbrev(fromDateStr);
  const end = formatDayOfWeekAbbrev(toDateStr);
  return start !== 'N/A' && end !== 'N/A' ? `${start}-${end}` : 'N/A';
}

/** Table display: whole-number percent or null (use —). */
function formatPctWhole(v) {
  if (v == null || !Number.isFinite(Number(v))) return null;
  return `${Math.round(Number(v))}%`;
}

function formatMoneyWhole(v) {
  if (v == null || !Number.isFinite(Number(v))) return null;
  return `$${Math.round(Number(v)).toLocaleString()}`;
}

/** Compact period for dense table: MM-DD or MM-DD→MM-DD within a year. */
function formatFuturePacingPeriodCompact(periodStart, periodEnd) {
  if (!periodStart || !periodEnd) return '—';
  if (periodStart === periodEnd) return periodStart.slice(5);
  const y1 = periodStart.slice(0, 4);
  const y2 = periodEnd.slice(0, 4);
  if (y1 === y2) return `${periodStart.slice(5)}→${periodEnd.slice(5)}`;
  return `${periodStart.slice(5)}→${periodEnd}`;
}

function normalizeUnbookedHorizonDays(raw) {
  const n = Number(raw);
  if (UNBOOKED_HORIZON_OPTIONS.includes(n)) return n;
  return DEFAULT_UNBOOKED_HORIZON_DAYS;
}

function getUnbookedHorizonDays() {
  const sel = document.getElementById('unbookedHorizon');
  if (sel && sel.value !== '') return normalizeUnbookedHorizonDays(sel.value);
  const stored = localStorage.getItem(STORAGE_KEYS.unbookedHorizonDays);
  return normalizeUnbookedHorizonDays(stored);
}

function updateUnbookedSectionHeading() {
  const unbookedHeading = document.getElementById('unbookedSectionHeading');
  const sel = document.getElementById('unbookedHorizon');
  if (!unbookedHeading) return;
  if (sel && sel.options && sel.selectedIndex >= 0) {
    const opt = sel.options[sel.selectedIndex];
    const label = opt ? opt.textContent.trim() : `${getUnbookedHorizonDays()} days`;
    unbookedHeading.textContent = `Unbooked dates (next ${label})`;
  } else {
    unbookedHeading.textContent = `Unbooked dates (next ${getUnbookedHorizonDays()} days)`;
  }
}

// --- UI ---
function loadStored() {
  const target = localStorage.getItem(STORAGE_KEYS.revenueTarget);
  const period = localStorage.getItem(STORAGE_KEYS.period) || 'month';
  const raw = localStorage.getItem(STORAGE_KEYS.reservations);
  return {
    revenueTarget: target ? parseFloat(target) : null,
    period,
    reservations: raw ? parseReservations(raw) : null,
    pacingTiers: loadPacingTiers(),
  };
}

function saveStored(revenueTarget, period, reservations) {
  if (revenueTarget != null) localStorage.setItem(STORAGE_KEYS.revenueTarget, String(revenueTarget));
  if (period) localStorage.setItem(STORAGE_KEYS.period, period);
  if (reservations && reservations.length)
    localStorage.setItem(STORAGE_KEYS.reservations, JSON.stringify(reservations));
}

function render(reservations, periodKey, revenueTarget) {
  updateUnbookedSectionHeading();
  const horizonDays = getUnbookedHorizonDays();
  const pacingTiers = loadPacingTiers();
  if (!reservations || !reservations.length) {
    document.getElementById('revenue').textContent = '—';
    document.getElementById('adr').textContent = '—';
    document.getElementById('occupancy').textContent = '—';
    document.getElementById('revpar').textContent = '—';
    document.getElementById('nightsBooked').textContent = '—';
    document.getElementById('revenuePace').textContent = '—';
    document.getElementById('occupancyPace').textContent = '—';
    document.getElementById('daysElapsed').textContent = '—';
    document.getElementById('daysTotal').textContent = 'of period';
    document.getElementById('pacingStatus').textContent = '—';
    document.getElementById('pacingStatus').className = 'metric-value';
    document.getElementById('pacingAction').textContent = 'Load data to see pacing.';
    document.getElementById('unbookedDates').innerHTML = 'Load data to see unbooked dates.';
    document.getElementById('recommendations').innerHTML =
      '<li>Upload calendar/revenue data or use sample data.</li>';
    const emptyHint = document.getElementById('futurePacingEmpty');
    const tableEl = document.getElementById('futurePacingTable');
    const chartsWrap = document.getElementById('futurePacingCharts');
    destroyFuturePacingCharts();
    if (chartsWrap) chartsWrap.hidden = true;
    if (emptyHint) emptyHint.hidden = false;
    if (tableEl) tableEl.hidden = true;
    renderMarketCsvPanel();
    return;
  }

  const { start, end } = getPeriodBounds(periodKey);
  const metrics = computePeriodMetrics(reservations, start, end);
  const pacing = computePacing(reservations, periodKey, revenueTarget);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const unbooked = getUnbookedDates(reservations, today, horizonDays);
  const futurePeriodsWithMarket = buildFuturePeriodsWithMarket(reservations, pacingTiers, horizonDays);
  const recs = getRecommendations(
    reservations,
    periodKey,
    revenueTarget,
    unbooked,
    pacingTiers,
    futurePeriodsWithMarket,
    horizonDays
  );

  document.getElementById('revenue').textContent = `$${Math.round(metrics.revenue).toLocaleString()}`;
  document.getElementById('adr').textContent =
    metrics.adr > 0 ? `$${Math.round(metrics.adr).toLocaleString()}` : '—';
  document.getElementById('occupancy').textContent = `${metrics.occupancy.toFixed(1)}%`;
  document.getElementById('revpar').textContent =
    metrics.revpar > 0 ? `$${metrics.revpar.toFixed(0)}` : '—';
  document.getElementById('nightsBooked').textContent = String(metrics.nights);

  document.getElementById('revenuePace').textContent =
    pacing.revenuePace != null ? `${pacing.revenuePace.toFixed(0)}%` : '—';
  document.getElementById('revenuePaceMeta').textContent =
    revenueTarget > 0 ? `of $${revenueTarget.toLocaleString()} target` : 'vs target';
  document.getElementById('occupancyPace').textContent = `${pacing.occupancyPace.toFixed(0)}%`;
  document.getElementById('occupancyPaceMeta').textContent = 'occupied this period';
  document.getElementById('daysElapsed').textContent = pacing.elapsed;
  document.getElementById('daysTotal').textContent = `of ${pacing.totalDays} days`;
  document.getElementById('pacingStatus').textContent = pacing.status;
  document.getElementById('pacingStatus').className = 'metric-value ' + pacing.statusClass;
  document.getElementById('pacingAction').textContent = pacing.action;

  document.getElementById('unbookedDates').innerHTML =
    unbooked.length === 0
      ? `No unbooked nights in the selected window (${horizonDays} nights from today).`
      : unbooked
          .slice(0, 120)
          .map((d) => `<span class="date-chip">${d}</span>`)
          .join('') +
          (unbooked.length > 120 ? ` <span class="date-chip">+${unbooked.length - 120} more</span>` : '');

  document.getElementById('recommendations').innerHTML = recs
    .map((r) => {
      const c = r.type === 'warning' ? 'rec-warning' : r.type === 'action' ? 'rec-action' : '';
      return `<li class="${c}">${r.text}</li>`;
    })
    .join('');

  const tbody = document.getElementById('futurePacingBody');
  const emptyHint = document.getElementById('futurePacingEmpty');
  const tableEl = document.getElementById('futurePacingTable');
  if (futurePeriodsWithMarket.length === 0) {
    if (emptyHint) emptyHint.hidden = false;
    if (tableEl) tableEl.hidden = true;
  } else {
    if (emptyHint) emptyHint.hidden = true;
    if (tableEl) tableEl.hidden = false;
    if (tbody) {
      tbody.innerHTML = futurePeriodsWithMarket
        .map((fp) => {
          const currentStr = formatPctWhole(fp.currentMarketOccupancy) ?? '—';
          const lytStr = formatPctWhole(fp.lastYearTodayOccupancy) ?? '—';
          const finalStr = formatPctWhole(fp.finalExpectedOccupancy) ?? '—';
          const remainingStr = formatPctWhole(fp.remainingPotential) ?? '—';
          const remainingLyStr = formatPctWhole(fp.remainingPotentialLY) ?? '—';
          const finStr = formatMoneyWhole(fp.finalPrice) ?? '—';
          const finPctStr = formatPctWhole(fp.finalPricePercentile) ?? '—';
          const p25Str = formatMoneyWhole(fp.priceP25) ?? '—';
          const p50Str = formatMoneyWhole(fp.priceP50) ?? '—';
          const p75Str = formatMoneyWhole(fp.priceP75) ?? '—';
          const p90Str = formatMoneyWhole(fp.priceP90) ?? '—';
          const tier = fp.tier;
          const pctStr = tier ? `${tier.minPercentile}-${tier.maxPercentile}%` : '—';
          const tierTitle = tier ? `${tier.minPercentile}-${tier.maxPercentile}% · ${tier.label}` : '';
          const tierTitleAttr = tierTitle
            ? ` title="${tierTitle.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"`
            : '';
          const tierClass = tier ? `tier-${tier.tierClass} pct-cell` : 'pct-cell';
          const dowRange = formatDowRange(fp.periodStart, fp.periodEnd);
          const periodCompact = formatFuturePacingPeriodCompact(fp.periodStart, fp.periodEnd);
          const periodTitle = `${fp.periodStart} – ${fp.periodEnd}`;
          return `<tr>
            <td class="future-pacing-period-cell" title="${periodTitle}">${periodCompact}</td>
            <td class="future-pacing-dow-cell" title="${dowRange}">${dowRange}</td>
            <td class="future-pacing-num-cell">${fp.unbookedNights}</td>
            <td class="pct-cell">${currentStr}</td>
            <td class="pct-cell">${lytStr}</td>
            <td class="pct-cell">${finalStr}</td>
            <td class="pct-cell">${remainingStr}</td>
            <td class="pct-cell">${remainingLyStr}</td>
            <td class="money-cell">${finStr}</td>
            <td class="pct-cell">${finPctStr}</td>
            <td class="money-cell">${p25Str}</td>
            <td class="money-cell">${p50Str}</td>
            <td class="money-cell">${p75Str}</td>
            <td class="money-cell">${p90Str}</td>
            <td class="${tierClass}"${tierTitleAttr}>${pctStr}</td>
          </tr>`;
        })
        .join('');
    }
  }
  renderFuturePacingCharts(futurePeriodsWithMarket);
  renderMarketCsvPanel();
}

function init() {
  const stored = loadStored();
  let reservations = stored.reservations;

  const periodSelect = document.getElementById('period');
  periodSelect.value = stored.period || 'month';

  const targetInput = document.getElementById('revenueTarget');
  if (stored.revenueTarget != null) targetInput.value = stored.revenueTarget;

  const unbookedSelect = document.getElementById('unbookedHorizon');
  if (unbookedSelect) {
    const saved = localStorage.getItem(STORAGE_KEYS.unbookedHorizonDays);
    const days = normalizeUnbookedHorizonDays(saved);
    unbookedSelect.value = String(days);
    unbookedSelect.addEventListener('change', () => {
      const d = normalizeUnbookedHorizonDays(unbookedSelect.value);
      localStorage.setItem(STORAGE_KEYS.unbookedHorizonDays, String(d));
      unbookedSelect.value = String(d);
      updateUnbookedSectionHeading();
      render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
    });
  }
  updateUnbookedSectionHeading();

  const marketCsvInput = document.getElementById('marketCsvInput');
  const marketDropzone = document.getElementById('marketCsvDropzone');
  const marketCsvStatus = document.getElementById('marketCsvStatus');

  const showMarketDropError = (msg) => {
    if (!marketCsvStatus) return;
    marketCsvStatus.className = 'market-csv-status market-csv-status--err';
    marketCsvStatus.textContent = msg;
  };

  const applyTopZoneTableRows = (tableRows, fileName) => {
    const trimmed = (tableRows || []).filter((r) => r && r.some((c) => String(c).trim() !== ''));
    if (!trimmed.length) {
      showMarketDropError('File is empty.');
      return;
    }
    if (trimmed.length < 2) {
      showMarketDropError('Need a header row and at least one data row.');
      return;
    }
    const headersOriginal = trimmed[0].map((h) => String(h ?? '').trim());
    const headersNorm = headersOriginal.map(normalizeCsvHeader);
    const shape = detectCsvShape(headersNorm);

    const nCol = headersOriginal.length;
    const dataRows = trimmed.slice(1).map((raw) => {
      const row = [];
      for (let i = 0; i < nCol; i++) {
        row[i] = raw[i] != null ? String(raw[i]).trim() : '';
      }
      return row;
    });

    if (shape.kind === 'market') {
      const result = parseMarketMetricsFromRows(headersOriginal, dataRows, shape);
      if (result.error) {
        showMarketDropError(result.error);
        return;
      }
      const extraHeaders = [...new Set(result.rows.flatMap((r) => Object.keys(r.extras || {})))].sort();
      saveMarketPacingBundle(result.rows, {
        fileName: fileName || 'market.csv',
        columnMap: describeMarketColumnMap(headersOriginal, shape),
        extraHeaders,
        detectedHeaders: headersOriginal,
        loadedAt: new Date().toISOString(),
      });
      render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
      return;
    }

    if (shape.kind === 'reservations') {
      const parsed = parseReservationsFromTableRows(trimmed);
      if (!parsed.length) {
        showMarketDropError('No valid reservation rows (need check_in and check_out columns).');
        return;
      }
      reservations = parsed;
      saveStored(parseFloat(targetInput.value) || null, periodSelect.value, reservations);
      render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
      return;
    }

    const hint =
      'Market file: a week column (e.g. week_start, period_start) plus at least one of market_occ, lyt, ly_final, or price columns (final_price, price_p25–p90). ' +
      'Booking file: check_in + check_out.';
    showMarketDropError(
      `Could not detect file type. Columns: ${headersOriginal.join(', ')}. ${hint}`
    );
  };

  const applyTopZoneCsvText = (text, fileName) => {
    applyTopZoneTableRows(csvTextToTableRows(text), fileName);
  };

  const ingestTopZoneFile = (file) => {
    if (isSpreadsheetFile(file)) {
      const reader = new FileReader();
      reader.onload = () => {
        const out = parseWorkbookArrayBufferToRows(reader.result);
        if (out.error) {
          showMarketDropError(out.error);
          return;
        }
        applyTopZoneTableRows(out.rows, file.name);
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = () => applyTopZoneCsvText(String(reader.result), file.name);
      reader.readAsText(file);
    }
  };

  if (marketDropzone && marketCsvInput) {
    marketDropzone.addEventListener('click', () => marketCsvInput.click());
    marketDropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        marketCsvInput.click();
      }
    });
    marketDropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      marketDropzone.classList.add('csv-dropzone--dragover');
    });
    marketDropzone.addEventListener('dragleave', () => {
      marketDropzone.classList.remove('csv-dropzone--dragover');
    });
    marketDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      marketDropzone.classList.remove('csv-dropzone--dragover');
      const f = e.dataTransfer.files?.[0];
      if (!f) return;
      ingestTopZoneFile(f);
    });
    marketCsvInput.addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      ingestTopZoneFile(f);
      e.target.value = '';
    });
  }

  const clearMarketCsvBtn = document.getElementById('clearMarketCsv');
  if (clearMarketCsvBtn) {
    clearMarketCsvBtn.addEventListener('click', () => {
      clearMarketPacingBundle();
      render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
    });
  }

  document.getElementById('saveTarget').addEventListener('click', () => {
    const v = parseFloat(targetInput.value);
    if (!isNaN(v) && v >= 0) {
      localStorage.setItem(STORAGE_KEYS.revenueTarget, String(v));
      render(reservations, periodSelect.value, v);
    }
  });

  periodSelect.addEventListener('change', () => {
    const p = periodSelect.value;
    localStorage.setItem(STORAGE_KEYS.period, p);
    render(reservations, p, parseFloat(targetInput.value) || null);
  });

  document.getElementById('useSample').addEventListener('click', () => {
    reservations = getSampleReservations();
    saveStored(
      parseFloat(targetInput.value) || null,
      periodSelect.value,
      reservations
    );
    render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
  });

  const fetchHospitableBtn = document.getElementById('fetchHospitable');
  if (fetchHospitableBtn) {
    fetchHospitableBtn.addEventListener('click', async () => {
      const originalText = fetchHospitableBtn.textContent;
      fetchHospitableBtn.textContent = 'Loading…';
      fetchHospitableBtn.disabled = true;
      try {
        const startDate = '1970-01-01';
        const endDate = (() => {
          const d = new Date();
          d.setFullYear(d.getFullYear() + 1);
          return formatLocalNoonToYMD(d);
        })();
        const res = await fetch(
          `${getApiBase()}/api/hospitable/reservations?startDate=${encodeURIComponent(startDate)}&endDate=${encodeURIComponent(endDate)}`
        );
        const text = await res.text();
        let data;
        try {
          data = text ? JSON.parse(text) : {};
        } catch {
          throw new Error(`Bad response (${res.status}): ${text.slice(0, 200)}`);
        }
        if (!res.ok) {
          const detail =
            typeof data.details === 'string'
              ? data.details
              : data.details && JSON.stringify(data.details);
          throw new Error(
            [data.error || `Request failed (${res.status})`, detail, data.hint].filter(Boolean).join(' — ')
          );
        }
        const raw = data.reservations || [];
        const diag = data.diagnostics;
        if (raw.length === 0 && diag && diag.rawReservationObjects > 0) {
          console.warn('Hospitable diagnostics:', diag);
          throw new Error(
            `API returned ${diag.rawReservationObjects} reservation row(s) but 0 usable stays (mapping failed). ` +
              `Dropped by mapping: ${diag.droppedByMapping}. ` +
              (diag.checkoutQueryFailed
                ? `Checkout query failed (check-in only): ${JSON.stringify(diag.checkoutQueryError)}. `
                : '') +
              'Open DevTools → Network → this request → Response for full diagnostics.'
          );
        }
        reservations = raw.map((r) => ({
          checkIn: r.checkIn,
          checkOut: r.checkOut,
          revenue: Number(r.revenue) || 0,
        }));
        saveStored(
          parseFloat(targetInput.value) || null,
          periodSelect.value,
          reservations
        );
        render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
      } catch (err) {
        alert('Fetch from Hospitable failed. Is the backend running (npm start)? ' + (err.message || String(err)));
      } finally {
        fetchHospitableBtn.textContent = originalText;
        fetchHospitableBtn.disabled = false;
      }
    });
  }

  document.getElementById('csvInput').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (isSpreadsheetFile(file)) {
      const reader = new FileReader();
      reader.onload = () => {
        const out = parseWorkbookArrayBufferToRows(reader.result);
        if (out.error) {
          alert(out.error);
          return;
        }
        reservations = parseReservationsFromTableRows(out.rows);
        if (!reservations.length) {
          alert('No valid reservation rows (need check_in and check_out columns in the first sheet).');
          return;
        }
        saveStored(
          parseFloat(targetInput.value) || null,
          periodSelect.value,
          reservations
        );
        render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = () => {
        reservations = parseCSV(reader.result);
        saveStored(
          parseFloat(targetInput.value) || null,
          periodSelect.value,
          reservations
        );
        render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
      };
      reader.readAsText(file);
    }
    e.target.value = '';
  });

  document.getElementById('openSettings').addEventListener('click', () => {
    openPacingSettingsModal();
  });
  document.getElementById('closeSettings').addEventListener('click', () => {
    document.getElementById('settingsOverlay').hidden = true;
  });
  document.getElementById('savePacingSettings').addEventListener('click', () => {
    savePacingSettingsFromForm();
    document.getElementById('settingsOverlay').hidden = true;
    render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
  });
  document.getElementById('resetPacingSettings').addEventListener('click', () => {
    savePacingTiers(JSON.parse(JSON.stringify(DEFAULT_PACING_TIERS)));
    populateTiersTable(DEFAULT_PACING_TIERS);
    render(reservations, periodSelect.value, parseFloat(targetInput.value) || null);
  });
  document.getElementById('settingsOverlay').addEventListener('click', (e) => {
    if (e.target.id === 'settingsOverlay') e.target.hidden = true;
  });

  render(
    reservations,
    periodSelect.value,
    stored.revenueTarget != null ? stored.revenueTarget : parseFloat(targetInput.value) || null
  );
}

function openPacingSettingsModal() {
  const tiers = loadPacingTiers();
  populateTiersTable(tiers);
  document.getElementById('settingsOverlay').hidden = false;
}

function populateTiersTable(tiers) {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  tbody.innerHTML = tiers
    .map(
      (t, i) => `
    <tr data-tier-id="${t.id || i}">
      <td>
        <div class="tier-range">
          <input type="number" min="0" max="100" value="${t.minFinalOccupancy}" data-field="minFinalOccupancy" />
          <span>% to</span>
          <input type="number" min="0" max="100" value="${t.maxFinalOccupancy}" data-field="maxFinalOccupancy" />
          <span>%</span>
        </div>
      </td>
      <td>
        <div class="tier-range">
          <input type="number" min="0" max="100" value="${t.minPercentile}" data-field="minPercentile" />
          <span>–</span>
          <input type="number" min="0" max="100" value="${t.maxPercentile}" data-field="maxPercentile" />
          <span>%</span>
        </div>
      </td>
    </tr>`
    )
    .join('');
}

function savePacingSettingsFromForm() {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  const tiers = loadPacingTiers();
  const rows = tbody.querySelectorAll('tr');
  const newTiers = [];
  rows.forEach((row, i) => {
    const minFinal = row.querySelector('[data-field="minFinalOccupancy"]');
    const maxFinal = row.querySelector('[data-field="maxFinalOccupancy"]');
    const minPct = row.querySelector('[data-field="minPercentile"]');
    const maxPct = row.querySelector('[data-field="maxPercentile"]');
    const base = tiers[i] || { id: `tier-${i}`, label: ['High demand', 'Mid demand', 'Low demand'][i] || 'Tier' };
    newTiers.push({
      ...base,
      minFinalOccupancy: minFinal ? parseInt(minFinal.value, 10) || 0 : base.minFinalOccupancy,
      maxFinalOccupancy: maxFinal ? parseInt(maxFinal.value, 10) || 100 : base.maxFinalOccupancy,
      minPercentile: minPct ? parseInt(minPct.value, 10) || 0 : base.minPercentile,
      maxPercentile: maxPct ? parseInt(maxPct.value, 10) || 100 : base.maxPercentile,
    });
  });
  if (newTiers.length) savePacingTiers(newTiers);
}

init();
