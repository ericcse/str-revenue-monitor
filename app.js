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
};

/** Allowed unbooked lookahead windows (≈ calendar months). Default 3 months. */
const UNBOOKED_HORIZON_OPTIONS = [30, 90, 180, 270, 365];
const DEFAULT_UNBOOKED_HORIZON_DAYS = 90;

/**
 * How far ahead we scan when building Fri–Thu unbooked rows for the Future table.
 * Must cover max dropdown (12 months) so a period that starts inside the selected window
 * but ends after the cutoff still shows as one full row (not truncated at lastYmd).
 */
const UNBOOKED_PERIOD_TABLE_LOOKAHEAD_DAYS = 370;

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

// --- Sample market occupancy by week (current vs final expected from "last year"). Monday-based weeks to match unbooked periods. ---
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
  for (let w = 0; w < weeksAhead; w++) {
    const weekStart = new Date(friday);
    weekStart.setDate(friday.getDate() + w * 7);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6); // Thursday
    const key = weekStart.toISOString().slice(0, 10); // Friday key
    const lastYearSameWeek = (52 + w * 2 + (weekStart.getMonth() % 3) * 5) % 100;
    const finalExpected = Math.min(92, Math.max(28, lastYearSameWeek + 35));
    const current = Math.min(finalExpected * 0.95, finalExpected * (0.2 + (w / weeksAhead) * 0.6));
    out.push({
      periodStart: key,
      periodEnd: weekEnd.toISOString().slice(0, 10),
      currentMarketOccupancy: Math.round(current),
      finalExpectedOccupancy: Math.round(finalExpected),
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
        checkIn: start.toISOString().slice(0, 10),
        checkOut: end.toISOString().slice(0, 10),
        revenue: Math.round(rev),
      });
    }

    i = j;
  }

  // Safety: keep only reservations that end in the future.
  const now = new Date();
  return out.filter((r) => new Date(r.checkOut) > now);
}

// --- CSV: expect columns like check_in, check_out, revenue (or check-in, total_payout, etc.) ---
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const header = lines[0].toLowerCase().split(',').map((h) => h.trim().replace(/\s+/g, '_'));
  const checkInIdx = header.findIndex((h) => /check.?in|arrival|start/.test(h));
  const checkOutIdx = header.findIndex((h) => /check.?out|departure|end/.test(h));
  const revIdx = header.findIndex((h) => /revenue|payout|total|amount|earn/.test(h));
  if (checkInIdx < 0 || checkOutIdx < 0) return [];

  const reservations = [];
  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(',').map((c) => c.trim().replace(/^["']|["']$/g, ''));
    const checkIn = row[checkInIdx];
    const checkOut = row[checkOutIdx];
    if (!checkIn || !checkOut) continue;
    const revenue = revIdx >= 0 ? parseFloat(String(row[revIdx]).replace(/[^0-9.-]/g, '')) || 0 : 0;
    reservations.push({
      checkIn: checkIn.slice(0, 10),
      checkOut: checkOut.slice(0, 10),
      revenue,
    });
  }
  return reservations;
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
  const marketData = getSampleMarketOccupancy(today, marketWeeks);
  const out = [];
  for (const p of allPeriods) {
    if (!periodOverlapsUnbookedWindow(p, firstYmd, lastYmd)) continue;
    const market = getMarketForPeriod(marketData, p.periodStart);
    const finalExpected = market ? market.finalExpectedOccupancy : null;
    const current = market ? market.currentMarketOccupancy : null;
    const remainingPotential = finalExpected != null && current != null ? Math.max(0, finalExpected - current) : null;
    const tier = finalExpected != null && pacingTiers && pacingTiers.length ? getPricingTier(finalExpected, pacingTiers) : null;
    out.push({
      ...p,
      currentMarketOccupancy: current,
      finalExpectedOccupancy: finalExpected,
      finalExpected: finalExpected,
      remainingPotential,
      tier,
    });
  }
  return out;
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
    if (emptyHint) emptyHint.hidden = false;
    if (tableEl) tableEl.hidden = true;
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
          const currentStr = fp.currentMarketOccupancy != null ? `${fp.currentMarketOccupancy}%` : '—';
          const finalStr = fp.finalExpectedOccupancy != null ? `${fp.finalExpectedOccupancy}%` : '—';
          const remainingStr = fp.remainingPotential != null ? `${fp.remainingPotential}%` : '—';
          const tier = fp.tier;
          const pctStr = tier ? `${tier.minPercentile}-${tier.maxPercentile}% (${tier.label})` : '—';
          const tierClass = tier ? `tier-${tier.tierClass}` : '';
          const dowRange = formatDowRange(fp.periodStart, fp.periodEnd);
          return `<tr>
            <td>${fp.periodStart} – ${fp.periodEnd}</td>
            <td>${dowRange}</td>
            <td>${fp.unbookedNights}</td>
            <td class="pct-cell">${currentStr}</td>
            <td class="pct-cell">${finalStr}</td>
            <td class="pct-cell">${remainingStr}</td>
            <td class="${tierClass}">${pctStr}</td>
          </tr>`;
        })
        .join('');
    }
  }
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
          return d.toISOString().slice(0, 10);
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
