/**
 * STR Revenue Monitor — Pacing & revenue management
 * Works with CSV upload or sample data. Ready for Hospitable/PriceLabs API later.
 */

const STORAGE_KEYS = {
  revenueTarget: 'str-monitor-revenue-target',
  reservations: 'str-monitor-reservations',
  /** Two-column pacing rules (replaces legacy range-based tiers). */
  pacingRows: 'str-monitor-pacing-rows',
  /** User-defined template for “Reset to defaults” (set via Save as Defaults). */
  pacingRowsDefaults: 'str-monitor-pacing-rows-defaults',
  unbookedHorizonDays: 'str-monitor-unbooked-horizon-days',
  marketPacing: 'str-monitor-market-pacing',
  /** ISO 8601 instant when reservations were last loaded from Hospitable (client clock). */
  hospitableAsOf: 'str-monitor-hospitable-as-of',
  /** Future pacing table: extra columns visible (Cur, LYT, Rm LY, P25–P90). */
  futurePacingTableDetailsVisible: 'str-monitor-future-table-details',
};

/** Revenue/pacing cards use this calendar window (month = current month only). */
const PACING_PERIOD_KEY = 'month';

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
let revenueMetricsChartInstance = null;
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

// Default thresholds (highest first). 0% = catch-all. Optional two-step ramp: Target → Lower @ Wks ahead → Lowest @ stay.
const DEFAULT_PACING_ROWS = [
  {
    id: 'r-high',
    finalOccupancyPct: 75,
    targetPricePercentile: 80,
    lowestPricePercentile: 50,
    weeksToStartReducing: 3,
    lowerPercentileTo: 65,
    weeksAhead: 1,
  },
  {
    id: 'r-mid',
    finalOccupancyPct: 50,
    targetPricePercentile: 55,
    lowestPricePercentile: 25,
    weeksToStartReducing: 3,
    lowerPercentileTo: 40,
    weeksAhead: 1,
  },
  {
    id: 'r-low',
    finalOccupancyPct: 0,
    targetPricePercentile: 30,
    lowestPricePercentile: 15,
    weeksToStartReducing: 3,
    lowerPercentileTo: 22,
    weeksAhead: 1,
  },
];

function clampPct0to100(n) {
  const x = Math.round(Number(n));
  if (!Number.isFinite(x)) return 0;
  return Math.max(0, Math.min(100, x));
}

/** Weeks before stay to start ramping target → lowest (0 = always use target). One decimal place. */
function clampWeeksNonNeg(n) {
  const x = Number(n);
  if (!Number.isFinite(x) || x < 0) return 0;
  const r = Math.round(x * 10) / 10;
  return Math.min(104, r);
}

/** Wks to Lower / Wks to Lowest shown with exactly one decimal (e.g. 3 → 3.0). */
function formatWeeksOneDecimal(n) {
  const x = Number(n);
  if (!Number.isFinite(x)) return '—';
  return (Math.round(x * 10) / 10).toFixed(1);
}

function newPacingRowId() {
  return `r-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`;
}

/**
 * Lower Pctl % and Wks to Lower (vs Wks to Lowest) for two-segment ramp.
 * If invalid or redundant, weeksAhead becomes 0 (single segment Target → Lowest).
 */
function finalizeLowerAndWeeksAhead(tp, low, wks, lowerRaw, weeksAheadRaw) {
  const tgt = clampPct0to100(tp);
  let lo = clampPct0to100(low);
  if (lo > tgt) lo = tgt;
  const wksR = clampWeeksNonNeg(wks);
  if (wksR <= 0) {
    return { lowerPercentileTo: lo, weeksAhead: 0 };
  }
  let lower =
    lowerRaw != null && lowerRaw !== '' ? clampPct0to100(lowerRaw) : Math.round((tgt + lo) / 2);
  lower = Math.max(lo, Math.min(lower, tgt));
  let wa =
    weeksAheadRaw != null && weeksAheadRaw !== ''
      ? clampWeeksNonNeg(weeksAheadRaw)
      : wksR > 1
        ? 1
        : 0;
  if (wa >= wksR) wa = 0;
  if (lower <= lo || lower >= tgt) wa = 0;
  return { lowerPercentileTo: lower, weeksAhead: wa };
}

/**
 * After linear blend of pacing rows: **always keep the lerped Wks to Lower** for display/tooltip.
 * Pricing uses computeRecPctlPercentileFromWksAway: bands from Wks Away vs Wks to Lower / Wks to Lowest.
 * between loRamp and tgt; if wa >= wksR (or lower invalid), ramp is single-segment — no need to
 * shrink wa to wksR−1 (that broke e.g. lerp 6→12 with small Wks to Lowest).
 */
function mergeInterpolatedLowerAndWeeksAhead(tpC, lowC, wksC, lowerInterp, weeksAheadInterp) {
  const tgt = clampPct0to100(tpC);
  const floor = clampPct0to100(lowC);
  const loRamp = Math.min(floor, tgt);
  const wksR = clampWeeksNonNeg(wksC);
  if (wksR <= 0) {
    return finalizeLowerAndWeeksAhead(tpC, lowC, wksC, lowerInterp, 0);
  }
  const wa = clampWeeksNonNeg(weeksAheadInterp);
  if (wa === 0) {
    return finalizeLowerAndWeeksAhead(tpC, lowC, wksC, lowerInterp, 0);
  }

  let lower = clampPct0to100(lowerInterp);
  lower = Math.max(floor, Math.min(lower, tgt));

  // Same mid-step validity as legacy two-phase ramp (used when normalizing rows)
  if (wa < wksR && lower > loRamp && lower < tgt) {
    return { lowerPercentileTo: lower, weeksAhead: wa };
  }

  if (wa < wksR && tgt > loRamp + 1) {
    lower = Math.round((tgt + loRamp) / 2);
    lower = Math.max(loRamp + 1, Math.min(tgt - 1, lower));
    if (lower > loRamp && lower < tgt) {
      return { lowerPercentileTo: lower, weeksAhead: wa };
    }
  }

  // wa >= wksR or still no valid mid: single-segment pricing; show true lerped Wks to Lower
  const fin = finalizeLowerAndWeeksAhead(tpC, lowC, wksC, lowerInterp, 0);
  return { lowerPercentileTo: fin.lowerPercentileTo, weeksAhead: wa };
}

/**
 * Keep row.weeksAhead (Wks to Lower) when the user saved it. finalizeLowerAndWeeksAhead() may zero it
 * for curve math, but that must not erase values on Save / Save as defaults / load from storage.
 * Saved Wks to Lower is preserved for the table formula even when finalize zeros it for curve rules.
 */
function coalesceStoredWeeksAhead(row, finWeeksAhead) {
  if (!row) return finWeeksAhead;
  const v = row.weeksAhead;
  if (v == null || v === '') return finWeeksAhead;
  const n = Number(v);
  if (!Number.isFinite(n)) return finWeeksAhead;
  return clampWeeksNonNeg(n);
}

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

/**
 * Forward `numMonths` calendar months starting with the current month (left → right).
 * Booked = guest nights (check-in inclusive, check-out exclusive). Unbooked = days in month − booked.
 */
function buildForwardMonthlyBookedUnbookedNights(reservations, numMonths) {
  const labels = [];
  const bookedArr = [];
  const unbookedArr = [];
  const anchor = new Date();
  anchor.setDate(1);
  anchor.setHours(12, 0, 0, 0);
  const n = Math.max(1, Math.min(12, Math.floor(Number(numMonths)) || 1));
  for (let i = 0; i < n; i++) {
    const d = new Date(anchor.getFullYear(), anchor.getMonth() + i, 1, 12, 0, 0, 0);
    const y = d.getFullYear();
    const m = d.getMonth();
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const bookedSet = new Set();
    for (const r of reservations) {
      if (!r.checkIn || !r.checkOut) continue;
      const cin = parseYMDLocal(r.checkIn);
      const cout = parseYMDLocal(r.checkOut);
      if (isNaN(cin.getTime()) || isNaN(cout.getTime()) || cout <= cin) continue;
      for (let cur = new Date(cin); cur < cout; cur.setDate(cur.getDate() + 1)) {
        if (cur.getFullYear() === y && cur.getMonth() === m) bookedSet.add(toYMDLocal(cur));
      }
    }
    const booked = bookedSet.size;
    const unbooked = Math.max(0, daysInMonth - booked);
    labels.push(d.toLocaleString(undefined, { month: 'short', year: '2-digit' }));
    bookedArr.push(booked);
    unbookedArr.push(unbooked);
  }
  return { labels, booked: bookedArr, unbooked: unbookedArr };
}

/**
 * Forward `numMonths` calendar months starting with current month.
 * Revenue attributed to each reservation's *check-in month* (not prorated by nights).
 */
function buildForwardMonthlyCheckInRevenue(reservations, numMonths) {
  const revenueArr = [];
  const firstRevenueLinesArr = [];
  const anchor = new Date();
  anchor.setDate(1);
  anchor.setHours(12, 0, 0, 0);
  const n = Math.max(1, Math.min(12, Math.floor(Number(numMonths)) || 1));
  for (let i = 0; i < n; i++) {
    const d = new Date(anchor.getFullYear(), anchor.getMonth() + i, 1, 12, 0, 0, 0);
    const y = d.getFullYear();
    const m = d.getMonth();
    let sum = 0;
    const monthEntries = [];
    for (const r of reservations) {
      if (!r.checkIn) continue;
      const cin = parseYMDLocal(r.checkIn);
      if (isNaN(cin.getTime())) continue;
      if (cin.getFullYear() === y && cin.getMonth() === m) {
        const rev = Number(r.revenue) || 0;
        sum += rev;
        monthEntries.push({ checkIn: r.checkIn, revenue: rev });
      }
    }
    revenueArr.push(sum);
    monthEntries.sort((a, b) => String(a.checkIn).localeCompare(String(b.checkIn)));
    firstRevenueLinesArr.push(
      monthEntries
        .slice(0, 4)
        .map((x) => `${String(x.checkIn).slice(5)}: $${Math.round(Number(x.revenue) || 0).toLocaleString()}`)
    );
  }
  return { totals: revenueArr, firstRevenueLines: firstRevenueLinesArr };
}

/** Maps Unbooked dropdown days → how many forward calendar months the revenue chart shows. */
function unbookedHorizonDaysToRevenueChartMonths(days) {
  const d = normalizeUnbookedHorizonDays(days);
  const map = { 30: 1, 120: 4, 180: 6, 270: 9, 365: 12 };
  return map[d] ?? Math.max(1, Math.min(12, Math.round(d / 30)));
}

function destroyRevenueMetricsChart() {
  if (revenueMetricsChartInstance) {
    revenueMetricsChartInstance.destroy();
    revenueMetricsChartInstance = null;
  }
}

function renderRevenueMetricsChart(reservations) {
  const wrap = document.getElementById('revenueMetricsChartWrap');
  const canvas = document.getElementById('revenueMetricsChart');
  if (!wrap || !canvas || typeof Chart === 'undefined') return;
  if (!reservations || !reservations.length) {
    destroyRevenueMetricsChart();
    wrap.hidden = true;
    return;
  }
  const horizonDays = getUnbookedHorizonDays();
  const numMonths = unbookedHorizonDaysToRevenueChartMonths(horizonDays);
  const { labels, booked, unbooked } = buildForwardMonthlyBookedUnbookedNights(reservations, numMonths);
  const { totals: checkInRevenue, firstRevenueLines } = buildForwardMonthlyCheckInRevenue(
    reservations,
    numMonths
  );
  wrap.hidden = false;
  destroyRevenueMetricsChart();
  ensureChartDataLabelsRegistered();
  const tickColor = '#9aa5b4';
  const gridColor = 'rgba(255, 255, 255, 0.06)';
  const revenueTickColor = 'rgba(90, 235, 120, 0.95)';
  revenueMetricsChartInstance = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Nights booked',
          data: booked,
          stack: 'nights',
          backgroundColor: 'rgba(118, 124, 136, 0.74)',
          borderRadius: { topLeft: 0, topRight: 0, bottomLeft: 3, bottomRight: 3 },
        },
        {
          label: 'Nights unbooked',
          data: unbooked,
          stack: 'nights',
          backgroundColor: 'rgba(88, 166, 255, 0.46)',
          borderRadius: { topLeft: 3, topRight: 3, bottomLeft: 0, bottomRight: 0 },
          datalabels: {
            display(ctx) {
              const v = ctx.dataset.data?.[ctx.dataIndex];
              return Number.isFinite(Number(v)) && Number(v) > 0;
            },
            formatter(v) {
              return `${Math.round(Number(v))}`;
            },
            color: 'rgba(235, 241, 248, 0.95)',
            font: { weight: '700', size: 10, family: "'JetBrains Mono', monospace" },
            anchor: 'center',
            align: 'center',
            clip: true,
          },
        },
        {
          label: 'Check-in revenue',
          data: checkInRevenue,
          yAxisID: 'y1',
          stack: 'revenue',
          backgroundColor: 'rgba(63, 185, 80, 0.42)',
          borderColor: 'rgba(63, 185, 80, 0.9)',
          borderWidth: 1,
          borderRadius: { topLeft: 3, topRight: 3, bottomLeft: 0, bottomRight: 0 },
          datalabels: {
            display(ctx) {
              const v = ctx.dataset.data?.[ctx.dataIndex];
              return Number.isFinite(Number(v)) && Number(v) > 0;
            },
            formatter(v) {
              return `$${Math.round(Number(v)).toLocaleString()}`;
            },
            color: 'rgba(90, 235, 120, 0.98)',
            font: { weight: '700', size: 10, family: "'JetBrains Mono', monospace" },
            anchor: 'end',
            align: 'top',
            offset: 2,
            clip: false,
          },
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        datalabels: { display: false },
        legend: {
          position: 'bottom',
          labels: { color: tickColor, boxWidth: 12, padding: 8, font: { size: 10 } },
        },
        tooltip: {
          callbacks: {
            label(ctx) {
              const raw = ctx.raw;
              if (ctx.dataset?.yAxisID === 'y1') {
                const v = Number(raw);
                return `${ctx.dataset.label}: $${Number.isFinite(v) ? Math.round(v).toLocaleString() : '—'}`;
              }
              const v = Number(raw);
              return `${ctx.dataset.label}: ${Number.isFinite(v) ? Math.round(v) : '—'}`;
            },
            footer(tooltipItems) {
              if (!tooltipItems.length) return '';
              const idx = tooltipItems[0].dataIndex;
              const b = booked[idx] ?? 0;
              const u = unbooked[idx] ?? 0;
              const rev = checkInRevenue[idx] ?? 0;
              const lines = [
                `Total nights in month: ${b + u} · Check-in revenue: $${Math.round(Number(rev) || 0).toLocaleString()}`,
              ];
              const first = firstRevenueLines[idx] || [];
              if (first.length) {
                lines.push('First check-ins:');
                lines.push(...first);
              }
              return lines;
            },
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          ticks: { color: tickColor, maxRotation: 45, minRotation: 0, font: { size: 10 } },
          grid: { display: false },
        },
        y: {
          stacked: true,
          beginAtZero: true,
          title: {
            display: true,
            text: 'Nights',
            color: tickColor,
            font: { size: 11, weight: '600' },
          },
          ticks: { color: tickColor, precision: 0 },
          grid: { color: gridColor },
        },
        y1: {
          position: 'right',
          beginAtZero: true,
          title: {
            display: true,
            text: 'Revenue ($) — check-ins',
            color: revenueTickColor,
            font: { size: 11, weight: '600' },
          },
          grid: { drawOnChartArea: false, color: gridColor },
          ticks: {
            color: revenueTickColor,
            callback(v) {
              return Number.isFinite(Number(v)) ? `$${Math.round(Number(v)).toLocaleString()}` : v;
            },
          },
        },
      },
    },
  });
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

/**
 * Normalize one pacing row into tier fields (target, ramp, etc.) for interpolation anchors.
 */
function tierFromPacingRow(r, idx) {
  const tp = clampPct0to100(r.targetPricePercentile);
  let low =
    r.lowestPricePercentile != null && r.lowestPricePercentile !== ''
      ? clampPct0to100(r.lowestPricePercentile)
      : Math.max(0, tp - 30);
  if (low > tp) low = tp;
  const wks = r.weeksToStartReducing != null ? clampWeeksNonNeg(r.weeksToStartReducing) : 3;
  const fin = finalizeLowerAndWeeksAhead(tp, low, wks, r.lowerPercentileTo, r.weeksAhead);
  const weeksAhead = coalesceStoredWeeksAhead(r, fin.weeksAhead);
  const th = Number(r.finalOccupancyPct);
  const id = String(r.id || `row-${idx}`).replace(/[^a-zA-Z0-9_-]/g, '');
  return {
    targetPricePercentile: tp,
    lowestPricePercentile: low,
    weeksToStartReducing: wks,
    lowerPercentileTo: fin.lowerPercentileTo,
    weeksAhead,
    thresholdFinalOcc: th,
    tierClass: id || `row-${idx}`,
  };
}

/**
 * Unique Final Occ. % thresholds, ascending (duplicate % → last row wins). Each anchor is a normalized tier.
 */
function buildPacingInterpolationAnchors(rows) {
  const valid = rows.filter(
    (r) =>
      r &&
      Number.isFinite(Number(r.finalOccupancyPct)) &&
      Number.isFinite(Number(r.targetPricePercentile))
  );
  if (!valid.length) return [];
  const byTh = new Map();
  valid.forEach((r) => {
    byTh.set(Number(r.finalOccupancyPct), r);
  });
  const sortedPairs = [...byTh.entries()].sort((a, b) => a[0] - b[0]);
  return sortedPairs.map(([th, r], i) => tierFromPacingRow(r, i));
}

/** Footer lines: how Rec. Pctl % is computed (shown in cell tooltip). */
function buildRecPctlFormulaTooltipLines() {
  return [
    '—',
    'Formula:',
    'If Wks Away ≥ max(Wks to Lower, Wks to Lowest) → Target Pctl %.',
    'If min(Wks to Lower, Wks to Lowest) ≤ Wks Away < max(...) → Rec = Target + u × (Lower Pctl − Target), u = (Wks Away − Wks to Lower) ÷ (Wks to Lowest − Wks to Lower), u clamped to [0, 1] (linear in Wks Away between (Wks to Lower, Target) and (Wks to Lowest, Lower Pctl)).',
    'If Wks Away < min(Wks to Lower, Wks to Lowest) → Rec = Lower Pctl + v × (Lowest Pctl − Lower Pctl), v = (Wks to Lowest − Wks Away) ÷ Wks to Lowest, v clamped to [0, 1] (linear in Wks Away between (Wks to Lowest, Lower Pctl) and (0, Lowest Pctl)).',
  ];
}

/**
 * Hover text for Rec. Pctl %: tier inputs, Wks Away / LY Fin context, and formula.
 */
function buildRecPctlCellTitle(tier, effectivePct, meta = {}) {
  if (!tier) return '';
  const lines = [];
  if (tier.pacingInterpolatedLyFinal != null && Number.isFinite(tier.pacingInterpolatedLyFinal)) {
    lines.push(
      `LY final ${Math.round(tier.pacingInterpolatedLyFinal)}% — linear blend between pacing rows (Final Occ. %)`
    );
    lines.push('—');
  }
  lines.push(`Target Pctl %: ${tier.targetPricePercentile}`);
  lines.push(`Lower Pctl %: ${tier.lowerPercentileTo}`);
  lines.push(`Wks to Lower: ${formatWeeksOneDecimal(tier.weeksAhead)}`);
  lines.push(`Lowest Pctl%: ${tier.lowestPricePercentile}`);
  lines.push(`Wks to Lowest: ${formatWeeksOneDecimal(tier.weeksToStartReducing)}`);
  if (meta.wksAway != null && Number.isFinite(Number(meta.wksAway))) {
    lines.push(`Wks Away: ${formatWeeksOneDecimal(meta.wksAway)}`);
  }
  if (meta.lyFinPct != null && Number.isFinite(Number(meta.lyFinPct))) {
    const lyStr = formatPctWhole(meta.lyFinPct);
    lines.push(`LY Fin: ${lyStr != null ? lyStr : '—'}`);
  }
  if (effectivePct != null && Number.isFinite(Number(effectivePct))) {
    lines.push('—');
    lines.push(`Rec. Pctl % (this period): ${effectivePct}%`);
  }
  lines.push(...buildRecPctlFormulaTooltipLines());
  return escapeHtmlAttr(lines.join('\n')).replace(/\n/g, '&#10;');
}

/**
 * Resolve pacing tier from LY final expected occupancy: linear interpolation on **Final Occ. %**
 * between pacing rows for Target / Lower / Wks to Lower / Lowest / Wks to Lowest.
 * Rec. Pctl % in the Future table uses Wks Away vs those week thresholds (see computeRecPctlPercentileFromWksAway).
 */
function resolvePacingTarget(finalExpectedOccupancy, rows) {
  const pct = Number(finalExpectedOccupancy);
  if (!Number.isFinite(pct) || !rows || !rows.length) return null;
  const anchors = buildPacingInterpolationAnchors(rows);
  if (!anchors.length) return null;

  const withLy = (tier) => ({ ...tier, pacingInterpolatedLyFinal: pct });

  if (anchors.length === 1) {
    return withLy({ ...anchors[0] });
  }

  const tMin = anchors[0].thresholdFinalOcc;
  const tMax = anchors[anchors.length - 1].thresholdFinalOcc;

  /** No extrapolation outside [tMin, tMax]: clamp to end anchors as-is. */
  if (pct < tMin) {
    return withLy({ ...anchors[0] });
  }
  if (pct > tMax) {
    return withLy({ ...anchors[anchors.length - 1] });
  }

  for (let i = 0; i < anchors.length - 1; i++) {
    const lo = anchors[i].thresholdFinalOcc;
    const hi = anchors[i + 1].thresholdFinalOcc;
    if (pct < lo || pct > hi) continue;
    if (hi <= lo) {
      return withLy({ ...anchors[i] });
    }
    const u = (pct - lo) / (hi - lo);
    const L = anchors[i];
    const R = anchors[i + 1];
    const tp = Math.round(L.targetPricePercentile + (R.targetPricePercentile - L.targetPricePercentile) * u);
    const lowI = Math.round(
      L.lowestPricePercentile + (R.lowestPricePercentile - L.lowestPricePercentile) * u
    );
    const wksC = clampWeeksNonNeg(
      L.weeksToStartReducing + (R.weeksToStartReducing - L.weeksToStartReducing) * u
    );
    const lowerI = Math.round(L.lowerPercentileTo + (R.lowerPercentileTo - L.lowerPercentileTo) * u);
    const wksA = clampWeeksNonNeg(L.weeksAhead + (R.weeksAhead - L.weeksAhead) * u);
    const tpC = clampPct0to100(tp);
    let lowC = clampPct0to100(lowI);
    if (lowC > tpC) lowC = tpC;
    const fin = mergeInterpolatedLowerAndWeeksAhead(tpC, lowC, wksC, lowerI, wksA);
    return withLy({
      targetPricePercentile: tpC,
      lowestPricePercentile: lowC,
      weeksToStartReducing: wksC,
      lowerPercentileTo: fin.lowerPercentileTo,
      weeksAhead: fin.weeksAhead,
      thresholdFinalOcc: Math.round(pct),
      tierClass: 'pacing-interpolated',
    });
  }

  return withLy({ ...anchors[anchors.length - 1] });
}

/**
 * Rec. Pctl % for a future period: **Wks Away** (as-of → period start) vs Wks to Lower / Wks to Lowest.
 * Beyond the farther week threshold → Target. Between the two thresholds → linear in Wks Away from
 * (Wks to Lower, Target Pctl) to (Wks to Lowest, Lower Pctl): u = (Wks Away − Wks to Lower) / (Wks to Lowest − Wks to Lower).
 * Below min(Wks to Lower, Wks to Lowest) → linear in Wks Away from (Wks to Lowest, Lower Pctl) to (0, Lowest Pctl).
 */
function computeRecPctlPercentileFromWksAway(tier, wksAway, _lyFinPct) {
  if (!tier || wksAway == null || !Number.isFinite(Number(wksAway))) return null;
  const tgt = clampPct0to100(tier.targetPricePercentile);
  const lower = clampPct0to100(tier.lowerPercentileTo);
  const floor = clampPct0to100(tier.lowestPricePercentile);
  const wL = clampWeeksNonNeg(tier.weeksAhead);
  const wF = clampWeeksNonNeg(tier.weeksToStartReducing);
  const hi = Math.max(wL, wF);
  const lo = Math.min(wL, wF);

  if (wksAway >= hi) return Math.round(tgt);

  if (wksAway < lo) {
    if (wF <= 1e-9) return Math.round(floor);
    const v = Math.max(0, Math.min(1, (wF - wksAway) / wF));
    return Math.round(clampPct0to100(lower + v * (floor - lower)));
  }

  const denom = wF - wL;
  if (Math.abs(denom) <= 1e-9) return Math.round(tgt);
  let u = (wksAway - wL) / denom;
  u = Math.max(0, Math.min(1, u));
  return Math.round(clampPct0to100(tgt + u * (lower - tgt)));
}

/** Mini number line: T/L/F = Target, Lower, Lowest; arrow = resolved Rec. Pctl % (when present). */
function buildRecPctlNumberLineHtml(tier, recPctResolved) {
  if (!tier) return '<span class="rec-pctl-line-empty" aria-hidden="true">—</span>';
  const tgt = clampPct0to100(tier.targetPricePercentile);
  const lower = clampPct0to100(tier.lowerPercentileTo);
  const floor = clampPct0to100(tier.lowestPricePercentile);
  const hasRec = recPctResolved != null && Number.isFinite(Number(recPctResolved));
  const rec = hasRec ? clampPct0to100(recPctResolved) : null;

  const VB_W = 200;
  const VB_H = 60;
  const PAD = 10;
  const BAR_Y = 26;
  const yMark = 34;
  const yWks = 44;
  const yNum = 56;
  const innerW = VB_W - 2 * PAD;
  /** Map percentile to x: higher % on the left, lower % on the right (mirrored 0–100). */
  const toX = (p) => VB_W - PAD - (clampPct0to100(p) / 100) * innerW;

  const xT = toX(tgt);
  const xL = toX(lower);
  const xF = toX(floor);
  const wToLowerStr = formatWeeksOneDecimal(clampWeeksNonNeg(tier.weeksAhead));
  const wToLowestStr = formatWeeksOneDecimal(clampWeeksNonNeg(tier.weeksToStartReducing));

  const anchors = [
    { pct: tgt, mark: 'T', tip: `Target ${tgt}%` },
    { pct: lower, mark: 'L', tip: `Lower ${lower}%` },
    { pct: floor, mark: 'F', tip: `Lowest ${floor}%` },
  ];

  const tipSuffix = hasRec ? ` · Arrow: Rec. ${rec}%` : '';
  let html = `<div class="rec-pctl-line-wrap" title="T = Target · L = Lower · F = Lowest · scale: high % left, low % right${tipSuffix}">`;
  html += `<svg class="rec-pctl-line" viewBox="0 0 ${VB_W} ${VB_H}" xmlns="http://www.w3.org/2000/svg" focusable="false" aria-hidden="true">`;
  html += `<line class="rec-pctl-line__axis" x1="${PAD}" y1="${BAR_Y}" x2="${VB_W - PAD}" y2="${BAR_Y}"/>`;

  for (const a of anchors) {
    const x = toX(a.pct);
    html += `<g class="rec-pctl-line__anchor">`;
    html += `<title>${a.tip.replace(/&/g, '&amp;')}</title>`;
    html += `<line class="rec-pctl-line__tick rec-pctl-line__tick--${a.mark === 'T' ? 'target' : a.mark === 'L' ? 'lower' : 'lowest'}" x1="${x}" y1="${BAR_Y - 8}" x2="${x}" y2="${BAR_Y + 8}"/>`;
    html += `<text class="rec-pctl-line__mark" x="${x}" y="${yMark}" text-anchor="middle" dominant-baseline="middle">${a.mark}</text>`;
    html += `<text class="rec-pctl-line__num" x="${x}" y="${yNum}" text-anchor="middle" dominant-baseline="middle">${a.pct}</text>`;
    html += `</g>`;
  }

  if (Math.abs(xT - xL) > 2) {
    const xMid = (xT + xL) / 2;
    html += `<g class="rec-pctl-line__wks-gap">`;
    html += `<title>Wks to Lower: ${wToLowerStr} (between Target and Lower)</title>`;
    html += `<text class="rec-pctl-line__wks-val" x="${xMid}" y="${yWks}" text-anchor="middle">${wToLowerStr}</text>`;
    html += `</g>`;
  }
  if (Math.abs(xL - xF) > 2) {
    const xMid = (xL + xF) / 2;
    html += `<g class="rec-pctl-line__wks-gap">`;
    html += `<title>Wks to Lowest: ${wToLowestStr} (between Lower and Lowest)</title>`;
    html += `<text class="rec-pctl-line__wks-val" x="${xMid}" y="${yWks}" text-anchor="middle">${wToLowestStr}</text>`;
    html += `</g>`;
  }

  if (hasRec) {
    const rx = toX(rec);
    html += `<g class="rec-pctl-line__rec">`;
    html += `<title>Rec. Pctl ${rec}%</title>`;
    html += `<line class="rec-pctl-line__rec-stem" x1="${rx}" y1="4" x2="${rx}" y2="12"/>`;
    html += `<polygon class="rec-pctl-line__rec-head" points="${rx},${BAR_Y - 1} ${rx - 6},12 ${rx + 6},12"/>`;
    html += `</g>`;
  }

  html += `</svg></div>`;
  return html;
}

/** Color hint on recommended pricing percentile (Rec. Pctl %) column by target level */
function pacingPercentileHeatClass(targetPct) {
  const n = Number(targetPct);
  if (!Number.isFinite(n)) return '';
  if (n >= 75) return 'tier-high';
  if (n >= 50) return 'tier-mid';
  return 'tier-low';
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

/**
 * Nightly $ at percentile `pct` (0–100) on the same P25–P90 ladder as Fin % (inverse of deriveFinalPricePercentile).
 */
function derivePriceFromPercentileLadder(pct, p25, p50, p75, p90) {
  const p = Number(pct);
  const q25 = Number(p25);
  const q50 = Number(p50);
  const q75 = Number(p75);
  const q90 = Number(p90);
  if (!Number.isFinite(p) || ![q25, q50, q75, q90].every(Number.isFinite) || q25 <= 0) return null;
  if (q25 > q50 || q50 > q75 || q75 > q90) return null;
  const eps = 1e-6;
  let price;
  if (p <= 50) {
    const d = q50 - q25;
    if (d < eps) price = q25;
    else price = q25 + ((p - 25) / 25) * d;
  } else if (p <= 75) {
    const d = q75 - q50;
    if (d < eps) price = q50;
    else price = q50 + ((p - 50) / 25) * d;
  } else if (p <= 90) {
    const d = q90 - q75;
    if (d < eps) price = q75;
    else price = q75 + ((p - 75) / 15) * d;
  } else {
    const d = q90 - q75;
    if (d < eps) price = q90;
    else price = q90 + ((p - 90) / 15) * d;
  }
  if (!Number.isFinite(price)) return null;
  return Math.round(Math.max(0, price));
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
    // Synthetic "last year same week" index (0–99) — used for price anchors only below.
    const lastYearSameWeek = (52 + w * 2 + (weekStart.getMonth() % 3) * 5) % 100;
    // Final expected occupancy: must NOT be mostly 91–92%. Old formula was
    // clamp(lastYearSameWeek + 35, 28, 92) which hit the 92% cap whenever lastYearSameWeek ≥ 58.
    // Spread ~34–90% in a stable, week-varying way (still fake sample data until PriceLabs).
    const spreadSeed = (w * 19 + weekStart.getMonth() * 37 + (weekStart.getFullYear() % 4) * 11) % 61;
    const finalExpected = Math.min(90, Math.max(34, 34 + spreadSeed));
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

/** Calendar YYYY-MM-DD for Future table “As Of”: Hospitable pull date if set, else today (local). */
function getFuturePacingAsOfYmd() {
  const iso = localStorage.getItem(STORAGE_KEYS.hospitableAsOf);
  if (iso) {
    const d = new Date(iso);
    if (!isNaN(d.getTime())) return toYMDLocal(d);
  }
  const t = new Date();
  t.setHours(0, 0, 0, 0);
  return toYMDLocal(t);
}

/** Fractional weeks from `fromYmd` to `toYmd` (local calendar): period start minus as-of. */
function weeksFromYmdToYmd(fromYmd, toYmd) {
  const a = parseYMDToLocalNoon(fromYmd);
  const b = parseYMDToLocalNoon(toYmd);
  if (!a || !b) return null;
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  return (b.getTime() - a.getTime()) / msPerWeek;
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
        text: `Week of ${fp.periodStart}: LY final ${fp.finalExpected}% → ~${tier.effectivePacingPercentile ?? tier.targetPricePercentile}th price pctl${
          tier.weeksToStartReducing > 0
            ? tier.weeksAhead > 0 && tier.lowerPercentileTo > tier.lowestPricePercentile
              ? ` (Target ${tier.targetPricePercentile}% → Lower ${tier.lowerPercentileTo}% @ Wks to Lower ${formatWeeksOneDecimal(tier.weeksAhead)} → Lowest ${tier.lowestPricePercentile}% / Wks to Lowest ${formatWeeksOneDecimal(tier.weeksToStartReducing)})`
              : ` (Target ${tier.targetPricePercentile}% → Lowest ${tier.lowestPricePercentile}% over Wks to Lowest ${formatWeeksOneDecimal(tier.weeksToStartReducing)})`
            : ` (Target ${tier.targetPricePercentile}%)`
        }. ${fp.remainingPotential > 30 ? 'Good remaining demand—consider holding rate.' : 'Limited remaining demand—consider promotions.'}`,
      });
    });
  }
  if (unbooked.length > 14) {
    recs.push({
      type: 'action',
      text: `${unbooked.length} unbooked nights in the selected window (${unbookedHorizonDays} days). Use the Future Unbooked Periods table to set price percentiles by week.`,
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
    recs.push({ type: 'ok', text: 'Pacing and occupancy look reasonable. Keep monitoring unbooked windows and pacing rules.' });
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

/** For HTML attribute values (e.g. title="..."). */
function escapeHtmlAttr(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;');
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
  const mapPlaceholder = document.getElementById('marketColumnMapPlaceholder');
  if (!statusEl || !mapWrap || !mapBody) return;

  const statusBase = 'market-csv-status market-csv-status--compact';

  const bundle = loadMarketPacingBundle();
  if (!bundle || !bundle.rows.length) {
    statusEl.textContent = '';
    statusEl.removeAttribute('title');
    statusEl.className = statusBase;
    mapWrap.hidden = true;
    mapBody.innerHTML = '';
    if (mapPlaceholder) mapPlaceholder.hidden = false;
    if (extraEl) {
      extraEl.hidden = true;
      extraEl.textContent = '';
    }
    if (clearBtn) clearBtn.hidden = true;
    if (dropzone) dropzone.classList.remove('csv-dropzone--has-file');
    return;
  }

  const meta = bundle.meta || {};
  statusEl.className = `${statusBase} market-csv-status--ok`;
  let shortName = meta.fileName || 'Market CSV';
  if (shortName.length > 14) shortName = `${shortName.slice(0, 12)}…`;
  statusEl.textContent = `${shortName} · ${bundle.rows.length} rows`;
  statusEl.title = `Future unbooked table uses this file instead of sample market data. Full name: ${meta.fileName || 'saved'}`;

  const columnMap = Array.isArray(meta.columnMap) ? meta.columnMap : [];
  mapBody.innerHTML = columnMap
    .map(
      (row) =>
        `<tr><td>${escapeHtml(row.field)}</td><td><code>${escapeHtml(row.header)}</code></td></tr>`
    )
    .join('');
  mapWrap.hidden = false;
  if (mapPlaceholder) mapPlaceholder.hidden = true;

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

// --- Pacing rows load/save (Final occupancy % threshold → Target price percentile) ---
function normalizePacingRowsFromStorage(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return null;
  const first = parsed[0];
  if (first && typeof first === 'object' && 'finalOccupancyPct' in first && 'targetPricePercentile' in first) {
    return parsed.map((r, i) => {
      const tp = clampPct0to100(r.targetPricePercentile);
      let low =
        r.lowestPricePercentile != null && r.lowestPricePercentile !== ''
          ? clampPct0to100(r.lowestPricePercentile)
          : Math.max(0, tp - 30);
      if (low > tp) low = tp;
      const wks = r.weeksToStartReducing != null ? clampWeeksNonNeg(r.weeksToStartReducing) : 3;
      const finN = finalizeLowerAndWeeksAhead(tp, low, wks, r.lowerPercentileTo, r.weeksAhead);
      const weeksAheadN = coalesceStoredWeeksAhead(r, finN.weeksAhead);
      return {
        id: String(r.id != null && String(r.id).trim() !== '' ? r.id : `row-${i}`),
        finalOccupancyPct: clampPct0to100(r.finalOccupancyPct),
        targetPricePercentile: tp,
        lowestPricePercentile: low,
        weeksToStartReducing: wks,
        lowerPercentileTo: finN.lowerPercentileTo,
        weeksAhead: weeksAheadN,
      };
    });
  }
  return null;
}

function loadPacingTiers() {
  const raw = localStorage.getItem(STORAGE_KEYS.pacingRows);
  if (!raw) return JSON.parse(JSON.stringify(DEFAULT_PACING_ROWS));
  try {
    const parsed = JSON.parse(raw);
    const norm = normalizePacingRowsFromStorage(parsed);
    return norm && norm.length ? norm : JSON.parse(JSON.stringify(DEFAULT_PACING_ROWS));
  } catch {
    return JSON.parse(JSON.stringify(DEFAULT_PACING_ROWS));
  }
}

function savePacingTiers(rows) {
  localStorage.setItem(STORAGE_KEYS.pacingRows, JSON.stringify(rows));
}

/** Defaults used by Reset (and empty-table fallback): saved template or built-in factory defaults. */
function loadUserPacingDefaults() {
  const raw = localStorage.getItem(STORAGE_KEYS.pacingRowsDefaults);
  if (!raw) return JSON.parse(JSON.stringify(DEFAULT_PACING_ROWS));
  try {
    const parsed = JSON.parse(raw);
    const norm = normalizePacingRowsFromStorage(parsed);
    if (norm && norm.length) return norm;
  } catch {
    /* ignore */
  }
  return JSON.parse(JSON.stringify(DEFAULT_PACING_ROWS));
}

function saveUserPacingDefaults(rows) {
  localStorage.setItem(STORAGE_KEYS.pacingRowsDefaults, JSON.stringify(rows));
}

function buildFuturePeriodsWithMarket(reservations, pacingTiers, horizonDays) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const asOfYmd = getFuturePacingAsOfYmd();
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
    const tierBase =
      finalExpected != null && pacingTiers && pacingTiers.length
        ? resolvePacingTarget(finalExpected, pacingTiers)
        : null;
    const weeksAwayFromAsOf = weeksFromYmdToYmd(asOfYmd, p.periodStart);
    const effectivePacingPercentile =
      tierBase != null && weeksAwayFromAsOf != null && Number.isFinite(weeksAwayFromAsOf)
        ? computeRecPctlPercentileFromWksAway(tierBase, weeksAwayFromAsOf, finalExpected)
        : null;
    const tier =
      tierBase != null
        ? { ...tierBase, effectivePacingPercentile }
        : null;
    const finalPrice = market ? market.finalPrice : null;
    const priceP25 = market ? market.priceP25 : null;
    const priceP50 = market ? market.priceP50 : null;
    const priceP75 = market ? market.priceP75 : null;
    const priceP90 = market ? market.priceP90 : null;
    const finalPricePercentile = deriveFinalPricePercentile(finalPrice, priceP25, priceP50, priceP75, priceP90);
    out.push({
      ...p,
      weeksAwayFromAsOf,
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
    const name = ctx.dataset.label || '';
    if (key === 'recPacingPercentile') {
      const fp = fps[ctx.dataIndex];
      const t = fp?.tier;
      const raw = t ? t.effectivePacingPercentile ?? t.targetPricePercentile : null;
      if (raw == null || !Number.isFinite(Number(raw))) return `${name}: —`;
      return `${name}: ${Math.round(Number(raw))}%`;
    }
    const raw = fps[ctx.dataIndex][key];
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
        {
          type: 'line',
          label: 'Rec. Pctl %',
          yAxisID: 'y1',
          order: 101,
          data: fps.map((fp) => {
            const t = fp.tier;
            if (!t) return null;
            const v = t.effectivePacingPercentile ?? t.targetPricePercentile;
            return v != null && Number.isFinite(Number(v)) ? Number(v) : null;
          }),
          _fieldKey: 'recPacingPercentile',
          borderColor: 'rgba(63, 210, 95, 0.95)',
          backgroundColor: 'rgba(63, 210, 95, 0.06)',
          borderWidth: 2.25,
          tension: 0.22,
          pointRadius: 3.5,
          pointHoverRadius: 5.5,
          pointBackgroundColor: 'rgba(90, 235, 120, 1)',
          pointBorderColor: 'rgba(20, 24, 30, 1)',
          pointBorderWidth: 2,
          spanGaps: false,
          fill: false,
          datalabels: {
            display(ctx) {
              const fp = fps[ctx.dataIndex];
              const t = fp?.tier;
              const v = t ? t.effectivePacingPercentile ?? t.targetPricePercentile : null;
              return v != null && Number.isFinite(Number(v));
            },
            formatter(_, ctx) {
              const fp = fps[ctx.dataIndex];
              const t = fp?.tier;
              const v = t ? t.effectivePacingPercentile ?? t.targetPricePercentile : null;
              return `${Math.round(Number(v))}%`;
            },
            color: 'rgba(140, 255, 170, 1)',
            backgroundColor: 'rgba(18, 22, 28, 0.94)',
            borderColor: 'rgba(63, 210, 95, 0.85)',
            borderWidth: 1,
            borderRadius: 6,
            padding: { top: 4, right: 6, bottom: 4, left: 6 },
            font: { weight: '700', size: 10, family: "'JetBrains Mono', monospace" },
            anchor: 'center',
            align: 'bottom',
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
            text: 'Fin % / Rec. Pctl %',
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

/** Rec−Fin cell style: stronger green/red as |Rec−Fin| grows (mix toward --success / --danger). */
function formatRecMinusFinDiffStyle(d) {
  const n = Math.round(Number(d));
  if (!Number.isFinite(n)) return '';
  const cap = 30;
  const t = Math.min(1, Math.abs(n) / cap);
  const mix = Math.round(12 + t * 88);
  if (n === 0) return 'color: var(--text-muted); font-weight: 600;';
  if (n > 0) {
    return `color: color-mix(in srgb, var(--success) ${mix}%, var(--text-muted)); font-weight: 600;`;
  }
  return `color: color-mix(in srgb, var(--danger) ${mix}%, var(--text-muted)); font-weight: 600;`;
}

function formatMoneyWhole(v) {
  if (v == null || !Number.isFinite(Number(v))) return null;
  return `$${Math.round(Number(v)).toLocaleString()}`;
}

function setTextById(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

const FUTURE_PACING_MONTH_SHORT = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

function parseFuturePacingYmdParts(ymd) {
  if (!ymd || typeof ymd !== 'string' || ymd.length < 10) return null;
  const y = Number(ymd.slice(0, 4));
  const m = Number(ymd.slice(5, 7));
  const d = Number(ymd.slice(8, 10));
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  return { y, m, d };
}

/** Table: abbreviated month(s), e.g. Mar or Mar–Apr when the range spans months. */
function formatFuturePacingMonthAbbrevRange(periodStart, periodEnd) {
  const a = parseFuturePacingYmdParts(periodStart);
  const b = parseFuturePacingYmdParts(periodEnd);
  if (!a || !b) return '—';
  if (a.y === b.y && a.m === b.m) return FUTURE_PACING_MONTH_SHORT[a.m - 1];
  return `${FUTURE_PACING_MONTH_SHORT[a.m - 1]}–${FUTURE_PACING_MONTH_SHORT[b.m - 1]}`;
}

/** Table: day-of-month only, e.g. 12 or 12→15. */
function formatFuturePacingPeriodDaysOnly(periodStart, periodEnd) {
  const a = parseFuturePacingYmdParts(periodStart);
  const b = parseFuturePacingYmdParts(periodEnd);
  if (!a || !b) return '—';
  if (periodStart === periodEnd) return String(a.d);
  return `${a.d}→${b.d}`;
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
  const raw = localStorage.getItem(STORAGE_KEYS.reservations);
  return {
    revenueTarget: target ? parseFloat(target) : null,
    reservations: raw ? parseReservations(raw) : null,
    pacingTiers: loadPacingTiers(),
  };
}

function saveStored(revenueTarget, reservations) {
  if (revenueTarget != null) localStorage.setItem(STORAGE_KEYS.revenueTarget, String(revenueTarget));
  if (reservations && reservations.length)
    localStorage.setItem(STORAGE_KEYS.reservations, JSON.stringify(reservations));
}

/** Record that booking data was just loaded from Hospitable (local device time). */
function setHospitableAsOfToNow() {
  localStorage.setItem(STORAGE_KEYS.hospitableAsOf, new Date().toISOString());
  refreshHospitableAsOfBanner();
}

/** Booking data is no longer from Hospitable (sample / CSV / etc.). */
function clearHospitableAsOf() {
  localStorage.removeItem(STORAGE_KEYS.hospitableAsOf);
  refreshHospitableAsOfBanner();
}

function refreshHospitableAsOfBanner() {
  const wrap = document.getElementById('hospitableAsOfBanner');
  const timeEl = document.getElementById('hospitableAsOfTime');
  if (!wrap || !timeEl) return;
  const iso = localStorage.getItem(STORAGE_KEYS.hospitableAsOf);
  if (!iso) {
    wrap.hidden = true;
    timeEl.removeAttribute('datetime');
    timeEl.textContent = '';
    return;
  }
  const d = new Date(iso);
  if (isNaN(d.getTime())) {
    wrap.hidden = true;
    return;
  }
  timeEl.dateTime = iso;
  timeEl.textContent = d.toLocaleString(undefined, {
    dateStyle: 'long',
    timeStyle: 'short',
  });
  wrap.hidden = false;
}

function syncFuturePacingTableDetailsUI() {
  const table = document.getElementById('futurePacingTable');
  const btn = document.getElementById('futurePacingDetailsToggle');
  if (!table || !btn) return;
  const visible = table.classList.contains('future-pacing-table--details-visible');
  btn.setAttribute('aria-pressed', visible ? 'true' : 'false');
  btn.textContent = visible ? 'Hide Details' : 'Show Details';
}

function syncFuturePacingTableToolbarVisibility() {
  const table = document.getElementById('futurePacingTable');
  const toolbar = document.getElementById('futurePacingTableToolbar');
  if (!toolbar || !table) return;
  toolbar.hidden = table.hidden;
}

/** Apply saved “Show details” state to the future pacing table. */
function applyFuturePacingTableDetailsFromStorage() {
  const table = document.getElementById('futurePacingTable');
  if (!table) return;
  if (localStorage.getItem(STORAGE_KEYS.futurePacingTableDetailsVisible) === '1') {
    table.classList.add('future-pacing-table--details-visible');
  } else {
    table.classList.remove('future-pacing-table--details-visible');
  }
  syncFuturePacingTableDetailsUI();
}

function render(reservations, periodKey, revenueTarget) {
  updateUnbookedSectionHeading();
  applyFuturePacingTableDetailsFromStorage();
  const horizonDays = getUnbookedHorizonDays();
  const pacingTiers = loadPacingTiers();
  if (!reservations || !reservations.length) {
    setTextById('revenue', '—');
    setTextById('adr', '—');
    setTextById('occupancy', '—');
    setTextById('revpar', '—');
    setTextById('nightsBooked', '—');
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
    destroyRevenueMetricsChart();
    const revChartWrap = document.getElementById('revenueMetricsChartWrap');
    if (revChartWrap) revChartWrap.hidden = true;
    if (chartsWrap) chartsWrap.hidden = true;
    if (emptyHint) emptyHint.hidden = false;
    if (tableEl) tableEl.hidden = true;
    syncFuturePacingTableToolbarVisibility();
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

  setTextById('revenue', `$${Math.round(metrics.revenue).toLocaleString()}`);
  setTextById('adr', metrics.adr > 0 ? `$${Math.round(metrics.adr).toLocaleString()}` : '—');
  setTextById('occupancy', `${metrics.occupancy.toFixed(1)}%`);
  setTextById('revpar', metrics.revpar > 0 ? `$${metrics.revpar.toFixed(0)}` : '—');
  setTextById('nightsBooked', String(metrics.nights));

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
      const futureTableAsOfYmd = getFuturePacingAsOfYmd();
      tbody.innerHTML = futurePeriodsWithMarket
        .map((fp) => {
          const currentStr = formatPctWhole(fp.currentMarketOccupancy) ?? '—';
          const lytStr = formatPctWhole(fp.lastYearTodayOccupancy) ?? '—';
          const finalStr = formatPctWhole(fp.finalExpectedOccupancy) ?? '—';
          const remainingStr = formatPctWhole(fp.remainingPotential) ?? '—';
          const remainingLyStr = formatPctWhole(fp.remainingPotentialLY) ?? '—';
          const wksAwayStr =
            fp.weeksAwayFromAsOf != null && Number.isFinite(fp.weeksAwayFromAsOf)
              ? (Math.round(fp.weeksAwayFromAsOf * 10) / 10).toFixed(1)
              : '—';
          const finStr = formatMoneyWhole(fp.finalPrice) ?? '—';
          const finPctStr = formatPctWhole(fp.finalPricePercentile) ?? '—';
          const p25Str = formatMoneyWhole(fp.priceP25) ?? '—';
          const p50Str = formatMoneyWhole(fp.priceP50) ?? '—';
          const p75Str = formatMoneyWhole(fp.priceP75) ?? '—';
          const p90Str = formatMoneyWhole(fp.priceP90) ?? '—';
          const tier = fp.tier;
          const eff = tier ? tier.effectivePacingPercentile ?? tier.targetPricePercentile : null;
          const pctStr = tier && eff != null ? `${eff}%` : '—';
          const tierTitleAttr = tier
            ? ` title="${buildRecPctlCellTitle(tier, eff, {
                wksAway: fp.weeksAwayFromAsOf,
                lyFinPct: fp.finalExpectedOccupancy,
              })}"`
            : '';
          const heat = tier && eff != null ? pacingPercentileHeatClass(eff) : '';
          const tierClass = tier ? `pct-cell ${heat}` : 'pct-cell';
          const recResolved =
            tier &&
            tier.effectivePacingPercentile != null &&
            Number.isFinite(Number(tier.effectivePacingPercentile))
              ? tier.effectivePacingPercentile
              : null;
          const recLineHtml = buildRecPctlNumberLineHtml(tier, recResolved);
          const recPriceVal =
            recResolved != null
              ? derivePriceFromPercentileLadder(recResolved, fp.priceP25, fp.priceP50, fp.priceP75, fp.priceP90)
              : null;
          const recDollarStr = formatMoneyWhole(recPriceVal) ?? '—';
          const recDollarTitle =
            recResolved != null
              ? escapeHtmlAttr(
                  recPriceVal != null
                    ? `Rec. Pctl % ${recResolved}%; same P25–P90 ladder as Fin %`
                    : `Rec. Pctl % ${recResolved}%; P25–P90 ladder unavailable for this period`
                )
              : '';
          const recDollarTitleAttr = recDollarTitle ? ` title="${recDollarTitle}"` : '';
          const finPctNum =
            fp.finalPricePercentile != null && Number.isFinite(Number(fp.finalPricePercentile))
              ? Number(fp.finalPricePercentile)
              : null;
          let recMinusFinStr = '—';
          let recMinusFinTitleAttr = '';
          let recMinusFinStyleAttr = '';
          if (tier != null && eff != null && finPctNum != null) {
            const d = Math.round(Number(eff) - finPctNum);
            recMinusFinStr = `${d > 0 ? '+' : ''}${d}%`;
            recMinusFinTitleAttr = ` title="${escapeHtmlAttr(`Rec. Pctl % (${eff}) − Fin % (${finPctNum})`)}"`;
            const diffCss = formatRecMinusFinDiffStyle(d);
            if (diffCss) recMinusFinStyleAttr = ` style="${escapeHtmlAttr(diffCss)}"`;
          }
          const dowRange = formatDowRange(fp.periodStart, fp.periodEnd);
          const monthAbbrev = formatFuturePacingMonthAbbrevRange(fp.periodStart, fp.periodEnd);
          const periodDays = formatFuturePacingPeriodDaysOnly(fp.periodStart, fp.periodEnd);
          const periodTitle = `${fp.periodStart} – ${fp.periodEnd}`;
          return `<tr>
            <td class="future-pacing-month-cell" title="${periodTitle}">${monthAbbrev}</td>
            <td class="future-pacing-period-cell" title="${periodTitle}">${periodDays}</td>
            <td class="future-pacing-dow-cell" title="${dowRange}">${dowRange}</td>
            <td class="future-pacing-num-cell">${fp.unbookedNights}</td>
            <td class="pct-cell future-pacing-detail-col">${currentStr}</td>
            <td class="pct-cell future-pacing-detail-col">${lytStr}</td>
            <td class="pct-cell">${remainingStr}</td>
            <td class="pct-cell future-pacing-detail-col">${remainingLyStr}</td>
            <td class="pct-cell">${finalStr}</td>
            <td class="future-pacing-num-cell" title="As Of ${futureTableAsOfYmd} → period start ${fp.periodStart}">${wksAwayStr}</td>
            <td class="money-cell">${finStr}</td>
            <td class="pct-cell">${finPctStr}</td>
            <td class="pct-cell future-pacing-rec-minus-fin"${recMinusFinTitleAttr}${recMinusFinStyleAttr}>${recMinusFinStr}</td>
            <td class="money-cell ${heat}"${recDollarTitleAttr}>${recDollarStr}</td>
            <td class="${tierClass}"${tierTitleAttr}>${pctStr}</td>
            <td class="future-pacing-rec-line-cell">${recLineHtml}</td>
            <td class="money-cell future-pacing-detail-col">${p25Str}</td>
            <td class="money-cell future-pacing-detail-col">${p50Str}</td>
            <td class="money-cell future-pacing-detail-col">${p75Str}</td>
            <td class="money-cell future-pacing-detail-col">${p90Str}</td>
          </tr>`;
        })
        .join('');
    }
  }
  syncFuturePacingTableToolbarVisibility();
  renderFuturePacingCharts(futurePeriodsWithMarket);
  renderRevenueMetricsChart(reservations);
  renderMarketCsvPanel();
}

function init() {
  const stored = loadStored();
  let reservations = stored.reservations;

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
      render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
    });
  }
  updateUnbookedSectionHeading();
  refreshHospitableAsOfBanner();

  document.getElementById('futurePacingDetailsToggle')?.addEventListener('click', () => {
    const table = document.getElementById('futurePacingTable');
    if (!table || table.hidden) return;
    table.classList.toggle('future-pacing-table--details-visible');
    if (table.classList.contains('future-pacing-table--details-visible')) {
      localStorage.setItem(STORAGE_KEYS.futurePacingTableDetailsVisible, '1');
    } else {
      localStorage.removeItem(STORAGE_KEYS.futurePacingTableDetailsVisible);
    }
    syncFuturePacingTableDetailsUI();
  });

  const marketCsvInput = document.getElementById('marketCsvInput');
  const marketDropzone = document.getElementById('marketCsvDropzone');
  const marketCsvStatus = document.getElementById('marketCsvStatus');

  const showMarketDropError = (msg) => {
    if (!marketCsvStatus) return;
    marketCsvStatus.className = 'market-csv-status market-csv-status--compact market-csv-status--err';
    marketCsvStatus.textContent = msg;
    marketCsvStatus.title = msg;
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
      render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
      return;
    }

    if (shape.kind === 'reservations') {
      const parsed = parseReservationsFromTableRows(trimmed);
      if (!parsed.length) {
        showMarketDropError('No valid reservation rows (need check_in and check_out columns).');
        return;
      }
      reservations = parsed;
      clearHospitableAsOf();
      saveStored(parseFloat(targetInput.value) || null, reservations);
      render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
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

  const marketColumnMapDialog = document.getElementById('marketColumnMapDialog');
  const openMarketColumnMapBtn = document.getElementById('openMarketColumnMap');
  const closeMarketColumnMapBtn = document.getElementById('closeMarketColumnMap');
  if (openMarketColumnMapBtn && marketColumnMapDialog) {
    openMarketColumnMapBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      marketColumnMapDialog.showModal();
    });
  }
  if (closeMarketColumnMapBtn && marketColumnMapDialog) {
    closeMarketColumnMapBtn.addEventListener('click', () => marketColumnMapDialog.close());
  }
  if (marketColumnMapDialog) {
    marketColumnMapDialog.addEventListener('click', (e) => {
      if (e.target === marketColumnMapDialog) marketColumnMapDialog.close();
    });
  }

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
      render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
    });
  }

  document.getElementById('saveTarget').addEventListener('click', () => {
    const v = parseFloat(targetInput.value);
    if (!isNaN(v) && v >= 0) {
      localStorage.setItem(STORAGE_KEYS.revenueTarget, String(v));
      render(reservations, PACING_PERIOD_KEY, v);
    }
  });

  document.getElementById('useSample').addEventListener('click', () => {
    reservations = getSampleReservations();
    clearHospitableAsOf();
    saveStored(parseFloat(targetInput.value) || null, reservations);
    render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
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
        setHospitableAsOfToNow();
        saveStored(parseFloat(targetInput.value) || null, reservations);
        render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
      } catch (err) {
        alert('Fetch from Hospitable failed. Is the backend running (npm start)? ' + (err.message || String(err)));
      } finally {
        fetchHospitableBtn.textContent = originalText;
        fetchHospitableBtn.disabled = false;
      }
    });
  }

  document.getElementById('openSettings').addEventListener('click', () => {
    openPacingSettingsModal();
  });
  document.getElementById('closeSettings').addEventListener('click', () => {
    document.getElementById('settingsOverlay').hidden = true;
  });
  document.getElementById('savePacingSettings').addEventListener('click', () => {
    savePacingSettingsFromForm();
    document.getElementById('settingsOverlay').hidden = true;
    render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
  });
  document.getElementById('resetPacingSettings').addEventListener('click', () => {
    const defs = JSON.parse(JSON.stringify(loadUserPacingDefaults()));
    savePacingTiers(defs);
    populateTiersTable(defs);
    render(reservations, PACING_PERIOD_KEY, parseFloat(targetInput.value) || null);
  });

  document.getElementById('savePacingDefaults')?.addEventListener('click', () => {
    savePacingDefaultsFromForm();
    const btn = document.getElementById('savePacingDefaults');
    if (btn) {
      const prev = btn.textContent;
      btn.textContent = 'Saved as defaults';
      btn.disabled = true;
      setTimeout(() => {
        btn.textContent = prev;
        btn.disabled = false;
      }, 1600);
    }
  });

  document.getElementById('addPacingRow')?.addEventListener('click', () => {
    appendPacingRowToTable({
      id: newPacingRowId(),
      finalOccupancyPct: 0,
      targetPricePercentile: 50,
      lowestPricePercentile: 35,
      weeksToStartReducing: 3,
      lowerPercentileTo: 42,
      weeksAhead: 1,
    });
  });

  document.getElementById('tiersTableBody')?.addEventListener('click', (e) => {
    const btn = e.target.closest('.btn-row-delete');
    if (!btn) return;
    const tr = btn.closest('tr');
    const tbody = document.getElementById('tiersTableBody');
    if (!tr || !tbody || tbody.querySelectorAll('tr').length <= 1) return;
    tr.remove();
  });
  initPacingTierRowDragDrop();
  document.getElementById('settingsOverlay').addEventListener('click', (e) => {
    if (e.target.id === 'settingsOverlay') e.target.hidden = true;
  });

  render(
    reservations,
    PACING_PERIOD_KEY,
    stored.revenueTarget != null ? stored.revenueTarget : parseFloat(targetInput.value) || null
  );
}

/**
 * Insert-before target for pacing row drag (y = pointer clientY). Excludes the row being dragged.
 */
function getTierDragAfterElement(tbody, dragRow, y) {
  const els = [...tbody.querySelectorAll('tr')].filter((tr) => tr !== dragRow);
  return els.reduce(
    (closest, child) => {
      const box = child.getBoundingClientRect();
      const offset = y - box.top - box.height / 2;
      if (offset < 0 && offset > closest.offset) {
        return { offset, element: child };
      }
      return closest;
    },
    { offset: Number.NEGATIVE_INFINITY, element: null }
  ).element;
}

/** HTML5 drag reorder for pacing settings rows (grip handle only). */
function initPacingTierRowDragDrop() {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody || tbody.dataset.tierDragInit) return;
  tbody.dataset.tierDragInit = '1';
  let dragRow = null;

  tbody.addEventListener('dragstart', (e) => {
    const handle = e.target.closest('.tier-drag-handle');
    if (!handle) return;
    dragRow = handle.closest('tr');
    if (!dragRow) return;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', dragRow.dataset.rowId || 'row');
    dragRow.classList.add('tier-row--dragging');
  });

  tbody.addEventListener('dragend', () => {
    if (dragRow) dragRow.classList.remove('tier-row--dragging');
    dragRow = null;
  });

  tbody.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (!dragRow) return;
    const after = getTierDragAfterElement(tbody, dragRow, e.clientY);
    if (after == null) {
      if (dragRow.nextSibling !== null) tbody.appendChild(dragRow);
    } else if (dragRow.nextSibling !== after) {
      tbody.insertBefore(dragRow, after);
    }
  });

  tbody.addEventListener('drop', (e) => e.preventDefault());
}

function openPacingSettingsModal() {
  const tiers = loadPacingTiers();
  populateTiersTable(tiers);
  document.getElementById('settingsOverlay').hidden = false;
}

function pacingRowHtml(r) {
  const id = String(r.id || newPacingRowId()).replace(/"/g, '');
  const fo = clampPct0to100(r.finalOccupancyPct);
  const tp = clampPct0to100(r.targetPricePercentile);
  let low =
    r.lowestPricePercentile != null && r.lowestPricePercentile !== ''
      ? clampPct0to100(r.lowestPricePercentile)
      : Math.max(0, tp - 30);
  if (low > tp) low = tp;
  const wks = r.weeksToStartReducing != null ? clampWeeksNonNeg(r.weeksToStartReducing) : 3;
  const fin = finalizeLowerAndWeeksAhead(tp, low, wks, r.lowerPercentileTo, r.weeksAhead);
  const lp = fin.lowerPercentileTo;
  // Must show stored Wks to Lower, not fin.weeksAhead (finalize often zeros it for curve rules).
  const wa = coalesceStoredWeeksAhead(r, fin.weeksAhead);
  return `
    <tr data-row-id="${id}">
      <td class="tiers-drag-cell">
        <span
          class="tier-drag-handle"
          draggable="true"
          title="Drag to reorder"
          aria-label="Drag to reorder row"
        >⋮⋮</span>
      </td>
      <td>
        <input type="number" min="0" max="100" step="1" value="${fo}" data-field="finalOccupancyPct" title="LY final at or above this % → use this row" />
      </td>
      <td>
        <input type="number" min="0" max="100" step="1" value="${tp}" data-field="targetPricePercentile" title="Target Pctl % when stay is ≥ Wks to Lowest away" />
      </td>
      <td>
        <input type="number" min="0" max="100" step="1" value="${lp}" data-field="lowerPercentileTo" title="Lower Pctl % — between Target and Lowest (two-step ramp)" />
      </td>
      <td>
        <input class="tiers-input-wide" type="number" min="0" max="104" step="0.1" value="${formatWeeksOneDecimal(wa)}" data-field="weeksAhead" title="Wks to Lower — weeks before stay when ramp hits Lower Pctl; 0 = one-step to Lowest" />
      </td>
      <td>
        <input type="number" min="0" max="100" step="1" value="${low}" data-field="lowestPricePercentile" title="Lowest Pctl% at first night of stay" />
      </td>
      <td>
        <input class="tiers-input-wide" type="number" min="0" max="104" step="0.1" value="${formatWeeksOneDecimal(wks)}" data-field="weeksToStartReducing" title="Wks to Lowest — start ramp from Target; 0 = always Target" />
      </td>
      <td>
        <button type="button" class="btn-row-delete btn-small" aria-label="Delete row">&times;</button>
      </td>
    </tr>`;
}

function appendPacingRowToTable(row) {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  const wrap = document.createElement('tbody');
  wrap.innerHTML = pacingRowHtml(row).trim();
  const tr = wrap.querySelector('tr');
  if (tr) tbody.appendChild(tr);
}

function populateTiersTable(rows) {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  tbody.innerHTML = (rows || []).map((r) => pacingRowHtml(r)).join('');
}

/** Read pacing rows from the modal table (empty → []). */
function readPacingRowsFromFormBody(tbody) {
  const trs = tbody.querySelectorAll('tr');
  const newRows = [];
  trs.forEach((row) => {
    const fo = row.querySelector('[data-field="finalOccupancyPct"]');
    const tp = row.querySelector('[data-field="targetPricePercentile"]');
    const lo = row.querySelector('[data-field="lowestPricePercentile"]');
    const wk = row.querySelector('[data-field="weeksToStartReducing"]');
    const lpIn = row.querySelector('[data-field="lowerPercentileTo"]');
    const waIn = row.querySelector('[data-field="weeksAhead"]');
    const rid = row.dataset.rowId || newPacingRowId();
    let tpp = clampPct0to100(tp ? tp.value : 0);
    let low = clampPct0to100(lo ? lo.value : Math.max(0, tpp - 30));
    if (low > tpp) low = tpp;
    const wksV = wk != null ? clampWeeksNonNeg(wk.value) : 3;
    const fin = finalizeLowerAndWeeksAhead(tpp, low, wksV, lpIn ? lpIn.value : null, waIn ? waIn.value : null);
    const weeksAheadStored = coalesceStoredWeeksAhead(
      waIn != null && waIn.value !== '' ? { weeksAhead: waIn.value } : {},
      fin.weeksAhead
    );
    newRows.push({
      id: rid,
      finalOccupancyPct: clampPct0to100(fo ? fo.value : 0),
      targetPricePercentile: tpp,
      lowestPricePercentile: low,
      weeksToStartReducing: wksV,
      lowerPercentileTo: fin.lowerPercentileTo,
      weeksAhead: weeksAheadStored,
    });
  });
  return newRows;
}

function savePacingSettingsFromForm() {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  let rows = readPacingRowsFromFormBody(tbody);
  if (rows.length === 0) {
    rows = JSON.parse(JSON.stringify(loadUserPacingDefaults()));
    populateTiersTable(rows);
  }
  savePacingTiers(rows);
}

/** Persist current table as the template used by “Reset to defaults”. */
function savePacingDefaultsFromForm() {
  const tbody = document.getElementById('tiersTableBody');
  if (!tbody) return;
  let rows = readPacingRowsFromFormBody(tbody);
  if (rows.length === 0) rows = JSON.parse(JSON.stringify(loadUserPacingDefaults()));
  saveUserPacingDefaults(rows);
}

init();
