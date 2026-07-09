// lib.js — date helpers and calendar-rendering engine for the Pack 127 calendar.
// Data-independent: consumes an ordered list of event objects and renders the
// two-column flowing body per the spec's formatting rules.

const { run, para } = require('./docx');

const MONTHS = ['January','February','March','April','May','June','July',
  'August','September','October','November','December'];
const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const WEEKDAY = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

// Build a Date at local noon (avoids DST/midnight edge cases in date math).
function d(y, m /*1-12*/, day) { return new Date(y, m - 1, day, 12, 0, 0); }
function ymd(dt) { return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`; }
function addDays(dt, n) { const r = new Date(dt); r.setDate(r.getDate() + n); return r; }
function sameDay(a, b) { return ymd(a) === ymd(b); }

// All dates in [start, end] (inclusive) falling on the given weekday (0=Sun..6=Sat).
function weekdaysBetween(start, end, weekday) {
  const out = [];
  let cur = new Date(start);
  while (cur.getDay() !== weekday) cur = addDays(cur, 1);
  while (cur <= end) { out.push(new Date(cur)); cur = addDays(cur, 7); }
  return out;
}

// nth given weekday of a month, or last one if n === 'last'.
function nthWeekdayOfMonth(year, month /*1-12*/, weekday, n) {
  const days = [];
  let cur = d(year, month, 1);
  while (cur.getMonth() === month - 1) {
    if (cur.getDay() === weekday) days.push(new Date(cur));
    cur = addDays(cur, 1);
  }
  return n === 'last' ? days[days.length - 1] : days[n - 1];
}

// ---- Calendar body rendering -------------------------------------------------
//
// Event shape:
//   { date: Date, kind: 'pack'|'den'|'committee'|'special'|'off',
//     label: string,          // e.g. "Pack Meeting (Cabin 6pm)" or "OFF - Thanksgiving"
//     wrap?: string }          // optional extra line rendered under the entry
//
// Month markers: month name on its own bold line. Per spec, the first and last
// month of the calendar, plus every December and January, include the year.

function monthNeedsYear(monthDate, firstMonthKey, lastMonthKey) {
  const key = `${monthDate.getFullYear()}-${monthDate.getMonth()}`;
  const m = monthDate.getMonth(); // 0=Jan, 11=Dec
  return key === firstMonthKey || key === lastMonthKey || m === 0 || m === 11;
}

function renderCalendarBody(events) {
  const sorted = [...events].sort((a, b) => a.date - b.date);
  const firstMonthKey = `${sorted[0].date.getFullYear()}-${sorted[0].date.getMonth()}`;
  const last = sorted[sorted.length - 1].date;
  const lastMonthKey = `${last.getFullYear()}-${last.getMonth()}`;

  const blocks = [];
  let curMonthKey = null;
  for (const ev of sorted) {
    const mk = `${ev.date.getFullYear()}-${ev.date.getMonth()}`;
    if (mk !== curMonthKey) {
      curMonthKey = mk;
      const withYear = monthNeedsYear(ev.date, firstMonthKey, lastMonthKey);
      const name = MONTHS[ev.date.getMonth()] + (withYear ? ` ${ev.date.getFullYear()}` : '');
      blocks.push(para([run(name, { bold: true })], { spaceBefore: 120, spaceAfter: 20, keepNext: true }));
    }
    // Entry: bold day number, then the label. Keep day+label together.
    const runs = [run(String(ev.date.getDate()) + ' ', { bold: true }), run(ev.label)];
    blocks.push(para(runs, { spaceAfter: 20 }));
    if (ev.wrap) blocks.push(para([run('    ' + ev.wrap)], { spaceAfter: 20 }));
  }
  return blocks;
}

module.exports = {
  MONTHS, MONTH_ABBR, WEEKDAY,
  d, ymd, addDays, sameDay, weekdaysBetween, nthWeekdayOfMonth,
  renderCalendarBody,
};
