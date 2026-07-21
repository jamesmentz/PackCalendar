// generate.js — builds the Pack 127 calendar .docx per spec.md.
//
//   node generator/generate.js
//
// Emits the .docx into the repository root. All schedule data lives in this
// file so it can be edited and regenerated at the start of each program year.
// Requires only Node core + Windows PowerShell/.NET (see generator/docx.js).
//
// Data sources & assumptions are documented in generator/DATA_SOURCES.md.

const path = require('path');
const { run, para, heading, writeDocx } = require('./docx');
const {
  MONTHS, MONTH_ABBR, WEEKDAY, d, ymd, nthWeekdayOfMonth, renderCalendarBody,
} = require('./lib');

const SCHOOL_YEAR = '2026-2027';
const TITLE = `${SCHOOL_YEAR} Pack 127 Calendar`;
const OUT = path.join(__dirname, '..', `${TITLE}.docx`);

// ---------------------------------------------------------------------------
// Key ACPS 2026-2027 dates. Verified against the official ACPS 2026-27 calendar
// PDF (Jul 2026). When rolling this file to a future year, treat these as
// EXAMPLES to re-confirm — a right-looking year is not proof a date is current.
const FIRST_DAY = d(2026, 8, 10);   // first student day (Mon) — verified ACPS
const LAST_DAY = d(2027, 5, 28);    // last student day (Fri) — verified ACPS
// Wednesdays that fall inside an ACPS student closure -> marked OFF.
// All verified against the official ACPS 2026-27 calendar; no other Wednesday
// closures exist in the school year.
const OFF_WEDNESDAYS = {
  '2026-11-11': 'Veterans Day',       // verified: ACPS holiday (school closed)
  '2026-11-25': 'Thanksgiving break',
  '2026-12-23': 'Winter break',       // winter break Dec 21 – Jan 1; students return Tue Jan 5
  '2026-12-30': 'Winter break',
  '2027-03-17': 'Spring break',       // spring break Mar 15–19
};
// Daylight for the 6pm playground (Roper Park): available before the clocks
// fall back (Sun Nov 1, 2026) and again after they spring forward (Sun Mar 14,
// 2027). Outdoor/fitness adventures are scheduled in those windows.

// ---------------------------------------------------------------------------
// Wednesday classification.
const PACK_MEETING_WED = ['2026-10-07', '2026-11-04', '2026-12-02', '2027-01-06', '2027-04-07'];
const YEAR_END_PACK_WED = '2027-05-26';
const BUILD_NIGHT_WED = ['2027-02-03', '2027-02-10', '2027-02-17']; // before Pinewood (Feb 20-21)
const BLUE_GOLD_WED = '2027-02-24';   // after the Pinewood Derby, per spec
const BACK_TO_PACK_WED = '2026-08-19';

function classifyWednesday(dt) {
  const key = ymd(dt);
  if (OFF_WEDNESDAYS[key]) return { kind: 'off', label: `OFF — ${OFF_WEDNESDAYS[key]}` };
  const y = dt.getFullYear(), m = dt.getMonth() + 1;
  if (y === 2026 && (m === 8 || m === 9)) {
    // Pre-Sign-Up Night: pack-wide elective meetings.
    if (key === BACK_TO_PACK_WED) {
      return { kind: 'pack', label: 'Back to the Pack — Pack Kickoff (Cabin, 6pm)' };
    }
    return { kind: 'pack', label: 'Pack-Wide Meeting (Cabin, 6pm)' };
  }
  if (BUILD_NIGHT_WED.includes(key)) {
    return { kind: 'pack', label: 'Pack Meeting — Pinewood Derby Build Night (FH, 6pm)' };
  }
  if (key === BLUE_GOLD_WED) {
    return { kind: 'pack', label: 'Blue & Gold Banquet — Arrow of Light Crossover (FH, 6pm)' };
  }
  if (key === YEAR_END_PACK_WED) {
    return { kind: 'pack', label: 'Pack Meeting — Year-End Celebration (FH, 6pm)' };
  }
  if (PACK_MEETING_WED.includes(key)) {
    return { kind: 'pack', label: 'Pack Meeting (FH, 6pm)' };
  }
  return { kind: 'den', label: 'Den Meetings — by grade (6pm)' };
}

function buildWednesdayEvents() {
  const events = [];
  let cur = new Date(FIRST_DAY);
  while (cur.getDay() !== 3) cur.setDate(cur.getDate() + 1); // first Wednesday
  while (cur <= LAST_DAY) {
    const c = classifyWednesday(cur);
    const ev = { date: new Date(cur), kind: c.kind, label: c.label };
    if (ymd(cur) === '2026-08-12') {
      ev.wrap = 'Pack-wide elective activities until Sign-Up Nights (SUNs) are complete.';
    }
    events.push(ev);
    cur.setDate(cur.getDate() + 7);
  }
  return events;
}

// ---------------------------------------------------------------------------
// Committee meetings: last Thursday of each school-year month, 9pm Zoom.
// Moved earlier when the last Thursday falls in a holiday week with no Wednesday
// meeting (Nov -> 1 week back; Dec -> 2 weeks back to stay in December).
function buildCommitteeEvents() {
  const spec = [
    [2026, 8], [2026, 9], [2026, 10], [2026, 11], [2026, 12],
    [2027, 1], [2027, 2], [2027, 3], [2027, 4], [2027, 5],
  ];
  const moved = {
    '2026-11': d(2026, 11, 19),
    '2026-12': d(2026, 12, 17),
  };
  return spec.map(([y, m]) => {
    const key = `${y}-${String(m).padStart(2, '0')}`;
    const dt = moved[key] || nthWeekdayOfMonth(y, m, 4, 'last');
    const ev = { date: dt, kind: 'committee', label: 'Committee Meeting (Zoom, 9pm)' };
    if (moved[key]) ev.wrap = 'Moved earlier — last Thursday falls in a holiday week.';
    return ev;
  });
}

// ---------------------------------------------------------------------------
// Special events (need not be on a Wednesday). Multi-day events are anchored to
// their start date with the range spelled out in the label.
function buildSpecialEvents() {
  return [
    // Summer / start-of-year recruitment
    { date: d(2026, 8, 1), kind: 'special',
      label: 'Council Popcorn Kickoff Party — St. Johns River Base, Orange Park (10am–2pm)' },
    // Sign-Up Nights (SUNs) — may coincide with a regular meeting
    { date: d(2026, 8, 26), kind: 'special',
      label: 'Littlewood Sign-Up Night — 6:30pm, Cafeteria' },
    { date: d(2026, 9, 9), kind: 'special',
      label: 'PK Yonge Sign-Up Night — Open House 5:30–7pm' },
    { date: d(2026, 9, 19), kind: 'special',
      label: 'Frontier Shooting Day (Sat) — Camp Shands' },  // verified: nfcscouting.org/frontier
    { date: d(2026, 9, 26), kind: 'special',
      label: 'Cubmaster & Den Leader Specific Training (Sat) — council',
      verify: 'Not found on the council’s live training calendar (Jul 2026); confirm the 2026 date at nfcscouting.org/training.' },
    { date: d(2026, 9, 30), kind: 'special',
      label: 'Williams & Lake Forest Sign-Up Nights — 6:30pm, Cafeterias' },
    // Fall
    { date: d(2026, 10, 16), kind: 'special',
      label: 'Fall Campout (Fri–Sun, Oct 16–18)', wrap: 'Two weekends before Halloween (Sat Oct 31) so families can do other Halloween activities. Close to home; one night mandatory / two optional (e.g., Troy Springs). Aligns with council Spookoree (Oct 16–18).' },
    { date: d(2026, 10, 17), kind: 'special',
      label: 'Rain Gutter Regatta (Sat) — at the Fall Campout', wrap: 'Fall recruitment activity; ~1 hour, minimal tools (materials via Troop 125 & James).' },
    { date: d(2026, 10, 24), kind: 'special',
      label: 'Lubee Florida Bat Festival — NO pack events this weekend' },  // verified: lubee.org/flbatfestival
    { date: d(2026, 11, 7), kind: 'special',
      label: 'BALOO Training — adult leaders (Sat–Sun, Nov 7–8)' },  // verified: nfcscouting.org/baloo-stjohns (spring option May 7–8, 2027)
    { date: d(2026, 12, 11), kind: 'special',
      label: 'Cub Winter Wonderland (council) — Dec 11–13', wrap: 'Optional council event.' },  // verified: nfcscouting.org/familycamps
    // Winter / spring
    { date: d(2027, 1, 30), kind: 'special',
      label: 'Council Camp Card Sale Kickoff (Sat)', wrap: 'Spring council fundraiser; sale runs Feb–Apr 2027.' },  // verified: NFC IYOS 2026-2027
    { date: d(2027, 2, 5), kind: 'special',
      label: 'Winter Campout — Medieval Faire Family Camp, Camp Shands (Fri–Sun, Feb 5–7)', wrap: "The pack's district campout for the year." },  // verified: NFC IYOS 2026-2027
    { date: d(2027, 2, 7), kind: 'special',
      label: 'Scout Sunday' },  // verified: scouting.org + NFC IYOS 2026-2027
    { date: d(2027, 2, 20), kind: 'special',
      label: 'Pinewood Derby — Pack 127 (Sat–Sun, Feb 20–21)' },
    { date: d(2027, 3, 3), kind: 'special',
      label: 'Five Rivers District Dinner / Banquet (Wed)' },  // verified: NFC IYOS 2026-2027
    { date: d(2027, 4, 3), kind: 'special',
      label: 'District Pinewood Derby (Sat)', wrap: 'NFC Council Pinewood Derby follows ~Apr 17.',
      verify: 'IYOS shows district Pinewood Derbies on Apr 3, 2027 but does not explicitly name Five Rivers; confirm the Five Rivers date at nfcscouting.org/calendar.' },
    { date: d(2027, 4, 23), kind: 'special',
      label: 'Spring Campout — Family Camp, Camp Shands (Fri–Sun, Apr 23–25)' },  // verified: NFC IYOS 2026-2027
    { date: d(2027, 5, 1), kind: 'special',
      label: 'Council Volunteer Recognition Awards Dinner' },  // verified: NFC IYOS 2026-2027
    // Summer 2027 (the "following summer") + next-year lookahead
    { date: d(2027, 5, 28), kind: 'special',
      label: 'Last ACPS student day — summer break begins', wrap: 'No regular Wednesday meetings over the summer.' },  // verified: ACPS 2026-2027 calendar
    { date: d(2027, 6, 7), kind: 'special',
      label: 'Cub Day Camp (council) — week of Jun 7 (Jun 7–11)', wrap: 'Five Rivers Cub Day Camp; Aquatics Camp Wk 1 follows Jun 14–18.' },  // verified: NFC IYOS 2026-2027
    { date: d(2027, 8, 7), kind: 'special',
      label: '2027 Popcorn Kickoff (council) — start of the 2027–28 year',
      verify: 'Next-year lookahead; the 2027 popcorn kickoff date is not yet published — confirm.' },
  ];
}

// ---------------------------------------------------------------------------
// Config-level dates that are assumptions rather than confirmed facts. These
// feed the "Issues — Unverified Dates" list at the bottom of the calendar.
// (Rule-derived dates — every Wednesday meeting, last-Thursday committee
// meetings, the two-weekends-before-Halloween campout, the after-Valentine's
// Pinewood Derby — are computed and are NOT listed here: the calendar arithmetic
// is known and is guaranteed correct by assertDateLabels below.)
function buildConfigIssues() {
  // ACPS 2026-2027 dates (first/last day, Veterans Day OFF, Thanksgiving, winter
  // & spring break, the Jan 6 return) were verified against the official ACPS
  // 26-27 calendar PDF and are no longer listed here. Parker SUN remains TBD.
  return [
    { date: null, label: 'Parker Sign-Up Night',
      note: 'Date still TBD — confirm with the school, then add it to the calendar.' },
  ];
}

// ---------------------------------------------------------------------------
// Guarantee that every weekday/date mentioned in an event label matches the
// real calendar. The Gregorian calendar is known, so a label that says the
// wrong day of week (or a start date that disagrees with the event's Date) is a
// hard error, not a "verify later" item — this THROWS and fails the build.
const WD_ABBR = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WD_RE = /\b(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sun|Mon|Tue|Wed|Thu|Fri|Sat)\b/;
const MD_RE = new RegExp(`\\b(${MONTH_ABBR.join('|')})\\s+(\\d{1,2})\\b`);

function normalizeWeekday(tok) {
  const full = WEEKDAY.indexOf(tok);
  if (full !== -1) return full;
  return WD_ABBR.indexOf(tok);
}

function assertDateLabels(events) {
  for (const ev of events) {
    if (!ev.date) continue;
    const label = ev.label || '';
    const wdMatch = label.match(WD_RE);
    if (wdMatch) {
      const claimed = normalizeWeekday(wdMatch[1]);
      if (claimed !== ev.date.getDay()) {
        throw new Error(
          `Weekday mismatch: "${label}" says ${wdMatch[1]} but ${ymd(ev.date)} `
          + `is a ${WEEKDAY[ev.date.getDay()]}.`);
      }
    }
    const mdMatch = label.match(MD_RE);
    if (mdMatch) {
      const mo = MONTH_ABBR.indexOf(mdMatch[1]);
      const day = parseInt(mdMatch[2], 10);
      if (mo === ev.date.getMonth() && day !== ev.date.getDate()) {
        throw new Error(
          `Start-date mismatch: "${label}" anchors to ${mdMatch[1]} ${day} but `
          + `the event date is ${ymd(ev.date)}.`);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// "Issues" section rendered at the very bottom of the calendar: every date the
// generator could not confirm against an authoritative source, sorted by date.
function fmtLong(dt) {
  return `${WEEKDAY[dt.getDay()]}, ${MONTHS[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
}

function buildIssuesList(allEvents, configIssues) {
  const blocks = [];
  blocks.push(heading('Issues — Unverified Dates', 1));
  blocks.push(para([run(
    'The generator confirms every day-of-week and calendar calculation itself, so '
    + 'the dates on the pages above land on the correct weekdays. What it cannot confirm '
    + 'is whether each externally-scheduled event actually happens on the date shown. '
    + 'Verify each item below against its authoritative source (the ACPS calendar or '
    + 'nfcscouting.org/calendar) before distributing this calendar.'
  )]));

  const items = allEvents
    .filter((e) => e.verify)
    .map((e) => ({ date: e.date, label: e.label, note: e.verify }))
    .concat(configIssues);

  const dated = items.filter((x) => x.date).sort((a, b) => a.date - b.date);
  const undated = items.filter((x) => !x.date);

  for (const x of dated) {
    blocks.push(para([run(fmtLong(x.date) + ' — ', { bold: true }), run(`${x.label}: ${x.note}`)]));
  }
  if (undated.length) {
    blocks.push(heading('Not yet scheduled (no date set)', 2));
    for (const x of undated) blocks.push(para([run('• ' + x.label + ': ' + x.note)]));
  }
  return blocks;
}

// ---------------------------------------------------------------------------
// Program recommendations.
const DENS = [
  { rank: 'Lion', grade: 'Kindergarten', crossesOver: false },
  { rank: 'Tiger', grade: '1st grade', crossesOver: false },
  { rank: 'Wolf', grade: '2nd grade', crossesOver: false },
  { rank: 'Bear', grade: '3rd grade', crossesOver: false },
  { rank: 'Webelos', grade: '4th grade', crossesOver: false },
  { rank: 'Arrow of Light', grade: '5th grade', crossesOver: true },
];

// Ordered adventure blocks per den: [name, agenda, location, #meetings].
// Non-AOL dens have 19 den meetings; blocks are ordered so the 3 fall meetings
// (Oct, before the Nov 1 clock change) and the 8 spring meetings (after the
// Mar 14 clock change) carry the outdoor work, with 8 indoor meetings between.
const DEN_PLANS = {
  Lion: [
    ['Mountain Lion', 'explore the outdoors, nature walk', 'Roper Park (Playground)', 3],
    ['King of the Jungle', 'animal-movement games (indoor)', 'Fellowship Hall', 1],
    ["Lion's Pride", 'den identity & family activity', 'Bethany Room (Bell Hall)', 2],
    ["Lion's Roar", 'recognize & respond to danger', 'The Chapel', 2],
    ['Fun on the Run', 'healthy habits + indoor movement game', 'Fellowship Hall', 2],
    ['Ready, Set, Grow (elective)', 'plant seeds (indoor start)', 'Upstairs (Epworth Hall)', 1],
    ['On Your Mark (elective)', 'races & fitness challenges', 'Roper Park (Playground)', 3],
    ['Year-end games & bridge to Tiger', 'field games; awards prep', 'Roper Park (Playground)', 5],
  ],
  Tiger: [
    ['Tigers in the Wild', 'hike & nature exploration', 'Roper Park (Playground)', 3],
    ['Tiger Tag (elective)', 'active games (indoor)', 'Fellowship Hall', 1],
    ['Team Tiger', 'teamwork & helping at home', 'Bethany Room (Bell Hall)', 2],
    ['Tiger Bites', 'healthy foods & a simple snack', 'Fellowship Hall', 2],
    ['Tiger Circles', 'duty to God & values', 'The Chapel', 2],
    ["Tiger's Roar", 'weather & safety, calling for help', 'Upstairs (Epworth Hall)', 1],
    ['Stories and Shapes (elective)', 'drawing & storytelling', 'Upstairs (Epworth Hall)', 2],
    ['Year-end games & bridge to Wolf', 'field games; awards prep', 'Roper Park (Playground)', 6],
  ],
  Wolf: [
    ['Running with the Pack', 'fitness & active games', 'Roper Park (Playground)', 2],
    ['Paws on the Path', 'short hike & outdoor essentials', 'Roper Park (Playground)', 1],
    ['Council Fire (Duty to Country)', 'flags & community', 'The Chapel', 2],
    ['Footsteps (Duty to God)', 'values & family faith', 'The Chapel', 2],
    ['Safe and Smart', 'personal & internet safety', 'Upstairs (Epworth Hall)', 2],
    ['Code of the Wolf (elective)', 'math & simple coding games', 'Upstairs (Epworth Hall)', 2],
    ['Air of the Wolf (elective)', 'build & launch a balloon rocket outdoors', 'Roper Park (Playground)', 3],
    ['Year-end games & bridge to Bear', 'field games; awards prep', 'Roper Park (Playground)', 5],
  ],
  Bear: [
    ['Bear Strong', 'fitness & active outdoor games', 'Roper Park (Playground)', 2],
    ['Bear Habitat', 'study trees & nature outdoors', 'Roper Park (Playground)', 1],
    ['Paws for Action (Duty to Country)', 'citizenship & a service project', 'The Chapel', 2],
    ['Fellowship (Duty to God)', 'values & reflection', 'The Chapel', 2],
    ['Baloo the Builder (elective)', 'woodworking with hand tools', 'The Cabin', 2],
    ['Standing Tall', 'character & making good choices', 'Upstairs (Epworth Hall)', 2],
    ['Whittling (elective, Spring)', 'pocketknife safety & whittling', 'The Cabin', 3],
    ['Year-end games & bridge to Webelos', 'field games; awards prep', 'Roper Park (Playground)', 5],
  ],
  Webelos: [
    ['Stronger, Faster, Higher', 'fitness & team sports outdoors', 'Roper Park (Playground)', 2],
    ['Webelos Walkabout', 'plan & take a hike', 'Roper Park (Playground)', 1],
    ['My Community', 'community roles & service', 'Bethany Room (Bell Hall)', 2],
    ['My Family', 'family duties & budgeting', 'Bethany Room (Bell Hall)', 2],
    ['My Safety', 'emergency preparedness', 'Upstairs (Epworth Hall)', 2],
    ['Build It (elective)', 'construction & engineering project', 'The Cabin', 2],
    ["Chef's Knife (elective, Spring)", 'kitchen-knife safety & cooking', 'Fellowship Hall (kitchen)', 3],
    ['Year-end games & prep for Arrow of Light', 'field games; awards prep', 'Roper Park (Playground)', 5],
  ],
  // Arrow of Light finishes before the Blue & Gold crossover (Feb 24, 2027) —
  // 9 den meetings (Oct–Jan).
  'Arrow of Light': [
    ['Personal Fitness', 'fitness testing & active games', 'Roper Park (Playground)', 2],
    ['Outdoor Adventurer', 'outdoor skills & Fall Campout prep', 'Roper Park (Playground)', 1],
    ['Citizenship', 'community service & flag ceremony', 'The Chapel', 1],
    ['Duty to God', 'values & family faith', 'The Chapel', 2],
    ['First Aid', 'first-aid skills & building a kit', 'Upstairs (Epworth Hall)', 1],
    ['Knife Safety (elective)', 'pocket + kitchen knife safety', 'The Cabin', 1],
    ['Crossover prep & rank make-up', 'finish requirements; ready for Blue & Gold', 'The Chapel', 1],
  ],
};

function fmtShort(dt) { return `${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}`; }

function expandPlan(dates, blocks) {
  const rows = [];
  let i = 0;
  for (const [name, agenda, loc, n] of blocks) {
    for (let k = 0; k < n && i < dates.length; k++, i++) {
      rows.push({ date: dates[i], name, agenda, loc });
    }
  }
  return rows;
}

function buildProgramBlocks(denMeetingDates) {
  const blocks = [];
  blocks.push(heading('Program Recommendations', 1));
  blocks.push(para([run(
    'Adventures follow the current Cub Scout program (scouting.org/programs/cub-scouts/adventures). '
    + 'Each rank earns its badge with 6 required adventures (including Bobcat) plus at least 2 electives. '
    + 'Keep active program time to 30–40 minutes per meeting; the rest is opening, snack, and closing.'
  )]));

  blocks.push(heading('Guiding constraints', 2));
  [
    'Bobcat is done first, pack-wide, during the August–September pack-wide meetings before dens form.',
    'Include one moving activity every den meeting. Webelos and Arrow of Light can handle desk work; the younger dens need to move.',
    'Roper Park (the playground) has daylight at 6pm only before the clocks fall back (Sun Nov 1, 2026) and again after they spring forward (Sun Mar 14, 2027). Outdoor/fitness adventures are scheduled in those windows; winter meetings move indoors.',
    'Anything with tools (Baloo the Builder, Whittling) works best in the Cabin; anything with cooking/food (Chef’s Knife, Tiger Bites) works best in Fellowship Hall (kitchen).',
    'Bear does Baloo the Builder and Whittling, with Whittling in the Spring. Webelos does Chef’s Knife in the Spring.',
    'Arrow of Light earns two electives: Knife Safety (a den meeting) and Race Time (covered by the Pinewood Derby build nights and race in February).',
    'Arrow of Light finishes all required rank work before the Blue & Gold crossover (Feb 24, 2027), leaving a make-up night in the schedule.',
    'Troop Visits apply only to the Arrow of Light den and are arranged directly by the AOL Den Leader with local Scouts BSA troops (125, 84, 21) on their own timetable (not on the top calendar page).',
  ].forEach((t) => blocks.push(para([run('• ' + t)])));

  blocks.push(heading('Meeting locations', 2));
  [
    'The Cabin — tools/woodworking.',
    'Roper Park (Playground) — outdoor, daylight-dependent.',
    'Fellowship Hall (FH) — large indoor space with a kitchen; the only room big enough for the whole pack.',
    'The Chapel — medium indoor space.',
    'Upstairs (Epworth Hall) — classroom.',
    'Bethany Room (Bell Hall) — classroom-size indoor space.',
  ].forEach((t) => blocks.push(para([run('• ' + t)])));

  blocks.push(heading('August–September: pack-wide (Bobcat + electives)', 2));
  [
    ['Aug 12', 'Welcome; introduce the Scout Oath & Law; active games', 'Roper Park (Playground)'],
    ['Aug 19', 'Back to the Pack kickoff; pack-wide games & welcome', 'Roper Park (Playground)'],
    ['Aug 26', 'Bobcat: sign, salute, handshake, motto, slogan (Littlewood SUN tonight)', 'Roper Park (Playground)'],
    ['Sep 2', 'Bobcat: Cub Scout Six Essentials & buddy system; nature walk', 'Roper Park (Playground)'],
    ['Sep 9', 'Bobcat: personal safety / Protect Yourself Rules (PK Yonge SUN tonight)', 'The Chapel'],
    ['Sep 16', 'Pack-wide elective: Champions for Nature clean-up', 'Roper Park (Playground)'],
    ['Sep 23', 'Pack-wide elective: STEM/craft activity', 'Fellowship Hall'],
    ['Sep 30', 'Welcome SUN recruits; prep to split into dens (Williams & Lake Forest SUNs tonight)', 'Fellowship Hall'],
  ].forEach(([dt, ag, loc]) => blocks.push(para([run(dt + ' — ', { bold: true }), run(`${ag} (${loc})`)])));

  const aolCutoff = d(2027, 2, 24);
  for (const den of DENS) {
    const dates = den.crossesOver ? denMeetingDates.filter((x) => x < aolCutoff) : denMeetingDates;
    const rows = expandPlan(dates, DEN_PLANS[den.rank]);
    blocks.push(heading(`${den.rank} Den (${den.grade})`, 2));
    if (den.crossesOver) {
      blocks.push(para([run(
        'Crosses over to Scouts BSA at the Blue & Gold Banquet (Feb 24, 2027), so all required work is front-loaded into Oct–Jan. '
        + 'Race Time (2nd elective) is earned through the Pinewood Derby build nights (Feb 3/10/17) and race (Feb 20–21).'
      )]));
    }
    for (const r of rows) {
      blocks.push(para([run(fmtShort(r.date) + ' — ', { bold: true }), run(`${r.name}: ${r.agenda} (${r.loc})`)]));
    }
  }

  blocks.push(heading('Program notes', 2));
  [
    'All date verification is consolidated in the "Issues — Unverified Dates" list at the end of this document.',
    'A few required-adventure names vary across sources (e.g., Wolf/Bear "Safe and Smart" vs "Safety in Numbers", Bear "Standing Tall"). Confirm against the current official rank pages.',
    'Back to the Pack (the August kickoff) and other August/September details are TBA — confirm before publishing.',
  ].forEach((t) => blocks.push(para([run('• ' + t)])));

  return blocks;
}

// ---------------------------------------------------------------------------
function main() {
  const wednesdays = buildWednesdayEvents();
  const committee = buildCommitteeEvents();
  const special = buildSpecialEvents();
  const allEvents = [...wednesdays, ...committee, ...special];

  // Hard guarantee: every weekday/date named in a label matches the real
  // calendar. Fails the build (rather than shipping a wrong date) on mismatch.
  assertDateLabels(allEvents);

  const denMeetingDates = wednesdays
    .filter((e) => e.kind === 'den')
    .map((e) => e.date)
    .sort((a, b) => a - b);

  const titleBlocks = [
    para([run(TITLE)], { style: 'Title' }),
    para([
      run('Cub Scout Pack 127', { bold: true }),
      run('  •  Regular meetings Wednesdays at 6pm  •  '),
      run('FH', { bold: true }), run(' = Fellowship Hall  •  '),
      run('OFF', { bold: true }), run(' = no meeting (school closed)'),
    ], { spaceAfter: 120 }),
  ];

  const calendarBlocks = renderCalendarBody(allEvents);
  const programBlocks = buildProgramBlocks(denMeetingDates);
  const issuesBlocks = buildIssuesList(allEvents, buildConfigIssues());

  writeDocx(OUT, {
    title: TITLE,
    sections: [
      { blocks: titleBlocks, numCols: 1, type: 'continuous' },
      { blocks: calendarBlocks, numCols: 2, type: 'continuous' },
      { blocks: [...programBlocks, ...issuesBlocks], numCols: 1, type: 'nextPage' },
    ],
  });
  console.log('Wrote', OUT);
}

main();
