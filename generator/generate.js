// generate.js — builds "2025-2026 Pack 127 Calendar.docx" per spec.md.
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
  MONTH_ABBR, d, ymd, nthWeekdayOfMonth, renderCalendarBody,
} = require('./lib');

const SCHOOL_YEAR = '2025-2026';
const TITLE = `${SCHOOL_YEAR} Pack 127 Calendar`;
const OUT = path.join(__dirname, '..', `${TITLE}.docx`);

// ---------------------------------------------------------------------------
// Key ACPS 2025-2026 dates (verified against the approved district calendar).
const FIRST_DAY = d(2025, 8, 11);   // first student day
const LAST_DAY = d(2026, 6, 2);     // last student day
// Wednesdays that fall inside an ACPS student closure -> marked OFF.
const OFF_WEDNESDAYS = {
  '2025-11-26': 'Thanksgiving break',
  '2025-12-24': 'Winter break',
  '2025-12-31': 'Winter break',
  '2026-03-18': 'Spring break',
};
// Daylight windows for the 6pm playground (Roper Park):
//   available Aug 13 - Nov 1 2025 (before fall-back Sun Nov 2, 2025) and
//   Mar 8 2026 onward (after spring-forward Sun Mar 8, 2026).

// ---------------------------------------------------------------------------
// Wednesday classification.
const PACK_MEETING_WED = ['2025-10-01', '2025-11-05', '2025-12-03', '2026-01-07', '2026-04-01'];
const YEAR_END_PACK_WED = '2026-05-27';
const BUILD_NIGHT_WED = ['2026-02-04', '2026-02-11', '2026-02-18'];
const BLUE_GOLD_WED = '2026-03-11';

// Time/venue tag shown on page 1 (Pack meetings need Fellowship Hall; see spec).
function classifyWednesday(dt) {
  const key = ymd(dt);
  if (OFF_WEDNESDAYS[key]) return { kind: 'off', label: `OFF — ${OFF_WEDNESDAYS[key]}` };
  const y = dt.getFullYear(), m = dt.getMonth() + 1;
  if (y === 2025 && (m === 8 || m === 9)) {
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

// Build the ordered list of school-year Wednesday events.
function buildWednesdayEvents() {
  const events = [];
  let cur = new Date(FIRST_DAY);
  while (cur.getDay() !== 3) cur.setDate(cur.getDate() + 1); // first Wednesday
  while (cur <= LAST_DAY) {
    const c = classifyWednesday(cur);
    const ev = { date: new Date(cur), kind: c.kind, label: c.label };
    if (ymd(cur) === '2025-08-13') {
      ev.wrap = 'Pack-wide elective activities until Sign-Up Nights (SUNs) are complete.';
    }
    events.push(ev);
    cur.setDate(cur.getDate() + 7);
  }
  return events;
}

// ---------------------------------------------------------------------------
// Committee meetings: last Thursday of each school-year month, 9pm Zoom.
// Move by the smallest whole number of weeks if that week has no Wed meeting.
function buildCommitteeEvents() {
  const spec = [
    [2025, 8], [2025, 9], [2025, 10], [2025, 11], [2025, 12],
    [2026, 1], [2026, 2], [2026, 3], [2026, 4], [2026, 5],
  ];
  const moved = {
    // last Thursday lands in a holiday week with no Wednesday meeting -> back 1 wk
    '2025-11': d(2025, 11, 20),
    '2025-12': d(2025, 12, 18),
  };
  return spec.map(([y, m]) => {
    const key = `${y}-${String(m).padStart(2, '0')}`;
    const dt = moved[key] || nthWeekdayOfMonth(y, m, 4, 'last');
    const ev = { date: dt, kind: 'committee', label: 'Committee Meeting (Zoom, 9pm)' };
    if (moved[key]) ev.wrap = 'Moved one week earlier — last Thursday falls in a holiday week.';
    return ev;
  });
}

// ---------------------------------------------------------------------------
// Special events (need not be on a Wednesday). Multi-day events are anchored to
// their start date with the range spelled out in the label.
function buildSpecialEvents() {
  return [
    { date: d(2025, 8, 2), kind: 'special',
      label: 'Council Popcorn Kickoff Party — St. Johns River Base, Orange Park (10am–2pm)' },
    { date: d(2025, 8, 9), kind: 'special',
      label: 'Pack 127 Kickoff Party (Sat) — location TBD' },
    { date: d(2025, 10, 24), kind: 'special',
      label: 'Fall Campout (Fri–Sun, Oct 24–26)', wrap: 'Location TBD — weekend before Halloween.' },
    { date: d(2026, 1, 23), kind: 'special',
      label: 'Winter Campout (Fri–Sun, Jan 23–25) — TBD', wrap: 'Committee to confirm site & details.' },
    { date: d(2026, 2, 21), kind: 'special',
      label: 'Pinewood Derby — Pack 127 (Sat–Sun, Feb 21–22)' },
    { date: d(2026, 3, 7), kind: 'special',
      label: 'Five Rivers District Pinewood Derby (Sat) — Kanapaha Presbyterian, Gainesville' },
    { date: d(2026, 3, 28), kind: 'special',
      label: 'BALOO Training — adult leaders (Sat–Sun, Mar 28–29), Ocala', wrap: 'Confirm exact date on council flyer.' },
    { date: d(2026, 4, 24), kind: 'special',
      label: 'BUC TUOCS Spring Campout (Fri–Sun, Apr 24–26) — Camp Shands, TBD' },
    // Summer 2026 (the "following summer") + next-year lookahead.
    { date: d(2026, 6, 2), kind: 'special',
      label: 'Last ACPS student day — summer break begins', wrap: 'No regular Wednesday meetings over the summer.' },
    { date: d(2026, 8, 8), kind: 'special',
      label: 'Pack 127 Kickoff Party for 2026–27 (Sat) — end of summer, TBD' },
    { date: d(2026, 9, 19), kind: 'special',
      label: 'Frontier Day — shooting sports, Camp Shands (Sat)', wrap: 'Next council date; verify — falls at the start of the 2026–27 year.' },
  ];
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
// Sequenced so outdoor blocks fall in the playground-daylight windows.
const DEN_PLANS = {
  Lion: [
    ['Mountain Lion', 'explore the outdoors, nature walk', 'Roper Park (Playground)', 3],
    ['King of the Jungle', 'animal-movement games & running', 'Roper Park (Playground)', 1],
    ["Lion's Pride", 'den identity & family activity', 'Bethany Room (Bell Hall)', 3],
    ["Lion's Roar", 'recognize & respond to danger', 'The Chapel', 2],
    ['Fun on the Run', 'healthy habits + indoor movement game', 'Fellowship Hall', 3],
    ['Ready, Set, Grow (elective)', 'plant seeds & garden care', 'Upstairs (Epworth Hall)', 2],
    ['On Your Mark (elective)', 'races & fitness challenges', 'Roper Park (Playground)', 3],
    ['Year-end games & bridge to Tiger', 'field games; awards prep', 'Roper Park (Playground)', 4],
  ],
  Tiger: [
    ['Tigers in the Wild', 'hike & nature exploration', 'Roper Park (Playground)', 3],
    ['Tiger Tag (elective)', 'outdoor moving games', 'Roper Park (Playground)', 1],
    ['Team Tiger', 'teamwork & helping at home', 'Bethany Room (Bell Hall)', 3],
    ['Tiger Bites', 'healthy foods & a simple snack', 'Fellowship Hall', 2],
    ['Tiger Circles', 'duty to God & values', 'The Chapel', 3],
    ["Tiger's Roar", 'weather & safety, calling for help', 'Upstairs (Epworth Hall)', 2],
    ['Stories and Shapes (elective)', 'drawing & storytelling', 'Upstairs (Epworth Hall)', 3],
    ['Year-end games & bridge to Wolf', 'field games; awards prep', 'Roper Park (Playground)', 4],
  ],
  Wolf: [
    ['Running with the Pack', 'fitness & active games', 'Roper Park (Playground)', 3],
    ['Paws on the Path', 'short hike & outdoor essentials', 'Roper Park (Playground)', 1],
    ['Council Fire (Duty to Country)', 'flags & community', 'The Chapel', 3],
    ['Footsteps (Duty to God)', 'values & family faith', 'The Chapel', 2],
    ['Safe and Smart', 'personal & internet safety', 'Upstairs (Epworth Hall)', 3],
    ['Code of the Wolf (elective)', 'math & simple coding games', 'Upstairs (Epworth Hall)', 2],
    ['Air of the Wolf (elective)', 'build & launch a balloon rocket outdoors', 'Roper Park (Playground)', 3],
    ['Year-end games & bridge to Bear', 'field games; awards prep', 'Roper Park (Playground)', 4],
  ],
  Bear: [
    ['Bear Strong', 'fitness & active outdoor games', 'Roper Park (Playground)', 3],
    ['Bear Habitat', 'study trees & nature outdoors', 'Roper Park (Playground)', 1],
    ['Paws for Action (Duty to Country)', 'citizenship & a service project', 'The Chapel', 3],
    ['Fellowship (Duty to God)', 'values & reflection', 'The Chapel', 2],
    ['Baloo the Builder (elective)', 'woodworking with hand tools', 'The Cabin', 3],
    ['Standing Tall', 'character & making good choices', 'Upstairs (Epworth Hall)', 2],
    ['Whittling (elective, Spring)', 'pocketknife safety & whittling', 'The Cabin', 3],
    ['Year-end games & bridge to Webelos', 'field games; awards prep', 'Roper Park (Playground)', 4],
  ],
  Webelos: [
    ['Stronger, Faster, Higher', 'fitness & team sports outdoors', 'Roper Park (Playground)', 3],
    ['Webelos Walkabout', 'plan & take a hike', 'Roper Park (Playground)', 1],
    ['My Community', 'community roles & service', 'Bethany Room (Bell Hall)', 3],
    ['My Family', 'family duties & budgeting', 'Bethany Room (Bell Hall)', 2],
    ['My Safety', 'emergency preparedness', 'Upstairs (Epworth Hall)', 3],
    ['Build It (elective)', 'construction & engineering project', 'The Cabin', 2],
    ["Chef's Knife (elective, Spring)", 'kitchen-knife safety & cooking', 'Fellowship Hall (kitchen)', 3],
    ['Year-end games & prep for Arrow of Light', 'field games; awards prep', 'Roper Park (Playground)', 4],
  ],
  // Arrow of Light must finish before the Blue & Gold crossover (Mar 11) — 13 meetings.
  'Arrow of Light': [
    ['Personal Fitness', 'fitness testing & active games', 'Roper Park (Playground)', 2],
    ['Outdoor Adventurer', 'outdoor skills & Fall Campout prep', 'Roper Park (Playground)', 2],
    ['Citizenship', 'community service & flag ceremony', 'The Chapel', 2],
    ['Duty to God', 'values & family faith', 'The Chapel', 2],
    ['First Aid', 'first-aid skills & building a kit', 'Upstairs (Epworth Hall)', 2],
    ['Knife Safety (elective)', 'pocket + kitchen knife safety', 'The Cabin', 2],
    ['Crossover prep & rank make-up', 'finish requirements; ready for Blue & Gold', 'The Chapel', 1],
  ],
};

function fmtShort(dt) { return `${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}`; }

// Expand a den's ordered blocks onto its list of meeting dates.
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
    'Roper Park (the playground) has daylight at 6pm only before the clocks fall back (Sun Nov 2, 2025) and again after they spring forward (Sun Mar 8, 2026). Outdoor/fitness adventures are scheduled in those windows; winter meetings move indoors.',
    'Anything with tools (Baloo the Builder, Whittling) works best in the Cabin; anything with cooking/food (Chef’s Knife, Tiger Bites) works best in Fellowship Hall (kitchen).',
    'Bear does Baloo the Builder and Whittling, with Whittling in the Spring. Webelos does Chef’s Knife in the Spring.',
    'Arrow of Light earns two electives: Knife Safety (a den meeting) and Race Time (covered by the Pinewood Derby build nights and race in February).',
    'Arrow of Light finishes all required rank work before the Blue & Gold crossover (Mar 11), leaving a make-up night in the schedule.',
    'Troop Visits apply only to the Arrow of Light den and are arranged directly by the AOL Den Leader with local Scouts BSA troops on their own timetable (not on the top calendar page).',
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
    ['Aug 13', 'Welcome; introduce the Scout Oath & Law; active games', 'Roper Park (Playground)'],
    ['Aug 20', 'Bobcat: sign, salute, handshake, motto, slogan; outdoor games', 'Roper Park (Playground)'],
    ['Aug 27', 'Bobcat: Cub Scout Six Essentials & buddy system; nature walk', 'Roper Park (Playground)'],
    ['Sep 3', 'Bobcat: personal safety / Protect Yourself Rules (age-appropriate)', 'The Chapel'],
    ['Sep 10', 'Pack-wide elective: Champions for Nature clean-up; Sign-Up Night push', 'Roper Park (Playground)'],
    ['Sep 17', 'Pack-wide elective: STEM/craft activity', 'Fellowship Hall'],
    ['Sep 24', 'Welcome SUN recruits; elective activity; prep to split into dens', 'Fellowship Hall'],
  ].forEach(([dt, ag, loc]) => blocks.push(para([run(dt + ' — ', { bold: true }), run(`${ag} (${loc})`)])));

  const aolCutoff = d(2026, 3, 5);
  for (const den of DENS) {
    const dates = den.crossesOver ? denMeetingDates.filter((x) => x < aolCutoff) : denMeetingDates;
    const rows = expandPlan(dates, DEN_PLANS[den.rank]);
    blocks.push(heading(`${den.rank} Den (${den.grade})`, 2));
    if (den.crossesOver) {
      blocks.push(para([run(
        'Crosses over to Scouts BSA at the Blue & Gold Banquet (Mar 11), so all required work is front-loaded into Oct–Mar. '
        + 'Race Time (2nd elective) is earned through the Pinewood Derby build nights (Feb 4/11/18) and race (Feb 21–22).'
      )]));
    }
    for (const r of rows) {
      blocks.push(para([run(fmtShort(r.date) + ' — ', { bold: true }), run(`${r.name}: ${r.agenda} (${r.loc})`)]));
    }
  }

  blocks.push(heading('Notes to verify before publishing', 2));
  [
    'A few required-adventure names vary across sources (e.g., Wolf/Bear "Safe and Smart" vs "Safety in Numbers", Bear "Standing Tall"). Confirm against the current official rank pages.',
    'BALOO Training date (Mar 28–29) is estimated from the Mar 24 registration close — confirm on the council flyer.',
    'Frontier Day is listed by the council for Sep 19, 2026 (start of the next program year); confirm whether a spring-2026 date exists.',
    'Winter Campout and BUC TUOCS dates are placeholders for the committee to finalize.',
  ].forEach((t) => blocks.push(para([run('• ' + t)])));

  return blocks;
}

// ---------------------------------------------------------------------------
function main() {
  const wednesdays = buildWednesdayEvents();
  const committee = buildCommitteeEvents();
  const special = buildSpecialEvents();
  const allEvents = [...wednesdays, ...committee, ...special];

  const denMeetingDates = wednesdays
    .filter((e) => e.kind === 'den')
    .map((e) => e.date)
    .sort((a, b) => a - b);

  // Section 1: full-width title + legend.
  const titleBlocks = [
    para([run(TITLE)], { style: 'Title' }),
    para([
      run('Cub Scout Pack 127', { bold: true }),
      run('  •  Regular meetings Wednesdays at 6pm  •  '),
      run('FH', { bold: true }), run(' = Fellowship Hall  •  '),
      run('OFF', { bold: true }), run(' = no meeting (school closed)'),
    ], { spaceAfter: 120 }),
  ];

  // Section 2: two-column flowing calendar.
  const calendarBlocks = renderCalendarBody(allEvents);

  // Section 3: single-column program recommendations (new page).
  const programBlocks = buildProgramBlocks(denMeetingDates);

  writeDocx(OUT, {
    title: TITLE,
    sections: [
      { blocks: titleBlocks, numCols: 1, type: 'continuous' },
      { blocks: calendarBlocks, numCols: 2, type: 'continuous' },
      { blocks: programBlocks, numCols: 1, type: 'nextPage' },
    ],
  });
  console.log('Wrote', OUT);
}

main();
