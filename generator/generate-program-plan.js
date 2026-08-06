// generate-program-plan.js — builds Program_Plan.docx from the ACTUAL published
// calendar (../calendar.md), NOT from spec.md.
//
//   node generator/generate-program-plan.js
//
// Two independent triggers live in this repo (see README.md / CLAUDE.md):
//
//   spec.md      --> generator/generate.js              --> "<year> Pack 127 Calendar.docx"
//   calendar.md  --> generator/generate-program-plan.js --> "Program_Plan.docx"
//
// spec.md describes how to BUILD a calendar; calendar.md is a calendar that has
// already been built and published. This generator reads the real meeting
// schedule out of calendar.md and lays the per-den program recommendations
// (the rules in spec.md's "Program Recommendations" section) on top of the real
// dates, so the plan always matches the calendar the pack is actually using.
//
// Reuses the stable writer (docx.js) and helpers (lib.js) — neither is modified.

const fs = require('fs');
const path = require('path');
const { run, para, heading, writeDocx } = require('./docx');
const { MONTH_ABBR, WEEKDAY, ymd, nthWeekdayOfMonth } = require('./lib');

const CALENDAR_MD = path.join(__dirname, '..', 'calendar.md');
const OUT = path.join(__dirname, '..', 'Program_Plan.docx');

// ---------------------------------------------------------------------------
// Parse the published calendar (calendar.md) into the schedule the program plan
// needs: the ordered list of den-meeting Wednesdays and the Blue & Gold / Arrow
// of Light crossover date. calendar.md is structured markdown (a bold date token
// leads each event line), so unlike spec.md it can be parsed directly.
const MONTH_NAMES = {
  January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
  July: 6, August: 7, September: 8, October: 9, November: 10, December: 11,
};
const MONTH_ABBR_MAP = {
  Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6,
  Aug: 7, Sep: 8, Sept: 8, Oct: 9, Nov: 10, Dec: 11,
};

function parseCalendar(md) {
  const lines = md.split(/\r?\n/);
  const monthHeader = /^\*\*([A-Z][a-z]+)\s+(\d{4})\*\*\s*$/;
  const firstBold = /\*\*([^*]+)\*\*/;

  let curMonth = null;
  let curYear = null;
  const denDates = [];
  let crossover = null;

  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;

    // Month header (e.g. "**September 2026**") sets the running month/year.
    const mh = line.match(monthHeader);
    if (mh && MONTH_NAMES[mh[1]] !== undefined) {
      curMonth = MONTH_NAMES[mh[1]];
      curYear = parseInt(mh[2], 10);
      continue;
    }

    // Event line: needs a bold date token somewhere.
    const bm = line.match(firstBold);
    if (!bm) continue;
    if (curMonth === null || curYear === null) continue;

    // Resolve the event's month/day from the bold token. Strip markdown escapes
    // (\& \-), clock times (1:00), and any 4-digit year before hunting the day.
    let scan = bm[1].replace(/\\/g, ' ').replace(/\d{1,2}:\d{2}/g, ' ').replace(/\b\d{4}\b/g, ' ');
    const monFull = scan.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/);
    const monAbbr = scan.match(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sept|Sep|Oct|Nov|Dec)\b/);
    let evMonth = curMonth;
    if (monFull) evMonth = MONTH_NAMES[monFull[1]];
    else if (monAbbr) evMonth = MONTH_ABBR_MAP[monAbbr[1]];
    const dayMatch = scan.match(/\b(\d{1,2})\b/);
    if (!dayMatch) continue;
    const day = parseInt(dayMatch[1], 10);
    const dt = new Date(curYear, evMonth, day, 12, 0, 0);

    // Classify on the whole line (some markers live inside the bold token).
    const text = line.replace(/\\/g, '');
    if (/Crossover/i.test(text) || /B&G/i.test(text)) {
      if (!crossover) crossover = dt;
    }
    if (/Den Meeting/i.test(text)) denDates.push(dt);
  }

  denDates.sort((a, b) => a - b);
  if (!denDates.length) throw new Error('No "Den Meeting" entries found in calendar.md.');
  if (!crossover) throw new Error('No Blue & Gold / crossover entry found in calendar.md.');
  return { denDates, crossover };
}

// ---------------------------------------------------------------------------
// Daylight windows for the 6pm playground (Roper Park): available before the
// clocks fall back (first Sunday of November) and again after they spring
// forward (second Sunday of March). Derived from the school year in calendar.md.
function daylightBounds(startYear) {
  return {
    fallBack: nthWeekdayOfMonth(startYear, 11, 0, 1),       // 1st Sunday of Nov
    springForward: nthWeekdayOfMonth(startYear + 1, 3, 0, 2), // 2nd Sunday of Mar
  };
}

// ---------------------------------------------------------------------------
// Dens. Each non-AOL den has all den meetings; the AOL den finishes at
// crossover. Adventure blocks are [name, agenda, location, count] and are
// ordered so a positional fill against the dates keeps outdoor work in the
// daylight windows and the spring-only adventures (Bear Whittling, Webelos
// Chef's Knife) in the spring.
//
// This plan is tuned to the CURRENT calendar.md shape: 6 fall / 9 indoor / 7
// spring den meetings (22 total) and 12 AOL meetings (6 fall / 6 indoor) before
// crossover. assertShape() below verifies calendar.md still has that shape; if a
// future calendar changes it, the build fails loudly so the blocks get re-tuned
// (in CI, Claude Code reconciles this file — see the workflow).
const SHAPE = { fall: 6, indoor: 9, spring: 7, aol: 12, aolFall: 6 };

const DENS = [
  { rank: 'Lion', grade: 'Kindergarten', crossesOver: false },
  { rank: 'Tiger', grade: '1st grade', crossesOver: false },
  { rank: 'Wolf', grade: '2nd grade', crossesOver: false },
  { rank: 'Bear', grade: '3rd grade', crossesOver: false },
  { rank: 'Webelos', grade: '4th grade', crossesOver: false },
  { rank: 'Arrow of Light', grade: '5th grade', crossesOver: true },
];

const DEN_PLANS = {
  Lion: [
    // Fall (6) — the 3 September meetings fall during recruiting, so they lead
    // with electives a late-joining scout can pick up; required work starts Oct.
    ['On Your Mark (elective)', 'races & fitness challenges', 'Roper Park (Playground)', 2],
    ['Ready, Set, Grow (elective)', 'plant seeds & explore growing things', 'Roper Park (Playground)', 1],
    ['Mountain Lion', 'nature walk & outdoor exploration', 'Roper Park (Playground)', 3],
    // Indoor (9)
    ['King of the Jungle', 'animal-movement games (indoor)', 'Fellowship Hall', 1],
    ["Lion's Pride", 'den identity & family activity', 'Bethany Room (Bell Hall)', 2],
    ["Lion's Roar", 'recognize & respond to danger', 'The Chapel', 2],
    ['Fun on the Run', 'healthy habits + indoor movement game', 'Fellowship Hall', 2],
    ['Build It Up (elective)', 'simple indoor building craft', 'Upstairs (Epworth Hall)', 2],
    // Spring (7)
    ['Count on Me (elective)', 'counting & STEM games outdoors', 'Roper Park (Playground)', 2],
    ['On Your Mark wrap-up (elective)', 'outdoor races & fitness', 'Roper Park (Playground)', 2],
    ['Year-end games & bridge to Tiger', 'field games; awards prep', 'Roper Park (Playground)', 3],
  ],
  Tiger: [
    // Fall (6)
    ['Team Tiger (elective start)', 'active outdoor teamwork games', 'Roper Park (Playground)', 2],
    ['Sky Is the Limit (elective)', 'outdoor sky & weather observation', 'Roper Park (Playground)', 1],
    ['Tigers in the Wild', 'hike & nature exploration', 'Roper Park (Playground)', 3],
    // Indoor (9)
    ['Tiger Bites', 'healthy foods & a simple snack', 'Fellowship Hall', 2],
    ['Team Tiger', 'teamwork & helping at home', 'Bethany Room (Bell Hall)', 1],
    ['Tiger Circles', 'duty to God & values', 'The Chapel', 2],
    ["Tiger's Roar", 'weather safety & calling for help', 'Upstairs (Epworth Hall)', 2],
    ['Stories and Shapes (elective)', 'drawing & storytelling', 'Upstairs (Epworth Hall)', 2],
    // Spring (7)
    ['Floats and Boats (elective)', 'build & test simple boats outdoors', 'Roper Park (Playground)', 2],
    ['Tigers in the Wild wrap-up', 'outdoor picnic games & nature', 'Roper Park (Playground)', 2],
    ['Year-end games & bridge to Wolf', 'field games; awards prep', 'Roper Park (Playground)', 3],
  ],
  Wolf: [
    // Fall (6)
    ['Air of the Wolf (elective)', 'build & launch balloon rockets outdoors', 'Roper Park (Playground)', 2],
    ['Running with the Pack', 'fitness & active games', 'Roper Park (Playground)', 1],
    ['Paws on the Path', 'short hike & outdoor essentials', 'Roper Park (Playground)', 3],
    // Indoor (9)
    ['Council Fire (Duty to Country)', 'flags & community', 'The Chapel', 2],
    ['Footsteps (Duty to God)', 'values & family faith', 'The Chapel', 2],
    ['Safe and Smart', 'personal & internet safety', 'Upstairs (Epworth Hall)', 2],
    ['Code of the Wolf (elective)', 'math & simple coding games', 'Upstairs (Epworth Hall)', 2],
    ['Germs Alive! (elective)', 'science experiment (indoor)', 'Fellowship Hall', 1],
    // Spring (7)
    ['Adventures in Coins (elective)', 'outdoor coin & history hunt', 'Roper Park (Playground)', 2],
    ['Paws on the Path wrap-up', 'nature hike & outdoor skills', 'Roper Park (Playground)', 2],
    ['Year-end games & bridge to Bear', 'field games; awards prep', 'Roper Park (Playground)', 3],
  ],
  Bear: [
    // Fall (6)
    ['Champions for Nature (elective)', 'outdoor conservation service', 'Roper Park (Playground)', 2],
    ['Bear Strong', 'fitness & active outdoor games', 'Roper Park (Playground)', 1],
    ['Bear Habitat', 'study trees & nature outdoors', 'Roper Park (Playground)', 3],
    // Indoor (9) — Baloo the Builder (tools) in the Cabin, per constraint 7
    ['Paws for Action (Duty to Country)', 'citizenship & a service project', 'The Chapel', 2],
    ['Fellowship (Duty to God)', 'values & reflection', 'The Chapel', 2],
    ['Standing Tall', 'character & making good choices', 'Upstairs (Epworth Hall)', 2],
    ['Baloo the Builder (elective)', 'woodworking with hand tools', 'The Cabin', 3],
    // Spring (7) — Whittling in the spring, per constraint 7
    ['Whittling (elective, Spring)', 'pocketknife safety & whittling', 'The Cabin', 3],
    ['Bear Habitat wrap-up', 'outdoor nature study & hike', 'Roper Park (Playground)', 2],
    ['Year-end games & bridge to Webelos', 'field games; awards prep', 'Roper Park (Playground)', 2],
  ],
  Webelos: [
    // Fall (6)
    ['Earth Rocks! (elective)', 'outdoor geology exploration', 'Roper Park (Playground)', 2],
    ['Stronger, Faster, Higher', 'fitness & team sports outdoors', 'Roper Park (Playground)', 1],
    ['Webelos Walkabout', 'plan & take a hike', 'Roper Park (Playground)', 3],
    // Indoor (9)
    ['My Community', 'community roles & service', 'Bethany Room (Bell Hall)', 2],
    ['My Family', 'family duties & budgeting', 'Bethany Room (Bell Hall)', 2],
    ['My Safety', 'emergency preparedness', 'Upstairs (Epworth Hall)', 2],
    ['Build It (elective)', 'construction & engineering project', 'The Cabin', 3],
    // Spring (7) — Chef's Knife in the spring, per constraint 6
    ["Chef's Knife (elective, Spring)", 'kitchen-knife safety & cooking', 'Fellowship Hall (kitchen)', 3],
    ['Aware and Care (elective)', 'accessibility awareness activity outdoors', 'Roper Park (Playground)', 1],
    ['Year-end games & prep for Arrow of Light', 'field games; awards prep', 'Roper Park (Playground)', 3],
  ],
  // Arrow of Light — crosses over at Blue & Gold, so all required work is
  // front-loaded before the crossover date. Two electives: Knife Safety +
  // Race Time (Pinewood Derby build nights & race). Constraint 5.
  'Arrow of Light': [
    // Fall (6)
    ['Personal Fitness', 'fitness testing & active games', 'Roper Park (Playground)', 2],
    ['Outdoor Adventurer', 'outdoor skills & Fall Campout prep', 'Roper Park (Playground)', 2],
    ['Scouting Adventure', 'visit / open house with a Scouts BSA troop (125, 84, 21)', 'Roper Park (Playground)', 2],
    // Indoor (6)
    ['Building a Better World (Citizenship)', 'community service & flag ceremony', 'The Chapel', 1],
    ['Duty to God in Action', 'values & family faith', 'The Chapel', 1],
    ['First Aid', 'first-aid skills & building a kit', 'Upstairs (Epworth Hall)', 1],
    ['Knife Safety (elective)', 'pocket + kitchen knife safety', 'The Cabin', 1],
    ['Race Time (elective) — Pinewood Derby', 'car build carried by the Pinewood Derby Build Nights & race', 'Fellowship Hall', 1],
    ['Crossover prep & rank make-up', 'finish requirements; ready for Blue & Gold', 'The Chapel', 1],
  ],
};

// ---------------------------------------------------------------------------
// Verify the parsed calendar still matches the shape the DEN_PLANS assume.
function assertShape(sched, bounds) {
  for (const dt of sched.denDates) {
    if (dt.getDay() !== 3) {
      throw new Error(`Den meeting ${ymd(dt)} is a ${WEEKDAY[dt.getDay()]}, not a Wednesday.`);
    }
  }
  const fall = sched.denDates.filter((x) => x < bounds.fallBack).length;
  const spring = sched.denDates.filter((x) => x > bounds.springForward).length;
  const indoor = sched.denDates.length - fall - spring;
  const aol = sched.denDates.filter((x) => x < sched.crossover).length;
  const aolFall = sched.denDates.filter((x) => x < sched.crossover && x < bounds.fallBack).length;
  if (fall !== SHAPE.fall || indoor !== SHAPE.indoor || spring !== SHAPE.spring
      || aol !== SHAPE.aol || aolFall !== SHAPE.aolFall) {
    throw new Error(
      `calendar.md shape changed: fall=${fall} indoor=${indoor} spring=${spring} `
      + `aol=${aol} aolFall=${aolFall}; expected ${JSON.stringify(SHAPE)}. `
      + 'Re-tune DEN_PLANS in generator/generate-program-plan.js to the new shape.');
  }
}

function fmtShort(dt) { return `${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`; }

// Positional fill: pair each block's meetings with successive dates.
function expandPlan(dates, blocks) {
  const rows = [];
  let i = 0;
  for (const [name, agenda, loc, n] of blocks) {
    for (let k = 0; k < n && i < dates.length; k++, i++) {
      rows.push({ date: dates[i], name, agenda, loc });
    }
  }
  if (i !== dates.length) {
    throw new Error(`Plan/date mismatch: ${dates.length} dates but blocks fill ${i}.`);
  }
  return rows;
}

// ---------------------------------------------------------------------------
function buildBlocks(sched) {
  const schoolStart = sched.denDates[0].getFullYear();
  const schoolYear = `${schoolStart}-${schoolStart + 1}`;
  const title = `${schoolYear} Pack 127 Program Plan`;
  const crossoverLong = `${WEEKDAY[sched.crossover.getDay()]} ${fmtShort(sched.crossover)}`;

  const blocks = [];

  blocks.push(para([run(title)], { style: 'Title' }));
  blocks.push(para([
    run('Cub Scout Pack 127', { bold: true }),
    run('  •  Program recommendations for each den at each meeting.'),
  ], { spaceAfter: 120 }));
  blocks.push(para([run(
    'This plan follows the program rules in spec.md but is keyed to the actual published '
    + 'calendar (calendar.md), not a freshly generated schedule. The meeting dates, the den/pack '
    + `split, and the Blue & Gold / Arrow of Light crossover (${crossoverLong}) are read directly out `
    + 'of calendar.md. Adventures follow the current Cub Scout program '
    + '(scouting.org/programs/cub-scouts/adventures): each rank earns its badge with 6 required adventures '
    + '(including Bobcat) plus at least 2 electives. Keep active program time to 30–40 minutes per meeting; '
    + 'the rest is opening, snack, and closing.'
  )]));

  blocks.push(heading('Guiding constraints', 2));
  [
    'Bobcat is done first, pack-wide, during the August–September pack-wide meetings before the dens split.',
    'September den meetings fall while Sign-Up Nights are still running, so they stay on elective / outdoor work that a late-joining scout can pick up — no required advancement is scheduled until October.',
    'Include one moving activity every den meeting. Webelos and Arrow of Light can handle desk work; the younger dens need to move.',
    'Roper Park (the playground) has daylight at 6pm only before the clocks fall back (first Sunday of November) and again after they spring forward (second Sunday of March). Outdoor/fitness adventures are scheduled in those windows; winter meetings move indoors.',
    'Anything with tools (Baloo the Builder, Whittling) works best in the Cabin; anything with cooking/food (Chef’s Knife, Tiger Bites) works best in Fellowship Hall (kitchen).',
    'Bear does Baloo the Builder and Whittling, with Whittling in the Spring. Webelos does Chef’s Knife in the Spring.',
    'Arrow of Light earns two electives: Knife Safety (a den meeting) and Race Time (covered by the Pinewood Derby build nights and race).',
    `Arrow of Light finishes all required rank work before the Blue & Gold crossover (${crossoverLong}), leaving a make-up night in the schedule.`,
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
  blocks.push(para([run(
    'Before the dens split, the pack meets together on the pack-wide Wednesdays in calendar.md. '
    + 'Use them for Bobcat (sign, salute, handshake, motto, slogan, the Cub Scout Six Essentials, '
    + 'the buddy system, and personal safety / Protect Yourself Rules) plus pack-wide elective games, '
    + 'so every scout — including those who join at a Sign-Up Night — completes Bobcat together.'
  )]));

  // Per-den plans.
  for (const den of DENS) {
    const dates = den.crossesOver
      ? sched.denDates.filter((x) => x < sched.crossover)
      : sched.denDates;
    const rows = expandPlan(dates, DEN_PLANS[den.rank]);
    blocks.push(heading(`${den.rank} Den (${den.grade})`, 2));
    if (den.crossesOver) {
      blocks.push(para([run(
        `Crosses over to Scouts BSA at the Blue & Gold Banquet (${crossoverLong}), so all required work is front-loaded before it. `
        + 'Race Time (2nd elective) is earned through the Pinewood Derby build nights and race.'
      )]));
    }
    for (const r of rows) {
      blocks.push(para([run(fmtShort(r.date) + ' — ', { bold: true }), run(`${r.name}: ${r.agenda} (${r.loc})`)]));
    }
  }

  blocks.push(heading('Program notes', 2));
  [
    'Dates and meeting types are read directly from the published calendar.md; verification of externally-scheduled dates is tracked in that calendar’s "Issues — Unverified Dates" list, not repeated here.',
    'A few required-adventure names vary across sources (e.g., Wolf/Bear "Safe and Smart" vs "Safety in Numbers", Bear "Standing Tall"). Confirm against the current official rank pages.',
    'Back to the Pack (the August kickoff) is already scheduled on the published calendar, with its date and venue; the pack-wide Bobcat meetings above run on the August–September Wednesdays that follow it.',
    'These recommendations are encouraged, not required — den leaders may reorder adventures to fit their scouts, as long as tool/cooking work stays in the Cabin/kitchen and outdoor work stays in the daylight windows.',
  ].forEach((t) => blocks.push(para([run('• ' + t)])));

  return { blocks, title };
}

function main() {
  const md = fs.readFileSync(CALENDAR_MD, 'utf8');
  const sched = parseCalendar(md);
  const bounds = daylightBounds(sched.denDates[0].getFullYear());
  assertShape(sched, bounds);

  const { blocks, title } = buildBlocks(sched);
  writeDocx(OUT, {
    title,
    sections: [{ blocks, numCols: 1, type: 'continuous' }],
  });
  console.log('Wrote', OUT, `(${sched.denDates.length} den meetings, crossover ${ymd(sched.crossover)})`);
}

main();
