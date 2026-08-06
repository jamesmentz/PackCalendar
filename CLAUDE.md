# CLAUDE.md

Guidance for Claude Code when working in this repository. These instructions
also drive the CI regeneration job (`.github/workflows/regenerate-calendar.yml`).

## What this repo is

It generates two Word documents for **Cub Scout Pack 127**, via two independent
pipelines with two different triggers:

| Trigger (source) | Generator | Output |
|---|---|---|
| `spec.md` — how to *build* the calendar (prose) | `generator/generate.js` | `YYYY-YYYY Pack 127 Calendar.docx` |
| `calendar.md` — the calendar already *built & published* | `generator/generate-program-plan.js` | `Program_Plan.docx` |

The calendar covers a school year and the following summer. The program plan
lays the per-den program recommendations over the real dates in `calendar.md`.
The two regeneration contracts are below — do not cross the wires: a `spec.md`
edit rebuilds the calendar; a `calendar.md` edit rebuilds the program plan.

## Source of truth

- **`spec.md` is the single source of truth.** It is prose: it describes the
  policies (meeting cadence, off weeks, committee-meeting rules), the specific
  dates, the special-event list, and the per-den program recommendations.
- **`generator/` renders the `.docx` deterministically** from structured data.
  Because `spec.md` is prose, no script parses it directly. Instead, whenever
  `spec.md` changes, Claude Code reconciles the generator with the spec and then
  the generator is run. This is the regeneration contract below.

## Regeneration contract (what CI runs, and what to do locally)

When `spec.md` changes:

1. Read `spec.md` in full.
2. Edit **`generator/generate.js`** so all data and scheduling logic match the
   spec. This is where everything configurable lives:
   - `SCHOOL_YEAR`, `FIRST_DAY`, `LAST_DAY`, `OFF_WEDNESDAYS`.
   - Wednesday classification constants (`PACK_MEETING_WED`, `BUILD_NIGHT_WED`,
     `BLUE_GOLD_WED`, `BACK_TO_PACK_WED`, `YEAR_END_PACK_WED`) and
     `classifyWednesday`.
   - `buildCommitteeEvents` (last-Thursday rule + holiday-week moves).
   - `buildSpecialEvents` (campouts, council/district events, Sign-Up Nights).
     Add a `verify: '...'` field to any event whose date you could not confirm;
     that pushes it onto the bottom-of-document "Issues — Unverified Dates" list.
   - `buildConfigIssues` (unverified config dates — ACPS start/last day, the
     Veterans Day OFF assumption, the January return, TBD dates) — also feeds the
     Issues list.
   - `DENS`, `DEN_PLANS`, and `buildProgramBlocks` (program recommendations).
   Edit `generator/lib.js` only if a change genuinely needs new rendering/date
   logic.
   Verify every date you can against its authoritative source before adding a
   `verify` note; the remaining unconfirmed dates become the Issues list. Do not
   hand-write a weekday or weekend range that disagrees with the real calendar —
   `assertDateLabels` fails the build on any weekday/start-date mismatch, because
   day-of-week and calendar arithmetic are known and must always be correct.
3. **Do not modify `generator/docx.js`** — the minimal OOXML `.docx` writer is
   stable and correct. It should never need changes to satisfy a spec edit.
4. Run `node generator/generate.js` to rebuild the `.docx`.
5. Be faithful to `spec.md`. Do not invent dates the spec does not support; when
   the spec says a date is TBD/TBA, reflect that rather than guessing. Follow
   the conventions and open questions in `generator/DATA_SOURCES.md`.

In CI, committing/pushing is handled by a dedicated workflow step — the agent
should only edit files and run the generator, not run git.

## Program-plan regeneration contract (trigger: `calendar.md`)

`calendar.md` is the **actual published calendar** — a calendar that has already
been built (by the `spec.md` pipeline or by hand) and distributed. It is the
source of truth for the **program plan**, and unlike `spec.md` it is structured
enough that a script parses it directly.

When `calendar.md` changes:

1. Read `calendar.md` in full and read `generator/generate-program-plan.js`.
2. The generator parses the den-meeting Wednesdays and the Blue & Gold /
   Arrow-of-Light crossover date straight out of `calendar.md`, then lays the
   per-den adventure blocks (`DEN_PLANS`) over those dates. The dates are always
   correct to the committed calendar; what needs reconciling is the program
   *content*. If the calendar changes shape — a different number of den meetings,
   a different fall/indoor/spring daylight split, or a different crossover — the
   `SHAPE` assertion fails on purpose. Re-tune `SHAPE` and the `DEN_PLANS` block
   counts so each den's blocks fill exactly its meeting dates, outdoor work stays
   in the daylight windows, and the spring-only adventures (Bear Whittling,
   Webelos Chef's Knife) stay in the spring.
3. **Do not modify `generator/docx.js` or `generator/generate.js`** — the program
   plan is a separate pipeline. Edit only `generator/generate-program-plan.js`
   (and `generator/lib.js` only if a change genuinely needs new shared logic).
4. Run `node generator/generate-program-plan.js` to rebuild `Program_Plan.docx`.
5. Be faithful to the program rules in `spec.md`; do not invent dates
   `calendar.md` does not contain.

## Files

| File | Purpose |
|---|---|
| `spec.md` | Source of truth for the **calendar** (prose). Edit to change the calendar. |
| `calendar.md` | The actual published calendar; source of truth for the **program plan**. |
| `generator/generate.js` | Calendar schedule data + logic; entry point (`node generator/generate.js`). |
| `generator/generate-program-plan.js` | Parses `calendar.md` → `Program_Plan.docx`; entry point (`node generator/generate-program-plan.js`). |
| `generator/lib.js` | Date helpers and the two-column calendar renderer (shared). |
| `generator/docx.js` | Minimal OOXML → `.docx` writer (shared). **Stable — do not edit for spec/calendar changes.** |
| `generator/DATA_SOURCES.md` | Where each date/fact came from; items to verify. |

## Running locally

```
node generator/generate.js               # calendar    (from generate.js data / spec.md)
node generator/generate-program-plan.js  # program plan (parses calendar.md)
```

Node core only — no npm packages and no external tools. `docx.js` writes the
`.docx` ZIP itself, and its output is byte-deterministic across platforms and
Node versions (so CI regenerating produces no spurious diff when nothing
substantive changed).
