# CLAUDE.md

Guidance for Claude Code when working in this repository. These instructions
also drive the CI regeneration job (`.github/workflows/regenerate-calendar.yml`).

## What this repo is

It generates the **Cub Scout Pack 127** annual calendar as a Word document
(`YYYY-YYYY Pack 127 Calendar.docx`) for a school year and the following summer.

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
   - `DENS`, `DEN_PLANS`, and `buildProgramBlocks` (program recommendations).
   Edit `generator/lib.js` only if a change genuinely needs new rendering/date
   logic.
3. **Do not modify `generator/docx.js`** — the minimal OOXML `.docx` writer is
   stable and correct. It should never need changes to satisfy a spec edit.
4. Run `node generator/generate.js` to rebuild the `.docx`.
5. Be faithful to `spec.md`. Do not invent dates the spec does not support; when
   the spec says a date is TBD/TBA, reflect that rather than guessing. Follow
   the conventions and open questions in `generator/DATA_SOURCES.md`.

In CI, committing/pushing is handled by a dedicated workflow step — the agent
should only edit files and run the generator, not run git.

## Files

| File | Purpose |
|---|---|
| `spec.md` | Source of truth (prose). Edit this to change the calendar. |
| `generator/generate.js` | Schedule data + logic; entry point (`node generator/generate.js`). |
| `generator/lib.js` | Date helpers and the two-column calendar renderer. |
| `generator/docx.js` | Minimal OOXML → `.docx` writer. **Stable — do not edit for spec changes.** |
| `generator/DATA_SOURCES.md` | Where each date/fact came from; items to verify. |

## Running locally

```
node generator/generate.js
```

Node core only — no npm packages. On Windows, `docx.js` shells out to
PowerShell/.NET for zip packaging.
