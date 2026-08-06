# PackCalendar

Generates two Word documents for **Cub Scout Pack 127**: the annual **calendar**
(from [`spec.md`](spec.md)) and a **program plan** (from the actual published
[`calendar.md`](calendar.md)).

## Two triggers, two outputs

This repo has **two independent pipelines**. They start from different source
files, run different generators, and produce different documents:

| Source (edit this) | Generator | Output |
|---|---|---|
| `spec.md` — *how to build* the calendar (prose policy) | `generator/generate.js` | **`<year> Pack 127 Calendar.docx`** |
| `calendar.md` — *the calendar already built and published* | `generator/generate-program-plan.js` | **`Program_Plan.docx`** |

- **`<year> Pack 127 Calendar.docx`** — the finished calendar:
  - a two-column, page-flowing calendar (Aug → the following summer) with Pack
    Meetings, Den Meetings, Committee Meetings, Special Events, Sign-Up Nights,
    council/district events, and OFF weeks;
  - program-recommendation pages with a suggested adventure, agenda, and location
    for every den at every meeting.
- **`Program_Plan.docx`** — a standalone program plan whose per-den, per-meeting
  adventure recommendations are laid over the **real dates** in `calendar.md`
  (so it matches the schedule the pack is actually running, including the real
  Blue & Gold crossover date), rather than a freshly generated schedule.

(The prior `2025-2026 Pack 127 Calendar.docx` is kept for reference.)

## How regeneration works

**Calendar — trigger: `spec.md`.** `spec.md` is the **single source of truth**
for the calendar, but it is prose — so no script parses it directly. Editing
`spec.md` and pushing triggers the **Regenerate Calendar** GitHub Action
(`.github/workflows/regenerate-calendar.yml`), which runs Claude Code to
reconcile `generator/generate.js` with the new spec, then runs the deterministic
generator to rebuild the `.docx` and commits it.

**Program plan — trigger: `calendar.md`.** `calendar.md` is the actual published
calendar. Editing it and pushing triggers the **Regenerate Program Plan** GitHub
Action (`.github/workflows/regenerate-program-plan.yml`). The generator parses
the meeting dates straight out of `calendar.md`; if the calendar changes shape,
Claude Code reconciles the per-den blocks in `generator/generate-program-plan.js`
first, then the generator rebuilds `Program_Plan.docx` and commits it.

Both Actions require a `CLAUDE_CODE_OAUTH_TOKEN` repo secret — from
`claude setup-token` — or an `ANTHROPIC_API_KEY`; see the workflow files and
`CLAUDE.md`.

To do either locally with Claude Code, ask it to "regenerate the calendar from
spec.md" or "regenerate the program plan from calendar.md" — it follows the
regeneration contract in `CLAUDE.md`.

## Generator
`generator/` builds the `.docx` with no external dependencies — only Node.js.
`docx.js` writes the ZIP itself, deterministically (byte-identical output on
every platform/Node version). No Word, Python, npm packages, or external tools.

```
node generator/generate.js               # calendar   (reads spec.md via generate.js data)
node generator/generate-program-plan.js  # program plan (reads calendar.md)
```

| File | Purpose |
|---|---|
| `generator/generate.js` | Calendar schedule data + logic; entry point. Kept in sync with `spec.md` (see `CLAUDE.md`). |
| `generator/generate-program-plan.js` | Parses `calendar.md` and renders `Program_Plan.docx`; entry point for the program-plan pipeline. |
| `generator/lib.js` | Date helpers and the two-column calendar renderer (shared). |
| `generator/docx.js` | Minimal OOXML → `.docx` writer (stable, shared). |
| `generator/DATA_SOURCES.md` | Where every date/fact came from and which items still need verifying. |

See `generator/DATA_SOURCES.md` for the ACPS calendar, council event, and Cub
Scout adventure sources, plus the flagged items to confirm before distributing.
