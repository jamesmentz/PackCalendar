# PackCalendar

Generates the **Cub Scout Pack 127** annual calendar as a Word document, per
[`spec.md`](spec.md).

## Output
- **`2026-2027 Pack 127 Calendar.docx`** — the finished calendar:
  - a two-column, page-flowing calendar (Aug 2026 → the following summer) with
    Pack Meetings, Den Meetings, Committee Meetings, Special Events, Sign-Up
    Nights, council/district events, and OFF weeks;
  - program-recommendation pages with a suggested adventure, agenda, and location
    for every den at every meeting.

(The prior `2025-2026 Pack 127 Calendar.docx` is kept for reference.)

## How regeneration works
`spec.md` is the **single source of truth**, but it is prose — so no script
parses it directly. Instead, editing `spec.md` and pushing triggers the
**Regenerate Calendar** GitHub Action, which runs Claude Code to reconcile
`generator/generate.js` with the new spec, then runs the deterministic generator
to rebuild the `.docx` and commits it. (Requires a `CLAUDE_CODE_OAUTH_TOKEN`
repo secret — from `claude setup-token` — or an `ANTHROPIC_API_KEY`; see
`.github/workflows/regenerate-calendar.yml` and `CLAUDE.md`.)

To do the same locally with Claude Code, ask it to "regenerate the calendar from
spec.md" — it follows the regeneration contract in `CLAUDE.md`.

## Generator
`generator/` builds the `.docx` with no external dependencies — only Node.js and
(on Windows) PowerShell/.NET for zip packaging. No Word, Python, or npm packages.

```
node generator/generate.js
```

| File | Purpose |
|---|---|
| `generator/generate.js` | Schedule data + logic; entry point. Kept in sync with `spec.md` (see `CLAUDE.md`). |
| `generator/lib.js` | Date helpers and the two-column calendar renderer. |
| `generator/docx.js` | Minimal OOXML → `.docx` writer (stable). |
| `generator/DATA_SOURCES.md` | Where every date/fact came from and which items still need verifying. |

See `generator/DATA_SOURCES.md` for the ACPS calendar, council event, and Cub
Scout adventure sources, plus the flagged items to confirm before distributing.
