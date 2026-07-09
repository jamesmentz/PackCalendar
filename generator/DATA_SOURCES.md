# Data sources & assumptions — 2025-2026 Pack 127 Calendar

The generator (`generator/generate.js`) builds the calendar from data hard-coded
in that file. This document records where each piece of data came from and the
judgment calls made where the spec was silent. **Verify the flagged items before
distributing the calendar.**

## School calendar (ACPS 2025-2026)
Source: Alachua County Public Schools approved 2025-2026 calendar
(https://www.alachuaschools.net/o/acps/page/calendars — approved 02/18/2025).

- First student day: Mon Aug 11, 2025. Last student day: Tue Jun 2, 2026.
- Only four Wednesdays fall inside a student closure and are marked **OFF**:
  Nov 26 (Thanksgiving), Dec 24 & Dec 31 (winter break), Mar 18 (spring break).
- No single-day holiday lands on a Wednesday, so no other meetings are skipped.
- Daylight (for the 6pm playground): clocks fall back **Sun Nov 2, 2025** and
  spring forward **Sun Mar 8, 2026** — the outdoor-scheduling windows.

## Meeting structure (from spec.md)
- Every school-year Wednesday at 6pm is a Pack or Den meeting unless it is a
  Special Event or an OFF week.
- **Aug & Sep** meetings are pack-wide (elective work) because Sign-Up Nights
  have not happened yet.
- **Pack Meetings** (≤ 1/month): 1st Wednesday of Oct, Nov, Dec, Jan, Apr; a
  year-end celebration on the last Wednesday of May. Page 1 shows `(FH, 6pm)`;
  pack meetings require Fellowship Hall (the only room big enough).
- **February** carries three Pinewood Derby "Build Night" pack meetings
  (Feb 4/11/18) — the spec's explicit exception to once-a-month.
- **Blue & Gold Banquet / AOL crossover**: Wed Mar 11 (a pack meeting, placed as
  early as practical in March while leaving AOLs time to finish rank work).
- **Committee Meetings**: last Thursday, 9pm, Zoom. Nov and Dec are moved one
  week earlier (Nov 20, Dec 18) because the true last Thursday falls in a holiday
  week with no Wednesday meeting.

## Special events
| Event | Date used | Basis |
|---|---|---|
| Council Popcorn Kickoff | Sat Aug 2, 2025 | nfcscouting.org/content/130287 (verified) |
| Pack 127 Kickoff Party | Sat Aug 9, 2025 | Spec: party end of summer before school (Aug 11). **Confirm.** |
| Fall Campout | Fri–Sun Oct 24–26, 2025 | Spec rule: weekend closest to Halloween (Fri Oct 31) that avoids trick-or-treat night. |
| Winter Campout | Fri–Sun Jan 23–25, 2026 | Spec: pick a weekend, mark TBD. **Placeholder.** |
| Pinewood Derby (Pack) | Sat–Sun Feb 21–22, 2026 | Spec rule: Sat/Sun after Valentine's Day (Sat Feb 14), ≥ 2 weeks before the district derby (Mar 7). |
| Five Rivers District Derby | Sat Mar 7, 2026 | nfcscouting.org/fiverivers1derby (verified). |
| BALOO Training (adults) | Sat–Sun Mar 28–29, 2026 | **ESTIMATE** from the Mar 24 registration close (nfcscouting.org/baloomarion). Confirm. |
| BUC TUOCS Spring Campout | Fri–Sun Apr 24–26, 2026 | Spec: spring camping at Camp Shands. **Placeholder — TBD.** |
| Frontier Day | Sat Sep 19, 2026 | nfcscouting.org/frontier lists this (next program year). **Verify** whether a spring-2026 date exists. |

Troop Visits (AOL only) are intentionally **not** on the calendar page — per spec
they are arranged by the AOL Den Leader and noted only in the program section.

## Cub Scout adventures
Source: current Cub Scout program (scouting.org/programs/cub-scouts/adventures and
individual `/cub-scout-adventures/` pages). Each rank = 6 required (incl. Bobcat,
done pack-wide in Aug–Sep) + ≥ 2 electives.

- Verified adventure placements: Bear **Whittling** (spring) + **Baloo the Builder**;
  Webelos **Chef's Knife** (spring); AOL **Knife Safety** + **Race Time** (via the
  Pinewood Derby). Rank↔grade: Lion K, Tiger 1, Wolf 2, Bear 3, Webelos 4, AOL 5.
- **To verify:** a few required-adventure names vary across secondary sources
  (Wolf/Bear "Safe and Smart" vs "Safety in Numbers"; Bear "Standing Tall").
  Confirm against the official current rank pages before publishing.

The per-den, per-meeting agendas are a suggested plan, sequenced so outdoor/fitness
adventures fall in the daylight windows and tool/cooking adventures use the Cabin /
Fellowship Hall. Adjust freely — they are recommendations, not fixed requirements.

## Regenerating
```
node generator/generate.js
```
Edit the data tables at the top of `generate.js` and re-run to produce an updated
`2025-2026 Pack 127 Calendar.docx` in the repo root. Requires Node.js and (on
Windows) PowerShell/.NET for zip packaging — no Word, Python, or npm packages.
```
