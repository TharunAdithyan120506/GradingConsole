# Test Report — BITS Digital Grading Console

**Result: 21/21 unit tests and 41/41 end-to-end tests pass.** Run date 30 Sep 2026, Chromium 141 (Playwright 1.56), Node 22, Python 3.11. The full suite and the screenshot review were re-run after the BITS Pilani Digital rebrand (official logo, brand palette, portal header); the single-file build test also checks that the embedded logo renders.

## How it was tested

- **Unit tests** (`tests/core.test.cjs`) exercise the DOM-free grading core directly: parsing real `.xlsx`/`.xls`/`.csv` files, validation rules, band logic, statistics, grading and CSV output.
- **End-to-end tests** (`tests/e2e.py`) drive the real page in headless Chromium: uploads through the file input, clicks, typing, keyboard-only navigation, mouse drags on the chart, downloads (the CSV is parsed and checked cell by cell), clipboard, reload persistence, reduced motion, dark theme and seven viewport sizes. Every test fails on any console error or warning, and the main flows scan the visible page for `NaN`, `undefined`, `null` or `Infinity`.
- **Fixtures** (`tests/make-fixtures.cjs`) are generated, not hand-edited: a clean three-course file with every default boundary mark (0, 19, 20, 29 … 79, 80, 100), a second file for re-upload tests, the brief's exact headers, a messy real-world sheet (title rows, extra column, text/decimal/negative/over-100/blank marks, duplicates, numeric course code, case variants, a course with no valid rows), structural failures (missing column, empty, header-only, corrupt, zero bytes, not a spreadsheet), degenerate distributions (one student, identical marks, 30 students on one mark) and a 5,000-row file.
- **Expected values** for statistics were computed independently in Python (`statistics.pstdev`, `median`), not copied from the app.
- **Visual review**: screenshots at 360, 390, 768, 1024, 1280, 1440 and 1920px in light and dark themes (saved to `tests/results/screens/`) were inspected; defects found this way are listed below.

## Coverage of the required checks

| Required check | Where it is verified |
|---|---|
| Valid upload and parsing | E2E *Valid upload…*, *Headers exactly as printed in the brief…*, *CSV, legacy .xls and data on a second sheet…*; unit *clean workbook…*, *csv, legacy xls…* |
| Invalid file handling | E2E *Invalid file types, empty, corrupt and zero-byte files…* (6 cases), *Spreadsheet library outage…*; unit *structural file errors…* |
| Repeated uploads | E2E *Re-uploading replaces everything…*, *A failed re-upload keeps the previous file…* |
| Unique course selection | E2E *Valid upload lists unique courses*, *Numeric course codes…* |
| Statistics and grade calculations | E2E *Statistics are correct for every boundary mark*; unit *statistics are correct and never NaN*, *analysis counts every student…* |
| Every grade boundary | Unit *default bands… grade every boundary*; E2E *Export… every boundary graded correctly in the CSV* (0, 19, 20, 29, 30, 39, 40, 49, 50, 59, 60, 69, 70, 79, 80, 100) |
| Invalid and corrected grade configurations | E2E *Reversed, out-of-range and blank values…*, *Single-mark bands…*; unit *band validation catches…*, *stored band settings…* |
| Dynamic chart updates | E2E *Dragging a cutoff handle updates bands, chart colours and counts*, *Chart scales crowded bins…*, *Normal curve toggles…*, *Hover readout…* |
| Reset and default restoration | E2E *Restore defaults is undoable…*, *Saved bands are restored…* |
| CSV contents and filename | E2E *Export: gated on instructor name…*, *Decimal marks…*, *Sample data… SAMPLE_ export*, *Copy CSV…*, *Issues list… downloads*; unit *grade CSV has the expected layout…*, *CSV fields are quoted…* |
| Empty states | E2E *Empty state renders…*, *A course with no valid students…* |
| Keyboard operation | E2E *Keyboard-only flow…*, *Cutoff handles work by keyboard…*, *File guide dialog…* |
| Responsive layouts | E2E *No horizontal overflow at 360–1920px* (14 renders) plus manual screenshot review |
| Reduced-motion behaviour | E2E *Reduced motion…*; all functional tests also run with reduced motion, so every workflow is proven usable without animation |
| Browser console cleanliness | Every E2E test asserts no console errors/warnings; *Full flow with motion enabled…* covers the animated path end to end |

## End-to-end results (Playwright, Chromium)

**41 of 41 passed.** Each test runs in a fresh browser context against the app served over HTTP.

| Area | Test | Result | Notes |
|---|---|---|---|
| Empty states | Empty state renders with no console errors; later steps are locked | Pass |  |
| Empty states | A course with no valid students shows an explanation and cannot export | Pass |  |
| Invalid files | Invalid file types, empty, corrupt and zero-byte files are explained | Pass | 6 failure modes |
| Upload & parsing | Valid upload lists unique courses; data checks summarised | Pass |  |
| Upload & parsing | Headers exactly as printed in the brief are recognised | Pass |  |
| Upload & parsing | CSV, legacy .xls and data on a second sheet all load | Pass |  |
| Repeated uploads | Re-uploading replaces everything: no duplicate or stale courses, selection cleared | Pass |  |
| Repeated uploads | A failed re-upload keeps the previous file loaded and says so | Pass |  |
| Data validation | Messy sheet: header below title rows, every problem row listed, nothing silently dropped | Pass |  |
| Data validation | Numeric course codes are selectable and graded (original bug) | Pass |  |
| Data validation | Decimal marks: rounding policy is explicit, applied and recorded in the export | Pass | 79.6 → A (nearest) / A− (kept); 80.2 → 81 (up) |
| Data validation | Duplicate IDs: identical rows counted once, conflicting rows excluded | Pass |  |
| Data validation | Issues list scopes to the course and downloads as CSV | Pass |  |
| Export gating | Excluded rows must be acknowledged before export | Pass |  |
| Statistics & grading | Statistics are correct for every boundary mark | Pass | {'count': '16', 'mean': '49.6', 'median': '49.5', 'sd': '25.7', 'max': '100', 'min': '0'} |
| Export | Export: gated on instructor name, then every boundary graded correctly in the CSV | Pass | CS-F111_grades_2026-09-30.csv; marks 0,19,20,29,…,79,80,100 all correct |
| Export | Copy CSV puts the same sheet on the clipboard | Pass |  |
| Export | Repeated exports report correct ordinals and elapsed time | Pass |  |
| Export | Timer starts when a course opens, pauses at export and resumes on change | Pass |  |
| Grade bands | Linked edges: editing a minimum moves the neighbouring maximum; counts update live | Pass |  |
| Grade bands | Reversed, out-of-range and blank values are caught in context and block export | Pass |  |
| Grade bands | Single-mark bands are allowed (original rejected Min = Max) | Pass |  |
| Grade bands | Restore defaults is undoable; undo/redo buttons and Ctrl+Z work | Pass |  |
| Grade bands | Cutoff handles work by keyboard (arrows, Shift, Home/End) with slider semantics | Pass |  |
| Grade bands | Dragging a cutoff handle updates bands, chart colours and counts | Pass |  |
| Persistence | Saved bands are restored for the same course after reload, with a way back to defaults | Pass |  |
| Analytics | Chart scales crowded bins, marks mean/median, table view mirrors the data | Pass |  |
| Analytics | Normal curve toggles on/off and is disabled when marks are all equal | Pass |  |
| Analytics | Hover readout names the mark, count, grade and students | Pass |  |
| Roster | Roster search, grade filter, borderline filter, sorting and pagination | Pass | 5,000-row file parsed and rendered in 0.47s |
| Roster | Borderline "show" link filters the roster to that band | Pass |  |
| Onboarding | Sample data is clearly labelled and flows through to a SAMPLE_ export | Pass |  |
| Onboarding | Remove file returns to the empty state and can be undone | Pass |  |
| Resilience | Spreadsheet library outage is reported instead of failing silently | Pass |  |
| Accessibility | File guide dialog opens, traps focus and closes with Escape | Pass |  |
| Accessibility | Keyboard-only flow: course, cutoff, name and export without a mouse | Pass | 4 tab stops checked for visible focus |
| Theming | Dark theme toggle applies, persists and keeps contrast tokens | Pass |  |
| Motion | Reduced motion: animations collapse to ~1ms and numbers render final values at once | Pass |  |
| Console | Full flow with motion enabled produces no console errors or warnings | Pass |  |
| Deployment | Single-file build (dist/grading-console.html) runs the full workflow | Pass |  |
| Responsive | No horizontal overflow at 360–1920px; screenshots captured in light and dark | Pass | phone-sm, phone, tablet, laptop-sm, laptop, desktop, wide |

## Unit results (Node test runner, `assets/core.js`)

**21 of 21 passed.**

- file checks reject wrong types, empty and oversized files
- clean workbook: unique courses, counts and no issues
- header names from the brief are recognised
- messy workbook: header below title rows, every bad row reported, nothing dropped silently
- header matching is word-aware
- structural file errors are explained
- csv, legacy xls, second sheet and large files parse
- numeric IDs and course codes become text; text numbers are accepted
- default bands are valid and grade every boundary correctly
- editing an edge moves the shared neighbour edge
- band validation catches reversed, gaps, overlaps, missing ends and bad values
- single-mark bands are allowed
- stored band settings are accepted only when valid
- rounding policies grade decimals without gaps
- analysis counts every student and flags borderline marks
- statistics are correct and never NaN
- CSV fields are quoted and guarded against formulas
- grade CSV has the expected layout and refuses invalid input
- issues CSV lists every problem in plain words
- file names, ordinals and durations
- sample workbook runs through the real parser and shows its planted problems

## Defects found during testing and fixed

| Found by | Problem | Fix |
|---|---|---|
| Unit test | “Remarks” matched as a marks column (substring “mark”). | Header matching made word-aware; test added. |
| E2E overflow check | Page scrolled sideways on phones: the CSV preview sized to its longest line because the tablet grid's `align-items: start` carried into the phone layout. | Stretch alignment at phone width; `min-width: 0` on side-column children. |
| E2E overflow check | Rail tooltips (hover-only) positioned to the right of the top-bar buttons pushed the page wider at ≤820px. | Tooltips drop below the buttons in the top-bar layout. |
| E2E theme check | Under reduced motion, a global 1ms transition still animated every property change. | Reduced motion now disables transitions outright. |
| Keyboard test | Choosing a course with Enter dropped focus to the page because the course list re-renders. | Focus is restored to the same course after re-render. |
| Screenshot review | Round cutoff handles hid the grade letters on phones. | Slim vertical grips; letters show only where they fit (grades remain identified by band cards, tooltip and table view). |
| Screenshot review | Warning icon on course cards rendered at full size; card names misaligned when a course had nothing to plot. | Icon sized; empty sparkline baseline keeps cards aligned. |
| Screenshot review | Headline tracking too tight in the display face; “students” axis label clipped at phone width. | Tracking loosened; label anchored inside the chart. |

## Palette validation

Brand tokens were checked for contrast: white on the primary violet `#5E49E2` is 5.95:1, body ink on white 16.4:1, muted text 4.8:1 or better, dark-theme ink on the light violet 7.1:1. Grade-band colours were checked with a colour-vision-deficiency validator (OKLab ΔE on adjacent bands, lightness band, chroma floor) for both themes. Light and dark palettes pass all separation checks; bar-to-background contrast below 3:1 on the lighter steps is covered by direct grade labels, the tooltip and the table view, as the validator requires.

## Limitations of this test run

- Only Chromium was available offline in the test environment (Firefox and WebKit downloads were blocked), so other engines were not automated. The code avoids engine-specific APIs; `color-mix()`, `<dialog>` and CSS transforms on SVG are supported in current Firefox and Safari.
- Web fonts were served from local copies of the same families during screenshots; the deployed page loads them from Google Fonts with system-font fallbacks.
- Screen-reader output was not recorded with a real screen reader; semantics were checked through roles, labels, live regions and keyboard tests.

## Re-running

```bash
npm install                                  # test-only dependency (SheetJS for fixtures and unit tests)
pip install playwright && python3 -m playwright install chromium
npm test                                     # build, fixtures, unit tests, end-to-end tests
```

Results are written to `tests/results/e2e-results.json`; screenshots to `tests/results/screens/`.
