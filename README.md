# BITS Digital Grading Console

A browser-based grading tool for BITS Pilani Digital. Upload a marks sheet, tune grade bands against the real distribution, and export a clean grade sheet. **Runs entirely in the browser — no data leaves your machine.**

## Project Structure

```
GradingConsole/
├── index.html          ← Entry point (replaces monolithic grading-console.html)
├── src/
│   ├── css/
│   │   └── styles.css  ← All application styles (design tokens, layout, components)
│   └── js/
│       ├── core.js     ← Pure grading logic (GradingCore) — DOM-free, testable
│       └── app.js      ← UI layer (depends on core.js and SheetJS CDN)
├── package.json        ← Dev server setup
├── BUGFIX_LOG.md
└── TEST_REPORT (1).md
```

## Getting Started

### Run locally

```bash
npm install
npm run dev
```

Then open [http://localhost:3000](http://localhost:3000) in your browser.

### Without npm

Just open `index.html` directly in a browser — it works as a pure static file too, except that the `.xlsx` upload requires a local server due to browser security restrictions on CORS/CDN scripts.

## Features

- Upload `.xlsx`, `.xls`, or `.csv` marks sheets
- Auto-detect student ID, name, and marks columns
- Interactive grade band editor with live histogram
- Grade presets: Absolute %, Relative (σ), Percentile
- Data quality issue detection
- Export to CSV or Excel
- Dark mode support
- Fully keyboard-accessible

## Architecture

| File | Purpose |
|------|---------|
| `src/js/core.js` | Pure functions: file parsing, validation, statistics, grading, CSV export. No DOM access — can be `require()`'d in Node.js tests. |
| `src/js/app.js` | UI wiring: file upload, step navigation, histogram rendering, band editor, grade sheet table. Depends on `window.GradingCore`. |
| `src/css/styles.css` | Design token system (CSS custom properties), layout, all component styles, responsive breakpoints, dark mode. |
