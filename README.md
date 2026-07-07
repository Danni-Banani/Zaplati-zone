# Zaplati-zone

Client-side payroll preparation tool. Takes the monthly productivity report
(«АНАЛИЗ ПРОИЗВОДИТЕЛНОСТ НА РАБОТНИЦИ», .xls) and the accounting report
(«ОСНОВНА ЗАПЛАТА», .xls), and produces:

1. **rezultat.xlsx** — normalized data: per-employee summary (earnings, bonus,
   accounting fields), per-day breakdown, full operation detail, cleaned
   accounting sheet.
2. **status file** (.txt or .json) — every check, warning and error from the run.

## Privacy

All parsing and calculation happens **in the browser** (SheetJS + vanilla JS).
No file is ever uploaded to a server. The host serves two static files and
nothing else, so hosting is interchangeable: GitHub Pages, Cloudflare Pages,
Netlify, or a local file.

## What processing does

- strips report pagination junk (`Стр. N`, repeated headers, print timestamps)
- converts Bulgarian text numbers («96,55%», «41,49») to real numbers
- fixes the ×1000 scaling of operation-level earnings (auto-detected by
  reconciling against the report's own day subtotals)
- recomputes % of norm from qty ÷ norm instead of trusting the report text
- computes the over-norm bonus: configurable threshold (default: strictly
  above 101% per day), rate (default 10% of day earnings), and level
  (day / operation / both — "both" exists only to reproduce the legacy
  internal tool, which double-counts)
- parses the Crystal Reports accounting export (dot-decimal text numbers),
  making the manual "swap dots for commas" step (file 4.1) obsolete
- joins the two files by employee name; unmatched or duplicate names are
  flagged in the status file
- reconciles everything against the report's own subtotals and control sums

## Files

| file | purpose |
|---|---|
| `index.html` | UI (Bulgarian), config panel, file inputs, downloads |
| `core.js` | all parsing/calculation logic; pure functions, no DOM; runs in Node too |
| `test/test.js` | verification harness (Node + `npm i xlsx`) |

## Tests

Test fixtures are real (anonymized) reports and are **not** committed.
Run locally:

```
npm i xlsx
node test/test.js /path/to/folder-with-F1..F4-files
```

The harness verifies: exact reproduction of the historical manual bonus
formulas (F2), reconciliation of all F1 subtotals, reproduction of the legacy
tool's totals (F3), and F4 control sums.

## Deploy

Any static host. For GitHub Pages: Settings → Pages → Deploy from branch →
`main` / root.
