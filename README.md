# playwright-excel-report-writer

A low-level Excel reporting engine for Playwright and test automation frameworks.

This package solves a long-standing ExcelJS limitation:
**rows containing images cannot be safely deleted or replaced**.

The engine rebuilds worksheets safely while preserving screenshots and
supports **parallel Playwright execution** via file locking.

---

## ✅ What this package does

- Safely updates Excel rows even when screenshots are present
- Preserves images for unaffected rows
- Supports sheet-per-browser or single-sheet strategies
- Parallel-worker safe (lock file)
- Does NOT impose any column structure or reporting format

---

## ❌ What this package does NOT do

- Does not decide column names
- Does not decide styling
- Does not integrate directly into Playwright hooks

👉 You must define your report structure in your own project.

---

## Module format

This package is published as **ESM-only**.

✅ Works with:
- Playwright
- TypeScript
- Node.js 18+

❌ Not supported:
- `require()` (CommonJS)

If you need CommonJS support, please open an issue.

---

## Installation

```bash
npm install playwright-excel-report-writer exceljs
```
