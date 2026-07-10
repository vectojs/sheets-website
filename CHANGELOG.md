# Changelog

All notable changes to this project are documented in this file.

## Unreleased

### Added

- Canvas-native virtualized 10,000 × 100 grid with frozen headers and responsive scene sizing.
- Formula model with references, ranges, aggregates, conditional logic, percent, exponentiation, and `&` concatenation.
- Cell and range selection, IME-safe editing, formula bar, TSV copy/paste, clearing, and undo/redo history.
- `?debug` VectoJS devtools hook, exposed state audit, and deterministic model/view tests.
- Pointer-captured canvas drag selection; Home/End, Ctrl/Command+Home/End and
  PageUp/PageDown navigation; used-range selection and cut support.
- Workbook tabs, sheet creation, and local versioned snapshot restoration.
- Canvas-rendered cell formatting with selection-wide Ctrl/Command+B/I actions
  and format-aware undo/redo.
- Exact-pinned core upgrade with versioned workbook JSON and CSV exchange.
