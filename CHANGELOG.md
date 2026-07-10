# Changelog

All notable changes to this project are documented in this file.

## Unreleased

### Changed

- Renamed the product, repository, npm dependency, runtime root, and canonical
  deployment contract to Numera and `@vectojs/numera-core`.
- Adopted Microsoft Excel as the primary reference for local workbook behavior
  and personal productivity; cloud collaboration remains outside product scope.
- Exact-pinned `@vectojs/numera-core@0.3.1` for Excel-compatible boolean cell
  literals required by XLSX exchange.
- Exact-pinned `@vectojs/numera-xlsx@0.1.0` with Canvas import/export actions,
  transient native file I/O, typed corrupt-file feedback, and lazy OOXML codec
  loading so ordinary spreadsheet startup does not load ExcelJS.

### Fixed

- Isolate Playwright on a configurable E2E port so concurrent Agent development
  servers cannot be mistaken for the test server.
- Removed the projected sheet-name input when its temporary canvas entity is
  dismissed, preventing it from intercepting later tab actions in WebKit.
- Wait for Cloudflare's final deployment marker before terminating Wrangler,
  preventing upload-only CI runs from leaving the production domain stale.

### Added

- Repository `Justfile` with shared development, verification, selectable local
  browser verification, full CI browser matrix, and Cloudflare Pages deployment
  entry points.
- Exact-pinned `@vectojs/numera-core@0.2.0` with common numeric, logical, and
  error-recovery formula functions plus undoable structural operations.
- Exact-pinned `@vectojs/devtools@0.3.0` with opt-in routing trace exposed by
  `?debug` for deterministic Numera interaction diagnosis.
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
- Canvas tab rename/delete controls with temporary IME-safe editor input.
- Canvas JSON/CSV export controls and comma-delimited clipboard import.
- Canvas row/column insertion and deletion controls backed by Core's sparse
  structural history, formula-reference rewriting, and viewport-bound
  reconciliation. Narrow viewports keep those controls while compacting
  lower-priority export buttons.
- Exact-pinned `@vectojs/numera-core@0.3.0` with sparse persistent row/column
  metrics, structural remapping, and undoable axis-size history.
- Canvas header-edge resize and selection fill-handle gestures with live VMT
  previews, one-transaction commits, variable-size virtualization, and
  125%/150% scale verification.
