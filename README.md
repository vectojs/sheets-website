# Native Sheets

> A canvas-native Google-Sheets-core forge built with VectoJS.

[![CI](https://github.com/vectojs/sheets-website/actions/workflows/ci.yml/badge.svg)](https://github.com/vectojs/sheets-website/actions/workflows/ci.yml)
[![MIT license](https://img.shields.io/badge/license-MIT-6366f1.svg)](./LICENSE)

Native Sheets stress-tests VectoJS with a virtualized 10,000 × 100 spreadsheet.
The visible sheet is one retained canvas entity, not a DOM table: every cell
position, hit test, selection rectangle, and scroll offset is explicit numeric
scene state. Document semantics come from the exact-pinned published
[`@vectojs/sheets-core`](https://www.npmjs.com/package/@vectojs/sheets-core)
package; this repository owns only the VectoJS UI/UX adapter.

## Included core behavior

- sparse formula document with scalar/range dependencies and cycle detection;
- formulas with arithmetic, percent, exponentiation, text concatenation, and
  `SUM`, `AVG`/`AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`, and `CONCAT`;
- common numeric, logical, and recovery formulas: `ABS`, `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `AND`, `OR`, `NOT`, and `IFERROR`;
- virtualized grid, fixed headers, scrolling, pointer-drag range selection,
  keyboard navigation (arrows, Shift, Home/End, Page, Ctrl/Command corners),
  and double-click/F2/typing editing;
- native IME, clipboard, selection, and undo support through VectoJS `Input`;
- TSV copy/paste, range clearing, and transactional undo/redo;
- Ctrl/Command+A selection of the sparse used range plus Ctrl/Command+X/C/V
  and Z/Y document shortcuts, without creating writes for blank cells;
- ordered multi-sheet workbooks with a canvas-native tab strip, stable sheet
  identities, creation, rename/delete interactions, and versioned local
  snapshot persistence with safe recovery;
- selection formatting for background/foreground, bold/italic, horizontal
  alignment, and number/currency/percent display, with Ctrl/Command+B/I and
  transactional undo/redo.
- versioned workbook JSON plus formula-preserving RFC 4180 CSV serialization
  primitives, with canvas JSON/CSV copy controls and CSV paste import, shared
  by future CLI and MCP adapters.
- structural row and column insertion/deletion with sparse-format preservation,
  A1-style formula-reference rewriting, and transactional undo/redo. The
  Canvas toolbar keeps these document commands on narrow screens while
  compacting lower-priority export controls.
- sparse per-row and per-column logical sizes, Canvas header-edge resize
  gestures, and a selection fill handle. Gesture previews remain numeric VMT
  state; releasing the pointer creates one undoable document transaction.

- responsive container measurement and `?debug` VMT inspection/audit.

This forge intentionally excludes collaboration, cloud persistence, comments,
charts, Apps Script, and every Google Sheets function. Those are separate
future projects, as are `sheets-cli`, `sheets-skills`, and `sheets-mcp`.

## Development

```bash
bun install
just dev
```

Open [http://localhost:2323/?debug](http://localhost:2323/?debug) to attach the
VectoJS devtools panel. The app also exposes `window.__app` with its `scene`,
`model`, `app`, and `audit()` function. In debug mode, `debugTrace()` returns
the bounded devtools pointer/wheel/keyboard routing records. Prefer inspecting
this state and `auditScene(scene)` before using screenshots.

## Verify

```bash
just verify
just browser-verify
just browser-verify firefox
just browser-verify-all
```

Browser verification uses the dedicated port `24323` by default so it does not
attach to or terminate an Agent's development server on `2323`. The test server
is never reused, and the `just` recipes run Playwright in deterministic CI mode
with web-server tracing so its child server remains attached in non-interactive
Agent shells. Set `PLAYWRIGHT_PORT` when a parallel worktree needs another
isolated E2E port.

## Deploy

The canonical application is <https://sheets-website.vectojs.org/>. Its
Cloudflare Pages project is `sheets-website`; the fallback deployment URL is
<https://sheets-website.pages.dev>. With an authenticated Wrangler session:

```bash
just deploy
```

The deployment script streams Wrangler output, waits for the final
`Deployment complete!` marker, and then exits instead of waiting indefinitely
for Wrangler logs. An upload-only message is not treated as a completed Pages
deployment. `just deploy` runs the local non-browser verification gates and
build before invoking that script. `just browser-verify` defaults to Chromium
and accepts a Playwright project such as `firefox`; `just browser-verify-all`
runs the complete Chromium, Firefox, and WebKit matrix used by CI.

## Repository family

This repository lives in the `vectojs-native/sheets/` forge family. The future
CLI, skills, MCP adapter, and published core document library are separate
repositories/packages so they share the pure model layer rather than automate
the canvas UI.

## License

[MIT](./LICENSE) © 2026 Xuepoo
