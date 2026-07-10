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
- responsive container measurement and `?debug` VMT inspection/audit.

This forge intentionally excludes collaboration, cloud persistence, comments,
charts, Apps Script, and every Google Sheets function. Those are separate
future projects, as are `sheets-cli`, `sheets-skills`, and `sheets-mcp`.

## Development

```bash
bun install
bun run dev
```

Open [http://localhost:2323/?debug](http://localhost:2323/?debug) to attach the
VectoJS devtools panel. The app also exposes `window.__app` with its `scene`,
`model`, `app`, and `audit()` function. Prefer inspecting this state and
`auditScene(scene)` before using screenshots.

## Verify

```bash
bun run format:check
bun run lint
bun run test
bun run build
```

## Deploy

The Cloudflare Pages project is `sheets-website` at
<https://sheets-website.pages.dev>. With an authenticated Wrangler session:

```bash
bun run build
bun run deploy
```

The deployment script streams Wrangler output, detects both legacy and v4
success messages, and exits instead of waiting indefinitely for Wrangler logs.

## Repository family

This repository lives in the `vectojs-native/sheets/` forge family. The future
CLI, skills, MCP adapter, and published core document library are separate
repositories/packages so they share the pure model layer rather than automate
the canvas UI.

## License

[MIT](./LICENSE) © 2026 Xuepoo
