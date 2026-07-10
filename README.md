# Native Sheets

> A canvas-native Google-Sheets-core forge built with VectoJS.

[![CI](https://github.com/vectojs/sheets-website/actions/workflows/ci.yml/badge.svg)](https://github.com/vectojs/sheets-website/actions/workflows/ci.yml)
[![MIT license](https://img.shields.io/badge/license-MIT-6366f1.svg)](./LICENSE)

Native Sheets stress-tests VectoJS with a virtualized 10,000 × 100 spreadsheet.
The visible sheet is one retained canvas entity, not a DOM table: every cell
position, hit test, selection rectangle, and scroll offset is explicit numeric
scene state.

## Included core behavior

- sparse formula document with scalar/range dependencies and cycle detection;
- formulas with arithmetic, percent, exponentiation, text concatenation, and
  `SUM`, `AVG`/`AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`, and `CONCAT`;
- virtualized grid, fixed headers, scrolling, cell/range selection, keyboard
  navigation, and double-click/F2/typing editing;
- native IME, clipboard, selection, and undo support through VectoJS `Input`;
- TSV copy/paste, range clearing, and transactional undo/redo;
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
bun test
bun run build
```

## Repository family

This repository lives in the `vectojs-native/sheets/` forge family. The future
CLI, skills, MCP adapter, and published core document library are separate
repositories/packages so they share the pure model layer rather than automate
the canvas UI.

## License

[MIT](./LICENSE) © 2026 Xuepoo
