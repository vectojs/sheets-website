<!-- Conventional Commit title, for example: feat(grid): add fill-handle interaction -->

## What and why

<!-- Describe the user interaction and link the issue: Closes #123. -->

## VectoJS scene and state changes

<!-- Name the entities, numeric layout/state, event routing, a11y projection, and core-model operations involved. -->

## Verification

- [ ] `bun run format:check`
- [ ] `bun run lint`
- [ ] `bun test`
- [ ] `bun run build`
- [ ] Chromium interaction tests pass locally.
- [ ] CI passes Chromium, Firefox, WebKit, and `auditScene(scene) === []`.

## Checklist

- [ ] The page remains canvas-native; no sibling DOM layout or CSS UI was introduced.
- [ ] Responsive behavior was checked at narrow/desktop viewports and browser zoom.
- [ ] Pointer, keyboard, IME, clipboard, and focus behavior are deterministic where affected.
- [ ] Exact-pinned `@vectojs/*` packages are used.
- [ ] `README.md`, `CHANGELOG.md`, and shared forge docs were updated when status changed.
- [ ] Documentation and non-obvious code comments are English.
