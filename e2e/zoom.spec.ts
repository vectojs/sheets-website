import { expect, test } from "@playwright/test";

for (const scale of [1.25, 1.5] as const) {
  const viewport = {
    width: Math.floor(1280 / scale),
    height: Math.floor(800 / scale),
  };

  test.describe(`logical layout at ${scale * 100}% scale`, () => {
    test.use({ viewport, deviceScaleFactor: scale });

    test("keeps Canvas geometry, hit testing, and audit state aligned", async ({
      page,
    }) => {
      await page.goto("/?debug");
      const state = await page.evaluate(() => {
        const canvas = document.querySelector("canvas");
        return {
          scene: [window.__app?.scene.width, window.__app?.scene.height],
          client: [canvas?.clientWidth, canvas?.clientHeight],
          backing: [canvas?.width, canvas?.height],
          cell: window.__app?.app.viewport.cellAt(48, 36),
          audit: window.__app?.audit(),
        };
      });

      expect(state.scene).toEqual([viewport.width, viewport.height]);
      expect(state.client).toEqual([viewport.width, viewport.height]);
      expect(state.backing[0]).toBeGreaterThanOrEqual(
        Math.floor(viewport.width * scale),
      );
      expect(state.backing[1]).toBeGreaterThanOrEqual(
        Math.floor(viewport.height * scale),
      );
      expect(state.cell).toEqual({ row: 0, col: 0 });
      expect(state.audit).toEqual([]);
    });
  });
}
