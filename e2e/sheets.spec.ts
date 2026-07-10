import { expect, test } from "@playwright/test";

test("keeps document, VMT semantics, and audit state aligned while editing", async ({
  page,
}) => {
  await page.goto("/?debug");

  const formulaBar = page.getByRole("textbox", { name: "Formula bar" });
  await expect(formulaBar).toHaveValue("Month");

  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Enter");
  const editor = page.getByRole("textbox", { name: "Cell editor" });
  await expect(editor).toHaveValue("Revenue");
  await editor.fill("=2+3");
  await editor.press("Enter");

  await expect
    .poll(() =>
      page.evaluate(() => ({
        raw: window.__app?.model.getRaw(0, 1),
        display: window.__app?.model.getDisplay(0, 1),
        audit: window.__app?.audit(),
        tree: window.__app?.scene.getA11yTree(),
      })),
    )
    .toEqual({
      raw: "=2+3",
      display: "5",
      audit: [],
      tree: expect.arrayContaining([
        expect.objectContaining({
          role: "application",
          label: "Spreadsheet grid, active cell B1",
        }),
      ]),
    });
});

test("resizes from the layout container at a narrow viewport", async ({
  page,
}) => {
  await page.setViewportSize({ width: 375, height: 667 });
  await page.goto("/");

  await expect
    .poll(() =>
      page.evaluate(() => ({
        scene: [window.__app?.scene.width, window.__app?.scene.height],
        canvas: [
          document.querySelector("canvas")?.clientWidth,
          document.querySelector("canvas")?.clientHeight,
        ],
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ scene: [375, 667], canvas: [375, 667], audit: [] });
});

test("drags a canvas range and applies spreadsheet navigation shortcuts", async ({
  page,
}) => {
  await page.goto("/");

  const grid = page.getByRole("application", { name: /Spreadsheet grid/ });
  const box = await grid.boundingBox();
  if (!box) throw new Error("Spreadsheet grid a11y surface is not measurable");

  // A1 → C3. Coordinates are derived from the fixed header and explicit cell
  // dimensions, not a DOM table or a screenshot.
  await page.mouse.move(box.x + 48, box.y + 36);
  await page.mouse.down();
  await page.mouse.move(box.x + 48 + 2 * 112, box.y + 36 + 2 * 24);
  await page.mouse.up();

  await expect
    .poll(() =>
      page.evaluate(() => ({
        selection: window.__app?.app.viewport.selectionRange(),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ selection: { r1: 0, c1: 0, r2: 2, c2: 2 }, audit: [] });

  await page.keyboard.press("Home");
  await page.keyboard.press("PageDown");
  await page.keyboard.press("Control+Home");
  await page.keyboard.press("Control+End");

  await expect
    .poll(() =>
      page.evaluate(() => ({
        active: window.__app?.app.viewport.selected,
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ active: { row: 9999, col: 99 }, audit: [] });
});
