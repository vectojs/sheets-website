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
