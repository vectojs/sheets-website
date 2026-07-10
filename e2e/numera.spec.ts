import { expect, test } from "@playwright/test";
import { Workbook } from "@vectojs/numera-core";
import { encodeXlsx } from "@vectojs/numera-xlsx";

async function importFixture(): Promise<Uint8Array> {
  const workbook = new Workbook({ name: "Imported", rows: 20, cols: 10 });
  workbook.activeSheet.model.setCell(0, 0, "Loaded from XLSX");
  workbook.activeSheet.model.setCell(0, 1, "=1+2");
  workbook.addSheet("Second").model.setCell(0, 0, "More data");
  return encodeXlsx(workbook);
}

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
        trace: window.__app?.debugTrace?.().map((entry) => ({
          type: entry.type,
          source: entry.source,
          key: entry.key,
        })),
        tree: window.__app?.scene.getA11yTree(),
      })),
    )
    .toEqual({
      raw: "=2+3",
      display: "5",
      audit: [],
      trace: expect.arrayContaining([
        expect.objectContaining({ type: "keydown", source: "a11y" }),
      ]),
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

  await page.keyboard.press("Control+Home");
  await page.keyboard.press("Control+b");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        format: window.__app?.model.getFormat(0, 0),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ format: { bold: true }, audit: [] });
});

test("applies undoable row structure through the canvas toolbar", async ({
  page,
}) => {
  await page.goto("/");
  const toolbar = page.getByRole("toolbar", { name: /structure and export/ });
  const box = await toolbar.boundingBox();
  if (!box) throw new Error("Spreadsheet toolbar is not measurable");

  await page.mouse.click(box.x + 132, box.y + 24);
  await expect
    .poll(() =>
      page.evaluate(() => ({
        rows: window.__app?.model.rows,
        movedHeader: window.__app?.model.getRaw(1, 0),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ rows: 10_001, movedHeader: "Month", audit: [] });

  await page.keyboard.press("Control+z");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        rows: window.__app?.model.rows,
        restoredHeader: window.__app?.model.getRaw(0, 0),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ rows: 10_000, restoredHeader: "Month", audit: [] });
});

test("resizes a row and fills cells through Canvas pointer gestures", async ({
  page,
}) => {
  await page.goto("/?debug");
  const grid = page.getByRole("application", { name: /Spreadsheet grid/ });
  const box = await grid.boundingBox();
  if (!box) throw new Error("Spreadsheet grid a11y surface is not measurable");

  await page.mouse.move(box.x + 20, box.y + 52);
  await page.mouse.down();
  await page.mouse.move(box.x + 20, box.y + 68);
  await page.mouse.up();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        rowSize: window.__app?.model.getAxisSize("row", 0),
        secondRowY: window.__app?.app.viewport.cellRect({ row: 1, col: 0 }).y,
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ rowSize: 40, secondRowY: 68, audit: [] });

  await page.mouse.move(box.x + 152, box.y + 68);
  await page.mouse.down();
  await page.mouse.move(box.x + 80, box.y + 132);
  await page.mouse.up();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        filled: window.__app?.model.getRaw(3, 0),
        selection: window.__app?.app.viewport.selectionRange(),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({
      filled: "Month",
      selection: { r1: 0, c1: 0, r2: 3, c2: 0 },
      audit: [],
    });

  await page.keyboard.press("Control+z");
  await page.keyboard.press("Control+z");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        filled: window.__app?.model.getRaw(3, 0),
        rowSize: window.__app?.model.getAxisSize("row", 0),
        trace: window.__app?.debugTrace?.().map((entry) => entry.type),
      })),
    )
    .toEqual({
      filled: "March",
      rowSize: 24,
      trace: expect.arrayContaining([
        "pointerdown",
        "pointermove",
        "pointerup",
      ]),
    });
});

test("switches and creates workbook sheets through the canvas tab strip", async ({
  page,
}) => {
  await page.goto("/");

  const tabs = page.getByRole("toolbar", { name: /Workbook sheets/ });
  const box = await tabs.boundingBox();
  if (!box)
    throw new Error("Workbook tab strip a11y surface is not measurable");

  await page.mouse.click(box.x + 120 + 16, box.y + 16);
  await expect
    .poll(() =>
      page.evaluate(() => ({
        active: window.__app?.workbook.activeSheet.name,
        raw: window.__app?.model.getRaw(0, 0),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({
      active: "Notes",
      raw: "Use the + tab to add a worksheet.",
      audit: [],
    });

  await page.mouse.click(box.x + 240 + 12, box.y + 16);
  await expect
    .poll(() =>
      page.evaluate(() => ({
        active: window.__app?.workbook.activeSheet.name,
        count: window.__app?.workbook.sheets.length,
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ active: "Sheet 3", count: 3, audit: [] });
});

test("imports and exports an XLSX workbook through Canvas toolbar intentions", async ({
  page,
}) => {
  await page.goto("/");
  const toolbar = page.getByRole("toolbar", { name: /structure and export/ });
  const box = await toolbar.boundingBox();
  if (!box) throw new Error("Spreadsheet toolbar is not measurable");

  const chooserPromise = page.waitForEvent("filechooser");
  await page.mouse.click(box.x + 322, box.y + 24);
  const chooser = await chooserPromise;
  await chooser.setFiles({
    name: "imported.xlsx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from(await importFixture()),
  });

  await expect
    .poll(() =>
      page.evaluate(() => ({
        sheets: window.__app?.workbook.sheets.map((sheet) => sheet.name),
        raw: window.__app?.model.getRaw(0, 0),
        display: window.__app?.model.getDisplay(0, 1),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({
      sheets: ["Imported", "Second"],
      raw: "Loaded from XLSX",
      display: "3",
      audit: [],
    });

  const downloadPromise = page.waitForEvent("download");
  await page.mouse.click(box.x + 385, box.y + 24);
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("numera-workbook.xlsx");
});

test("reports corrupt XLSX imports through the Canvas toolbar state", async ({
  page,
}) => {
  await page.goto("/");
  const toolbar = page.getByRole("toolbar", { name: /structure and export/ });
  const box = await toolbar.boundingBox();
  if (!box) throw new Error("Spreadsheet toolbar is not measurable");

  const chooserPromise = page.waitForEvent("filechooser");
  await page.mouse.click(box.x + 322, box.y + 24);
  const chooser = await chooserPromise;
  await chooser.setFiles({
    name: "corrupt.xlsx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from([1, 2, 3]),
  });

  await expect(
    page.getByRole("toolbar", { name: /Import failed \(INVALID_ARCHIVE\)/ }),
  ).toBeVisible();
});

test("renames and deletes a sheet through the canvas tab strip", async ({
  page,
}) => {
  await page.goto("/");
  const tabs = page.getByRole("toolbar", { name: /Workbook sheets/ });
  const box = await tabs.boundingBox();
  if (!box)
    throw new Error("Workbook tab strip a11y surface is not measurable");

  await page.mouse.dblclick(box.x + 120 + 16, box.y + 16);
  const editor = page.getByRole("textbox", { name: "Sheet name" });
  await editor.fill("Archive");
  await editor.press("Enter");
  await expect
    .poll(() => page.evaluate(() => window.__app?.workbook.sheets[1].name))
    .toBe("Archive");

  await page.mouse.click(box.x + 120 + 104, box.y + 16);
  await expect
    .poll(() =>
      page.evaluate(() => ({
        active: window.__app?.workbook.activeSheet.name,
        count: window.__app?.workbook.sheets.length,
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({ active: "Revenue", count: 1, audit: [] });
});
