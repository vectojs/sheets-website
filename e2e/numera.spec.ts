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

test("yields projected text copy and traces its content route", async ({
  page,
}) => {
  await page.goto("/?debug");
  const status = "Native selectable status";
  await page.evaluate((text) => {
    const debug = window.__app;
    if (!debug) throw new Error("Numera debug surface is unavailable");
    debug.app.toolbar.setStatus(text);
  }, status);
  await expect
    .poll(() =>
      page.evaluate(
        (text) =>
          Array.from(
            document.querySelectorAll<HTMLElement>("[data-vecto-content]"),
          ).some((element) => element.textContent === text),
        status,
      ),
    )
    .toBe(true);

  const result = await page.evaluate(async (text) => {
    const debug = window.__app;
    if (!debug) throw new Error("Numera debug surface is unavailable");

    const content = Array.from(
      document.querySelectorAll<HTMLElement>("[data-vecto-content]"),
    ).find((element) => element.textContent === text);
    if (!content) throw new Error("Projected toolbar status is unavailable");

    const range = document.createRange();
    range.selectNodeContents(content);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);

    const clipboard = new DataTransfer();
    const copyEvent = new Event("copy", { bubbles: true, cancelable: true });
    Object.defineProperty(copyEvent, "clipboardData", { value: clipboard });
    content.dispatchEvent(copyEvent);

    const rect = content.getBoundingClientRect();
    content.dispatchEvent(
      new PointerEvent("pointerdown", {
        bubbles: true,
        clientX: rect.left + 2,
        clientY: rect.top + 2,
        pointerId: 1,
      }),
    );
    await Promise.resolve();

    return {
      selection: selection?.toString(),
      copyPrevented: copyEvent.defaultPrevented,
      clipboardText: clipboard.getData("text/plain"),
      trace: debug
        .debugTrace?.()
        .findLast((entry) => entry.type === "pointerdown"),
    };
  }, status);

  expect(result).toMatchObject({
    selection: "Native selectable status",
    copyPrevented: false,
    clipboardText: "",
    trace: {
      type: "pointerdown",
      source: "content",
      defaultPrevented: false,
    },
  });
  expect(result.trace?.targetId).toBeTruthy();
  expect(result.trace?.targetPath).toContain("Scene > SheetToolbarEntity");
  expect(result.trace?.targetPath).toContain("Text#");
});

test("keeps every command reachable across responsive toolbar breakpoints", async ({
  page,
}) => {
  const commands = [
    "Export workbook as JSON",
    "Export selection as CSV",
    "Insert rows",
    "Delete rows",
    "Insert columns",
    "Delete columns",
    "Sort selection ascending",
    "Sort selection descending",
    "Import XLSX workbook",
    "Export XLSX workbook",
  ];

  for (const width of [375, 600, 759, 760]) {
    await page.setViewportSize({ width, height: 667 });
    await page.goto("/");
    for (const name of commands) {
      const button = page.getByRole("button", { name });
      await expect(button).toBeVisible();
      const box = await button.boundingBox();
      if (!box) throw new Error(`${name} is not measurable at ${width}px`);
      expect(box.height).toBeGreaterThanOrEqual(44);
      expect(box.x).toBeGreaterThanOrEqual(0);
      expect(box.x + box.width).toBeLessThanOrEqual(width);
    }

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
      .toEqual({ scene: [width, 667], canvas: [width, 667], audit: [] });
  }
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
  await page.getByRole("button", { name: "Insert rows" }).click();
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

test("sorts selected rows through Canvas and keyboard intentions", async ({
  page,
}) => {
  await page.goto("/?debug");
  await page.evaluate(() => {
    const app = window.__app?.app;
    const model = window.__app?.model;
    if (!app || !model) throw new Error("Numera debug surface is unavailable");
    model.setCell(1, 0, "Beta");
    model.setCell(1, 1, "20");
    model.setCell(1, 2, "=B2*2");
    model.setCell(2, 0, "Alpha");
    model.setCell(2, 1, "10");
    model.setCell(2, 2, "=B3*2");
    app.controller.select({ row: 1, col: 2 });
    app.controller.extendSelection({ row: 2, col: 0 });
  });
  await page.keyboard.press("Control+Alt+Shift+ArrowUp");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        first: window.__app?.model.getRaw(1, 0),
        second: window.__app?.model.getRaw(2, 0),
      })),
    )
    .toEqual({ first: "Beta", second: "Alpha" });

  await page.getByRole("button", { name: "Sort selection ascending" }).click();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        first: window.__app?.model.getRaw(1, 0),
        firstFormula: window.__app?.model.getRaw(1, 2),
        second: window.__app?.model.getRaw(2, 0),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({
      first: "Alpha",
      firstFormula: "=B2*2",
      second: "Beta",
      audit: [],
    });
  await expect(
    page.getByRole("status", { name: /Sorted 2 rows ascending by A/ }),
  ).toBeVisible();

  await page.keyboard.press("Control+z");
  await expect(page.getByRole("status")).toHaveCount(0);
  await page.keyboard.press("Alt+Shift+ArrowDown");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        first: window.__app?.model.getRaw(1, 0),
        second: window.__app?.model.getRaw(2, 0),
      })),
    )
    .toEqual({ first: "Beta", second: "Alpha" });
});

test("commits projected editor drafts before a toolbar sort", async ({
  page,
}) => {
  await page.goto("/?debug");
  await page.evaluate(() => {
    const app = window.__app?.app;
    const model = window.__app?.model;
    if (!app || !model) throw new Error("Numera debug surface is unavailable");
    model.setCell(1, 0, "Beta");
    model.setCell(2, 0, "Alpha");
    app.controller.select({ row: 2, col: 0 });
    app.controller.extendSelection({ row: 1, col: 0 });
  });

  await page.getByRole("textbox", { name: "Formula bar" }).fill("Zulu");
  await page.getByRole("button", { name: "Sort selection ascending" }).click();
  await expect
    .poll(() =>
      page.evaluate(() => ({
        first: window.__app?.model.getRaw(1, 0),
        second: window.__app?.model.getRaw(2, 0),
      })),
    )
    .toEqual({ first: "Alpha", second: "Zulu" });

  await page.evaluate(() => {
    const app = window.__app?.app;
    const model = window.__app?.model;
    if (!app || !model) throw new Error("Numera debug surface is unavailable");
    model.setCell(1, 0, "Beta");
    model.setCell(2, 0, "Alpha");
    app.controller.select({ row: 1, col: 0 });
    app.controller.extendSelection({ row: 2, col: 0 });
    (
      app as unknown as {
        beginEdit(): void;
      }
    ).beginEdit();
  });
  await page.getByRole("textbox", { name: "Cell editor" }).fill("Zulu");
  await page.getByRole("button", { name: "Sort selection ascending" }).click();
  await expect(page.getByRole("textbox", { name: "Cell editor" })).toHaveCount(
    0,
  );
  await expect
    .poll(() =>
      page.evaluate(() => ({
        first: window.__app?.model.getRaw(1, 0),
        second: window.__app?.model.getRaw(2, 0),
      })),
    )
    .toEqual({ first: "Beta", second: "Zulu" });
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

test("routes internal clipboard ranges through formula-aware Core transfer", async ({
  page,
}) => {
  await page.goto("/?debug");
  await page.evaluate(() => {
    const app = window.__app?.app;
    const model = window.__app?.model;
    if (!app || !model) throw new Error("Numera debug surface is unavailable");
    model.setCell(0, 0, "5");
    model.setCell(0, 1, "=A1*2");
    model.setFormat(0, 1, { bold: true, numberFormat: "currency" });
    app.controller.select({ row: 0, col: 0 });
    app.controller.extendSelection({ row: 0, col: 1 });

    const clipboard = new DataTransfer();
    const copyEvent = new Event("copy", { cancelable: true });
    Object.defineProperty(copyEvent, "clipboardData", { value: clipboard });
    window.dispatchEvent(copyEvent);
    app.controller.select({ row: 2, col: 2 });
    const pasteEvent = new Event("paste", { cancelable: true });
    Object.defineProperty(pasteEvent, "clipboardData", { value: clipboard });
    window.dispatchEvent(pasteEvent);
  });

  await expect
    .poll(() =>
      page.evaluate(() => ({
        raw: window.__app?.model.getRaw(2, 3),
        display: window.__app?.model.getDisplay(2, 3),
        format: window.__app?.model.getFormat(2, 3),
        selection: window.__app?.app.viewport.selectionRange(),
        audit: window.__app?.audit(),
      })),
    )
    .toEqual({
      raw: "=C3*2",
      display: "$10",
      format: { bold: true, numberFormat: "currency" },
      selection: { r1: 2, c1: 2, r2: 2, c2: 3 },
      audit: [],
    });

  await page.keyboard.press("Control+z");
  await expect
    .poll(() =>
      page.evaluate(() => ({
        left: window.__app?.model.getRaw(2, 2),
        right: window.__app?.model.getRaw(2, 3),
      })),
    )
    .toEqual({ left: "=B3*1.1", right: "" });
});

test("does not reuse stale internal formats for matching external text", async ({
  page,
}) => {
  await page.goto("/?debug");
  await page.evaluate(() => {
    const app = window.__app?.app;
    const model = window.__app?.model;
    if (!app || !model) throw new Error("Numera debug surface is unavailable");
    model.setCell(0, 0, "same");
    model.setFormat(0, 0, { bold: true });
    app.controller.select({ row: 0, col: 0 });
    const internalClipboard = new DataTransfer();
    const copyEvent = new Event("copy", { cancelable: true });
    Object.defineProperty(copyEvent, "clipboardData", {
      value: internalClipboard,
    });
    window.dispatchEvent(copyEvent);

    const externalClipboard = new DataTransfer();
    externalClipboard.setData("text/plain", "same");
    app.controller.select({ row: 3, col: 3 });
    const pasteEvent = new Event("paste", { cancelable: true });
    Object.defineProperty(pasteEvent, "clipboardData", {
      value: externalClipboard,
    });
    window.dispatchEvent(pasteEvent);
  });

  await expect
    .poll(() =>
      page.evaluate(() => ({
        raw: window.__app?.model.getRaw(3, 3),
        format: window.__app?.model.getFormat(3, 3),
      })),
    )
    .toEqual({ raw: "same", format: {} });
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

  const chooserPromise = page.waitForEvent("filechooser");
  await page.getByRole("button", { name: "Import XLSX workbook" }).click();
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
  await page.getByRole("button", { name: "Export XLSX workbook" }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("numera-workbook.xlsx");
});

test("reports corrupt XLSX imports through the Canvas toolbar state", async ({
  page,
}) => {
  await page.goto("/");

  const chooserPromise = page.waitForEvent("filechooser");
  await page.getByRole("button", { name: "Import XLSX workbook" }).click();
  const chooser = await chooserPromise;
  await chooser.setFiles({
    name: "corrupt.xlsx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from([1, 2, 3]),
  });

  await expect(
    page.getByRole("status", { name: /Import failed \(INVALID_ARCHIVE\)/ }),
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
