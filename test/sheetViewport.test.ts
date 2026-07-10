import { describe, expect, it } from "bun:test";
import { SheetViewport } from "../src/view/SheetViewport";

describe("SheetViewport", () => {
  it("clamps scrolling and reports only the visible cell rectangle", () => {
    const view = new SheetViewport({
      rows: 10_000,
      cols: 100,
      rowHeight: 24,
      colWidth: 100,
      rowHeaderWidth: 40,
      columnHeaderHeight: 28,
    });
    view.resize(440, 328);

    view.scrollBy(450, 50_000);

    expect(view.scrollX).toBe(450);
    expect(view.scrollY).toBe(50_000);
    expect(view.visibleRange()).toEqual({
      rowStart: 2083,
      rowEnd: 2095,
      colStart: 4,
      colEnd: 8,
    });
  });

  it("maps local grid coordinates to cells and rejects frozen headers", () => {
    const view = new SheetViewport({
      rows: 20,
      cols: 10,
      rowHeight: 24,
      colWidth: 100,
    });
    view.resize(340, 148);
    view.scrollBy(100, 48);

    expect(view.cellAt(39, 28)).toBeNull();
    expect(view.cellAt(40, 27)).toBeNull();
    expect(view.cellAt(40, 28)).toEqual({ row: 2, col: 1 });
    expect(view.cellAt(339, 147)).toEqual({ row: 6, col: 3 });
  });

  it("keeps selection within the sheet bounds", () => {
    const view = new SheetViewport({
      rows: 2,
      cols: 3,
      rowHeight: 24,
      colWidth: 100,
    });

    view.moveSelection(-1, -1);
    expect(view.selected).toEqual({ row: 0, col: 0 });
    view.moveSelection(20, 20);
    expect(view.selected).toEqual({ row: 1, col: 2 });
  });

  it("reconciles selection and scroll position when sheet dimensions change", () => {
    const view = new SheetViewport({
      rows: 20,
      cols: 10,
      rowHeight: 24,
      colWidth: 100,
    });
    view.resize(340, 148);
    view.select({ row: 19, col: 9 });
    view.scrollTo(900, 456);

    view.setBounds(5, 3);

    expect(view.rows).toBe(5);
    expect(view.cols).toBe(3);
    expect(view.selected).toEqual({ row: 4, col: 2 });
    expect(view.selectionRange()).toEqual({ r1: 4, c1: 2, r2: 4, c2: 2 });
    expect(view.scrollX).toBe(0);
    expect(view.scrollY).toBe(0);
  });

  it("scrolls the selected cell into the visible body", () => {
    const view = new SheetViewport({
      rows: 100,
      cols: 100,
      rowHeight: 24,
      colWidth: 100,
    });
    view.resize(340, 148);

    view.ensureVisible({ row: 20, col: 10 });

    expect(view.cellAt(40, 28)).toEqual({ row: 16, col: 8 });
    expect(view.visibleRange()).toEqual({
      rowStart: 16,
      rowEnd: 20,
      colStart: 8,
      colEnd: 10,
    });
  });

  it("normalizes a keyboard or shift-click cell range", () => {
    const view = new SheetViewport({
      rows: 20,
      cols: 10,
      rowHeight: 24,
      colWidth: 100,
    });

    view.select({ row: 4, col: 3 });
    view.extendSelection({ row: 1, col: 1 });

    expect(view.selectionRange()).toEqual({ r1: 1, c1: 1, r2: 4, c2: 3 });
    expect(view.selected).toEqual({ row: 1, col: 1 });
  });

  it("derives a page navigation stride from the visible body", () => {
    const view = new SheetViewport({
      rows: 100,
      cols: 10,
      rowHeight: 24,
      colWidth: 100,
    });
    view.resize(340, 148);

    expect(view.pageRows()).toBe(5);
  });
});
