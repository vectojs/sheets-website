import { describe, expect, it } from "bun:test";
import { SheetModel, Workbook } from "@vectojs/numera-core";
import { SheetController, NumeraApp } from "../src/view/NumeraApp";
import { SheetViewport } from "../src/view/SheetViewport";

function createController(): {
  model: SheetModel;
  controller: SheetController;
} {
  const model = new SheetModel(20, 10);
  const viewport = new SheetViewport({
    rows: model.rows,
    cols: model.cols,
    rowHeight: 24,
    colWidth: 100,
  });
  viewport.resize(340, 148);
  return { model, controller: new SheetController(model, viewport) };
}

function workbookWith(model: SheetModel): Workbook {
  const workbook = new Workbook({ rows: model.rows, cols: model.cols });
  for (const cell of model.getCellsInRange({
    r1: 0,
    c1: 0,
    r2: model.rows - 1,
    c2: model.cols - 1,
  }))
    workbook.activeSheet.model.setCell(cell.row, cell.col, cell.raw);
  return workbook;
}

describe("SheetController", () => {
  it("commits an edit to the active cell and recalculates its formula", () => {
    const { model, controller } = createController();
    controller.select({ row: 0, col: 0 });

    controller.beginEdit();
    controller.commitEdit("=1+2");

    expect(controller.editing).toBeNull();
    expect(model.getRaw(0, 0)).toBe("=1+2");
    expect(model.getDisplay(0, 0)).toBe("3");

    controller.undo();
    expect(model.getRaw(0, 0)).toBe("");
    controller.redo();
    expect(model.getDisplay(0, 0)).toBe("3");
  });

  it("cancels an edit without replacing the stored raw value", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "original");
    controller.beginEdit();
    controller.cancelEdit();

    expect(controller.editing).toBeNull();
    expect(model.getRaw(0, 0)).toBe("original");
  });

  it("moves selection and scrolls it into view", () => {
    const { controller } = createController();

    controller.moveSelection(12, 8);

    expect(controller.viewport.selected).toEqual({ row: 12, col: 8 });
    expect(controller.viewport.visibleRange()).toEqual({
      rowStart: 8,
      rowEnd: 12,
      colStart: 6,
      colEnd: 8,
    });
  });

  it("copies a selection and pastes it as one undoable transaction", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "1");
    model.setCell(0, 1, "=A1+1");
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 0, col: 1 });

    expect(controller.copySelection()).toBe("1\t=A1+1");
    controller.select({ row: 2, col: 0 });
    controller.paste("3\t=A3+1");
    expect(model.getDisplay(2, 1)).toBe("4");

    controller.undo();
    expect(model.getRaw(2, 0)).toBe("");
    expect(model.getRaw(2, 1)).toBe("");
  });

  it("clears the selected range as one undoable transaction", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "left");
    model.setCell(0, 1, "right");
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 0, col: 1 });

    controller.clearSelection();
    expect(model.getRaw(0, 0)).toBe("");
    expect(model.getRaw(0, 1)).toBe("");

    controller.undo();
    expect(model.getRaw(0, 0)).toBe("left");
    expect(model.getRaw(0, 1)).toBe("right");
  });

  it("clears only stored cells when a large range is selected", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "first");
    model.setCell(19, 9, "last");
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 19, col: 9 });

    controller.clearSelection();

    expect(model.cellCount).toBe(0);
    controller.undo();
    expect(model.getRaw(0, 0)).toBe("first");
    expect(model.getRaw(19, 9)).toBe("last");
  });

  it("selects the used range and pages through the viewport", () => {
    const { model, controller } = createController();
    model.setCell(2, 1, "left");
    model.setCell(10, 5, "right");

    controller.selectAll();
    expect(controller.viewport.selectionRange()).toEqual({
      r1: 2,
      c1: 1,
      r2: 10,
      c2: 5,
    });

    controller.select({ row: 0, col: 0 });
    controller.pageSelection(1);
    expect(controller.viewport.selected).toEqual({ row: 5, col: 0 });
  });

  it("applies a style to the selection as an undoable transaction", () => {
    const { model, controller } = createController();
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 0, col: 1 });

    controller.applyFormat({ bold: true, background: "#fef3c7" });
    expect(model.getFormat(0, 1)).toEqual({
      bold: true,
      background: "#fef3c7",
    });
    controller.undo();
    expect(model.hasFormat(0, 0)).toBe(false);
  });

  it("applies selected row structure through Core and reconciles the viewport", () => {
    const { model, controller } = createController();
    model.setCell(1, 0, "10");
    model.setCell(1, 1, "=A2*2");
    controller.select({ row: 1, col: 0 });

    expect(controller.insertRows()).toBe(true);
    expect(model.rows).toBe(21);
    expect(controller.viewport.rows).toBe(21);
    expect(model.getRaw(2, 0)).toBe("10");
    expect(model.getRaw(2, 1)).toBe("=A3*2");
    expect(controller.viewport.selected).toEqual({ row: 1, col: 0 });

    controller.undo();
    expect(model.rows).toBe(20);
    expect(controller.viewport.rows).toBe(20);
    expect(model.getRaw(1, 1)).toBe("=A2*2");

    controller.redo();
    expect(model.rows).toBe(21);
    expect(controller.viewport.rows).toBe(21);
    expect(model.getRaw(2, 1)).toBe("=A3*2");
  });

  it("does not remove the final row or column through structural commands", () => {
    const model = new SheetModel(1, 1);
    const viewport = new SheetViewport({
      rows: model.rows,
      cols: model.cols,
      rowHeight: 24,
      colWidth: 112,
      rowMetrics: model.rowMetrics,
      columnMetrics: model.columnMetrics,
    });
    const controller = new SheetController(model, viewport);

    expect(controller.deleteRows()).toBe(false);
    expect(controller.deleteColumns()).toBe(false);
    expect(model.rows).toBe(1);
    expect(model.cols).toBe(1);
  });

  it("resizes axes through Core history and refreshes viewport geometry", () => {
    const { model, controller } = createController();

    controller.resizeRow(2, 40);
    controller.resizeColumn(1, 160);
    expect(model.getAxisSize("row", 2)).toBe(40);
    expect(model.getAxisSize("column", 1)).toBe(160);
    expect(controller.viewport.cellRect({ row: 3, col: 2 })).toEqual({
      x: 312,
      y: 116,
      width: 112,
      height: 24,
    });

    controller.undo();
    expect(model.getAxisSize("column", 1)).toBe(112);
    expect(model.getAxisSize("row", 2)).toBe(40);
    controller.undo();
    expect(model.getAxisSize("row", 2)).toBe(24);
    controller.redo();
    controller.redo();
    expect(model.getAxisSize("row", 2)).toBe(40);
  });

  it("fills translated formulas and exact formats as one transaction", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "first");
    model.setCell(0, 1, '=A1&"!"');
    model.setFormat(0, 0, { bold: true });
    model.setFormat(0, 1, { background: "#fef3c7" });
    model.setCell(2, 3, "before");
    model.setFormat(2, 3, { italic: true });
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 0, col: 1 });

    controller.fillSelection({ r1: 1, c1: 0, r2: 2, c2: 3 });
    expect(model.getRaw(1, 0)).toBe("first");
    expect(model.getRaw(1, 1)).toBe('=A2&"!"');
    expect(model.getRaw(2, 2)).toBe("first");
    expect(model.getRaw(2, 3)).toBe('=C3&"!"');
    expect(model.getFormat(2, 2)).toEqual({ bold: true });
    expect(model.getFormat(2, 3)).toEqual({ background: "#fef3c7" });

    controller.undo();
    expect(model.getRaw(1, 0)).toBe("");
    expect(model.getRaw(2, 3)).toBe("before");
    expect(model.getFormat(2, 3)).toEqual({ italic: true });
  });

  it("pastes an internal range with formulas and formats as one transaction", () => {
    const { model, controller } = createController();
    model.setCell(0, 0, "5");
    model.setCell(0, 1, "=A1*2");
    model.setFormat(0, 1, { bold: true, numberFormat: "currency" });
    controller.select({ row: 0, col: 0 });
    controller.extendSelection({ row: 0, col: 1 });
    const copied = controller.copySelectionPayload();

    controller.select({ row: 2, col: 2 });
    controller.pasteRange(copied.payload);

    expect(model.getRaw(2, 2)).toBe("5");
    expect(model.getRaw(2, 3)).toBe("=C3*2");
    expect(model.getFormat(2, 3)).toEqual({
      bold: true,
      numberFormat: "currency",
    });
    controller.undo();
    expect(model.getRaw(2, 2)).toBe("");
    expect(model.getRaw(2, 3)).toBe("");
  });
});

describe("NumeraApp", () => {
  it("reflects the initially selected cell in the formula bar", () => {
    const model = new SheetModel();
    model.setCell(0, 0, "Month");
    const scene = {
      add: () => scene,
      markDirty: () => undefined,
      remove: () => scene,
      resize: () => undefined,
    };

    const app = new NumeraApp(scene as never, workbookWith(model));

    expect(app.formulaBar.value).toBe("Month");
  });

  it("mounts a cell editor as a temporary grid child rather than an overlapping scene sibling", () => {
    const model = new SheetModel();
    const scene = {
      add: () => scene,
      markDirty: () => undefined,
      remove: () => scene,
      resize: () => undefined,
      getA11yElement: () => undefined,
    };
    const originalFrame = globalThis.requestAnimationFrame;
    globalThis.requestAnimationFrame = (callback) => {
      callback(0);
      return 0;
    };
    try {
      const app = new NumeraApp(scene as never, workbookWith(model));
      app.resize(400, 300);
      app.grid.emit("pointerdown", { localX: 40, localY: 28 });
      app.grid.emit("pointerdown", { localX: 40, localY: 28 });

      expect(app.grid.children).toHaveLength(1);
    } finally {
      globalThis.requestAnimationFrame = originalFrame;
    }
  });

  it("connects Canvas resize and fill gestures to document transactions", () => {
    const model = new SheetModel(20, 10);
    model.setCell(0, 0, "seed");
    const scene = {
      add: () => scene,
      markDirty: () => undefined,
      remove: () => scene,
      resize: () => undefined,
    };
    const app = new NumeraApp(scene as never, workbookWith(model));
    app.resize(440, 300);

    app.grid.emit("pointerdown", { localX: 20, localY: 52 });
    app.grid.emit("pointermove", { localX: 20, localY: 68 });
    app.grid.emit("pointerup", {});
    expect(app.model.getAxisSize("row", 0)).toBe(40);

    const handle = app.viewport.cellRect({ row: 0, col: 0 });
    app.grid.emit("pointerdown", {
      localX: handle.x + handle.width,
      localY: handle.y + handle.height,
    });
    app.grid.emit("pointermove", { localX: 80, localY: 116 });
    app.grid.emit("pointerup", {});
    expect(app.model.getRaw(3, 0)).toBe("seed");
    expect(app.viewport.selectionRange()).toEqual({
      r1: 0,
      c1: 0,
      r2: 3,
      c2: 0,
    });
  });

  it("replaces the workbook and rebuilds the Canvas adapters after an import", () => {
    const scene = {
      width: 440,
      height: 300,
      add: () => scene,
      markDirty: () => undefined,
      remove: () => scene,
      resize: () => undefined,
    };
    const app = new NumeraApp(scene as never, new Workbook({ name: "Before" }));
    const replacement = new Workbook({ name: "Imported", rows: 20, cols: 10 });
    replacement.activeSheet.model.setCell(0, 0, "XLSX value");

    app.replaceWorkbook(replacement);

    expect(app.workbook.activeSheet.name).toBe("Imported");
    expect(app.model.getRaw(0, 0)).toBe("XLSX value");
    expect(app.formulaBar.value).toBe("XLSX value");
  });
});
