import { describe, expect, it } from "bun:test";
import { SheetModel } from "@vectojs/sheets-core";
import { SheetController, SheetsApp } from "../src/view/SheetsApp";
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
});

describe("SheetsApp", () => {
  it("reflects the initially selected cell in the formula bar", () => {
    const model = new SheetModel();
    model.setCell(0, 0, "Month");
    const scene = {
      add: () => scene,
      markDirty: () => undefined,
      remove: () => scene,
      resize: () => undefined,
    };

    const app = new SheetsApp(scene as never, model);

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
      const app = new SheetsApp(scene as never, model);
      app.resize(400, 300);
      app.grid.emit("pointerdown", { localX: 40, localY: 28 });
      app.grid.emit("pointerdown", { localX: 40, localY: 28 });

      expect(app.grid.children).toHaveLength(1);
    } finally {
      globalThis.requestAnimationFrame = originalFrame;
    }
  });
});
