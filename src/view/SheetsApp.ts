import { Scene } from "@vectojs/core";
import { Input } from "@vectojs/ui";
import {
  type CellFormat,
  copyRange,
  fromCsv,
  pasteText,
  SheetHistory,
  SheetModel,
  toCsv,
  toWorkbookJson,
  Workbook,
} from "@vectojs/sheets-core";
import { SheetGridEntity } from "./SheetGridEntity";
import { SheetTabsEntity } from "./SheetTabsEntity";
import { SheetToolbarEntity } from "./SheetToolbarEntity";
import { type CellPosition, SheetViewport } from "./SheetViewport";

const TOOLBAR_HEIGHT = 48;
const TABS_HEIGHT = 32;

/** Pure interaction state shared by canvas input and the native editor. */
export class SheetController {
  editing: CellPosition | null = null;
  draft = "";
  readonly history: SheetHistory;

  constructor(
    readonly model: SheetModel,
    readonly viewport: SheetViewport,
  ) {
    this.history = new SheetHistory(model);
  }

  select(cell: CellPosition): void {
    this.viewport.select(cell);
    this.viewport.ensureVisible(this.viewport.selected);
  }

  extendSelection(cell: CellPosition): void {
    this.viewport.extendSelection(cell);
    this.viewport.ensureVisible(this.viewport.selected);
  }

  moveSelection(rowDelta: number, colDelta: number, extend = false): void {
    this.viewport.moveSelection(rowDelta, colDelta, extend);
    this.viewport.ensureVisible(this.viewport.selected);
  }

  pageSelection(direction: -1 | 1, extend = false): void {
    this.moveSelection(direction * this.viewport.pageRows(), 0, extend);
  }

  selectAll(): void {
    const used = this.model.getUsedRange();
    if (!used) return;
    this.viewport.select({ row: used.r1, col: used.c1 });
    this.viewport.extendSelection({ row: used.r2, col: used.c2 });
    this.viewport.ensureVisible(this.viewport.selected);
  }

  scroll(deltaX: number, deltaY: number): void {
    this.viewport.scrollBy(deltaX, deltaY);
  }

  beginEdit(initialValue?: string): string {
    this.editing = { ...this.viewport.selected };
    this.draft =
      initialValue ?? this.model.getRaw(this.editing.row, this.editing.col);
    return this.draft;
  }

  setDraft(value: string): void {
    this.draft = value;
  }

  commitEdit(value = this.draft): void {
    if (!this.editing) return;
    this.history.apply([
      { row: this.editing.row, col: this.editing.col, raw: value },
    ]);
    this.editing = null;
    this.draft = "";
  }

  cancelEdit(): void {
    this.editing = null;
    this.draft = "";
  }

  writeSelected(raw: string): void {
    const { row, col } = this.viewport.selected;
    this.history.apply([{ row, col, raw }]);
  }

  undo(): void {
    this.history.undo();
  }

  redo(): void {
    this.history.redo();
  }

  copySelection(): string {
    return copyRange(this.model, this.viewport.selectionRange());
  }

  paste(text: string): void {
    this.history.apply(pasteText(text, this.viewport.selected, this.model));
  }

  clearSelection(): void {
    const range = this.viewport.selectionRange();
    const writes = this.model
      .getCellsInRange(range)
      .map(({ row, col }) => ({ row, col, raw: "" }));
    this.history.apply(writes);
  }

  applyFormat(format: CellFormat): void {
    const range = this.viewport.selectionRange();
    const writes = [];
    for (let row = range.r1; row <= range.r2; row++) {
      for (let col = range.c1; col <= range.c2; col++)
        writes.push({ row, col, format });
    }
    this.history.applyFormats(writes);
  }
}

/** Canvas shell, formula bar and short-lived native cell editor. */
export class SheetsApp {
  viewport: SheetViewport;
  controller: SheetController;
  grid: SheetGridEntity;
  readonly tabs: SheetTabsEntity;
  readonly toolbar: SheetToolbarEntity;
  readonly formulaBar: Input;

  private editor: Input | null = null;
  private tabEditor: Input | null = null;
  private lastPointer: { cell: CellPosition; at: number } | null = null;
  private readonly keyboardListener: (event: KeyboardEvent) => void;
  private readonly copyListener: (event: ClipboardEvent) => void;
  private readonly cutListener: (event: ClipboardEvent) => void;
  private readonly pasteListener: (event: ClipboardEvent) => void;

  constructor(
    readonly scene: Scene,
    readonly workbook: Workbook,
  ) {
    const model = workbook.activeSheet.model;
    this.viewport = new SheetViewport({
      rows: model.rows,
      cols: model.cols,
      rowHeight: 24,
      colWidth: 112,
    });
    this.controller = new SheetController(model, this.viewport);
    this.grid = this.createGrid(model);
    this.tabs = new SheetTabsEntity(workbook, {
      onSelect: (id) => this.selectSheet(id),
      onAdd: () => this.addSheet(),
      onRename: (id) => this.beginSheetRename(id),
      onDelete: (id) => this.deleteSheet(id),
    });
    this.toolbar = new SheetToolbarEntity((format) => this.copyExport(format));
    this.formulaBar = new Input({
      width: 320,
      height: 32,
      placeholder: "Formula bar",
      bg: "#ffffff",
      border: "#cbd5e1",
      color: "#0f172a",
      onChange: (value) => this.controller.setDraft(value),
    });
    this.formulaBar.on(
      "keydown",
      (event: { key?: string; preventDefault(): void }) => {
        if (event.key !== "Enter") return;
        event.preventDefault();
        this.commitFormulaBar();
      },
    );
    this.formulaBar.on("blur", () => this.commitFormulaBar());

    this.scene.add(this.grid);
    this.scene.add(this.toolbar);
    this.scene.add(this.formulaBar);
    this.scene.add(this.tabs);
    this.syncFormulaBar();
    this.keyboardListener = (event) => this.handleKeyboard(event);
    this.copyListener = (event) => this.handleCopy(event);
    this.cutListener = (event) => this.handleCut(event);
    this.pasteListener = (event) => this.handlePaste(event);
    if (typeof window !== "undefined") {
      window.addEventListener("keydown", this.keyboardListener);
      window.addEventListener("copy", this.copyListener);
      window.addEventListener("cut", this.cutListener);
      window.addEventListener("paste", this.pasteListener);
    }
  }

  resize(width: number, height: number): void {
    this.scene.resize(width, height);
    this.grid.setPosition(0, TOOLBAR_HEIGHT);
    this.grid.resize(width, Math.max(0, height - TOOLBAR_HEIGHT - TABS_HEIGHT));
    this.toolbar.setPosition(0, 0);
    this.toolbar.resize(112, TOOLBAR_HEIGHT);
    this.formulaBar.setPosition(120, 8);
    this.formulaBar.width = Math.max(120, width - 132);
    this.tabs.setPosition(0, Math.max(TOOLBAR_HEIGHT, height - TABS_HEIGHT));
    this.tabs.resize(width, TABS_HEIGHT);
    this.scene.markDirty();
  }

  get model(): SheetModel {
    return this.workbook.activeSheet.model;
  }

  selectSheet(id: string): void {
    if (id === this.workbook.activeSheetId) return;
    this.removeEditor(false);
    this.workbook.setActiveSheet(id);
    this.scene.remove(this.grid);
    this.viewport = new SheetViewport({
      rows: this.model.rows,
      cols: this.model.cols,
      rowHeight: 24,
      colWidth: 112,
    });
    this.controller = new SheetController(this.model, this.viewport);
    this.grid = this.createGrid(this.model);
    this.scene.add(this.grid);
    this.resize(this.scene.width, this.scene.height);
    this.syncFormulaBar();
    this.scene.markDirty();
  }

  addSheet(): void {
    const sheet = this.workbook.addSheet();
    this.selectSheet(sheet.id);
  }

  deleteSheet(id: string): void {
    this.removeTabEditor(false);
    const wasActive = id === this.workbook.activeSheetId;
    this.workbook.deleteSheet(id);
    if (wasActive) this.rebuildActiveSheet();
    this.scene.markDirty();
  }

  destroy(): void {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", this.keyboardListener);
      window.removeEventListener("copy", this.copyListener);
      window.removeEventListener("cut", this.cutListener);
      window.removeEventListener("paste", this.pasteListener);
    }
    this.removeEditor(false);
    this.removeTabEditor(false);
  }

  private handleCellPointerDown(cell: CellPosition, extend: boolean): void {
    const now = performance.now();
    const isDoublePointer =
      this.lastPointer !== null &&
      this.lastPointer.cell.row === cell.row &&
      this.lastPointer.cell.col === cell.col &&
      now - this.lastPointer.at < 350;
    this.lastPointer = { cell, at: now };
    if (extend) this.controller.extendSelection(cell);
    else this.controller.select(cell);
    if (isDoublePointer) this.beginEdit();
    else this.syncFormulaBar();
    this.scene.markDirty();
  }

  private createGrid(model: SheetModel): SheetGridEntity {
    return new SheetGridEntity(model, this.viewport, {
      onCellPointerDown: (cell, extend) =>
        this.handleCellPointerDown(cell, extend),
      onCellPointerMove: (cell) => this.handleCellPointerMove(cell),
      onScroll: (x, y) => {
        this.controller.scroll(x, y);
        this.scene.markDirty();
      },
    });
  }

  private beginSheetRename(id: string): void {
    if (this.tabEditor) return;
    const index = this.workbook.sheets.findIndex((sheet) => sheet.id === id);
    if (index < 0) return;
    const sheet = this.workbook.getSheet(id);
    const editor = new Input({
      width: 108,
      height: 26,
      value: sheet.name,
      placeholder: "Sheet name",
      bg: "#ffffff",
      border: "#1a73e8",
      color: "#0f172a",
      radius: 4,
      padding: 6,
    });
    editor.setPosition(index * 120 + 4, 3);
    editor.on("keydown", (event: { key?: string; preventDefault(): void }) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.removeTabEditor(true);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.removeTabEditor(false);
      }
    });
    editor.on("blur", () => this.removeTabEditor(true));
    this.tabEditor = editor;
    this.tabs.add(editor);
    requestAnimationFrame(() => this.scene.getA11yElement(editor.id)?.focus());
  }

  private removeTabEditor(commit: boolean): void {
    if (!this.tabEditor) return;
    const editor = this.tabEditor;
    this.tabEditor = null;
    if (commit) {
      const index = Math.floor(editor.x / 120);
      const sheet = this.workbook.sheets[index];
      if (sheet) {
        try {
          this.workbook.renameSheet(sheet.id, editor.value);
        } catch {
          // Invalid rename leaves the current name intact; the next edit can retry.
        }
      }
    }
    this.tabs.remove(editor);
    this.scene.markDirty();
  }

  private rebuildActiveSheet(): void {
    this.scene.remove(this.grid);
    this.viewport = new SheetViewport({
      rows: this.model.rows,
      cols: this.model.cols,
      rowHeight: 24,
      colWidth: 112,
    });
    this.controller = new SheetController(this.model, this.viewport);
    this.grid = this.createGrid(this.model);
    this.scene.add(this.grid);
    this.resize(this.scene.width, this.scene.height);
    this.syncFormulaBar();
  }

  private handleCellPointerMove(cell: CellPosition): void {
    if (this.editor) return;
    this.controller.extendSelection(cell);
    this.scene.markDirty();
  }

  private handleKeyboard(event: KeyboardEvent): void {
    if (isNativeTextTarget(event.target)) return;
    const modifier = event.ctrlKey || event.metaKey;
    if (modifier && event.key.toLowerCase() === "a") {
      event.preventDefault();
      this.controller.selectAll();
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (modifier && event.key.toLowerCase() === "z") {
      event.preventDefault();
      if (event.shiftKey) this.controller.redo();
      else this.controller.undo();
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (modifier && event.key.toLowerCase() === "y") {
      event.preventDefault();
      this.controller.redo();
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (
      modifier &&
      (event.key.toLowerCase() === "b" || event.key.toLowerCase() === "i")
    ) {
      event.preventDefault();
      const current = this.model.getFormat(
        this.viewport.selected.row,
        this.viewport.selected.col,
      );
      this.controller.applyFormat(
        event.key.toLowerCase() === "b"
          ? { bold: !current.bold }
          : { italic: !current.italic },
      );
      this.scene.markDirty();
      return;
    }
    if (event.key === "Home" || event.key === "End") {
      event.preventDefault();
      const target = modifier
        ? event.key === "Home"
          ? { row: 0, col: 0 }
          : { row: this.model.rows - 1, col: this.model.cols - 1 }
        : {
            row: this.viewport.selected.row,
            col: event.key === "Home" ? 0 : this.model.cols - 1,
          };
      if (event.shiftKey) this.controller.extendSelection(target);
      else this.controller.select(target);
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (event.key === "PageUp" || event.key === "PageDown") {
      event.preventDefault();
      this.controller.pageSelection(
        event.key === "PageUp" ? -1 : 1,
        event.shiftKey,
      );
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      this.controller.clearSelection();
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    const movement: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Tab: [0, event.shiftKey ? -1 : 1],
      Enter: [1, 0],
    };
    if (event.key === "F2") {
      event.preventDefault();
      this.beginEdit();
      return;
    }
    if (event.key === "Enter" && !this.controller.editing) {
      event.preventDefault();
      this.beginEdit();
      return;
    }
    const delta = movement[event.key];
    if (delta) {
      event.preventDefault();
      this.controller.moveSelection(delta[0], delta[1], event.shiftKey);
      this.syncFormulaBar();
      this.scene.markDirty();
      return;
    }
    if (
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey
    ) {
      event.preventDefault();
      this.beginEdit(event.key);
    }
  }

  private handleCopy(event: ClipboardEvent): void {
    if (isNativeTextTarget(document.activeElement)) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", this.controller.copySelection());
  }

  private handleCut(event: ClipboardEvent): void {
    if (isNativeTextTarget(document.activeElement)) return;
    event.preventDefault();
    event.clipboardData?.setData("text/plain", this.controller.copySelection());
    this.controller.clearSelection();
    this.syncFormulaBar();
    this.scene.markDirty();
  }

  private handlePaste(event: ClipboardEvent): void {
    if (isNativeTextTarget(document.activeElement)) return;
    const text = event.clipboardData?.getData("text/plain");
    if (text === undefined) return;
    event.preventDefault();
    if (!text.includes("\t") && text.includes(","))
      this.controller.history.apply(
        fromCsv(text, this.viewport.selected, this.model),
      );
    else this.controller.paste(text);
    this.syncFormulaBar();
    this.scene.markDirty();
  }

  private beginEdit(initialValue?: string): void {
    if (this.editor) return;
    const value = this.controller.beginEdit(initialValue);
    const rect = this.viewport.cellRect(this.viewport.selected);
    const editor = new Input({
      width: rect.width,
      height: rect.height,
      value,
      placeholder: "Cell editor",
      bg: "#ffffff",
      border: "#1a73e8",
      color: "#0f172a",
      radius: 0,
      padding: 6,
      onChange: (next) => this.controller.setDraft(next),
    });
    editor.setPosition(rect.x, rect.y);
    editor.on("keydown", (event: { key?: string; preventDefault(): void }) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.removeEditor(true);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.removeEditor(false);
      }
    });
    editor.on("blur", () => this.removeEditor(true));
    this.editor = editor;
    this.grid.add(editor);
    this.scene.markDirty();
    requestAnimationFrame(() => this.scene.getA11yElement(editor.id)?.focus());
  }

  private removeEditor(commit: boolean): void {
    if (!this.editor) return;
    const editor = this.editor;
    this.editor = null;
    if (commit) this.controller.commitEdit(editor.value);
    else this.controller.cancelEdit();
    this.grid.remove(editor);
    this.syncFormulaBar();
    this.scene.markDirty();
  }

  private commitFormulaBar(): void {
    if (this.editor) return;
    this.controller.select(this.viewport.selected);
    this.controller.writeSelected(this.formulaBar.value);
    this.syncFormulaBar();
    this.scene.markDirty();
  }

  private syncFormulaBar(): void {
    const { row, col } = this.viewport.selected;
    this.formulaBar.value = this.model.getRaw(row, col);
  }

  private copyExport(format: "json" | "csv"): void {
    const content =
      format === "json"
        ? toWorkbookJson(this.workbook)
        : toCsv(this.model, this.viewport.selectionRange());
    void navigator.clipboard?.writeText(content);
  }
}

function isNativeTextTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement
  );
}
