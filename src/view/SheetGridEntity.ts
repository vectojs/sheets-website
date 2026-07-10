import { Entity, type A11yAttributes, type IRenderer } from "@vectojs/core";
import { colName, SheetModel } from "@vectojs/sheets-core";
import { type CellPosition, SheetViewport } from "./SheetViewport";

export interface SheetGridEvents {
  onCellPointerDown?: (cell: CellPosition, extend: boolean) => void;
  onCellPointerMove?: (cell: CellPosition) => void;
  onCellPointerUp?: () => void;
  onScroll?: (deltaX: number, deltaY: number) => void;
}

/** Number of cells that the current grid frame will inspect and render. */
export function visibleCellCount(viewport: SheetViewport): number {
  const range = viewport.visibleRange();
  if (range.rowEnd < range.rowStart || range.colEnd < range.colStart) return 0;
  return (
    (range.rowEnd - range.rowStart + 1) * (range.colEnd - range.colStart + 1)
  );
}

/** Canvas rectangle for the normalized selected range. */
export function selectionPixelRect(viewport: SheetViewport): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  const range = viewport.selectionRange();
  const topLeft = viewport.cellRect({ row: range.r1, col: range.c1 });
  return {
    x: topLeft.x,
    y: topLeft.y,
    width: (range.c2 - range.c1 + 1) * viewport.colWidth,
    height: (range.r2 - range.r1 + 1) * viewport.rowHeight,
  };
}

/**
 * A single canvas entity for the sheet surface. It renders only the rows and
 * columns intersecting the viewport; the 10,000 by 100 document never becomes
 * a matching entity tree or a DOM cell grid.
 */
export class SheetGridEntity extends Entity {
  private pointerDragging = false;

  constructor(
    readonly model: SheetModel,
    readonly viewport: SheetViewport,
    private readonly events: SheetGridEvents = {},
  ) {
    super();
    this.interactive = true;
    this.on(
      "pointerdown",
      (event: { localX?: number; localY?: number; shiftKey?: boolean }) => {
        if (event.localX === undefined || event.localY === undefined) return;
        const cell = this.viewport.cellAt(event.localX, event.localY);
        if (!cell) return;
        this.pointerDragging = true;
        this.events.onCellPointerDown?.(cell, event.shiftKey ?? false);
      },
    );
    this.on("pointermove", (event: { localX?: number; localY?: number }) => {
      if (
        !this.pointerDragging ||
        event.localX === undefined ||
        event.localY === undefined
      )
        return;
      const cell = this.viewport.cellAt(event.localX, event.localY);
      if (cell) this.events.onCellPointerMove?.(cell);
    });
    this.on("pointerup", () => {
      if (!this.pointerDragging) return;
      this.pointerDragging = false;
      this.events.onCellPointerUp?.();
    });
    this.on("pointerleave", () => {
      // Pointer capture on the projected grid keeps normal drags alive after
      // leaving its bounds. This reset covers synthetic/non-captured events.
      if (!this.pointerDragging) return;
      this.pointerDragging = false;
      this.events.onCellPointerUp?.();
    });
    this.on(
      "wheel",
      (event: { deltaX?: number; deltaY?: number; preventDefault(): void }) => {
        event.preventDefault();
        this.events.onScroll?.(event.deltaX ?? 0, event.deltaY ?? 0);
      },
    );
  }

  resize(width: number, height: number): void {
    this.width = width;
    this.height = height;
    this.viewport.resize(width, height);
  }

  isPointInside(sceneX: number, sceneY: number): boolean {
    const local = this.worldToLocal(sceneX, sceneY);
    return (
      local !== null &&
      local.x >= 0 &&
      local.y >= 0 &&
      local.x < this.width &&
      local.y < this.height
    );
  }

  getA11yAttributes(): A11yAttributes {
    const { row, col } = this.viewport.selected;
    return {
      role: "application",
      label: `Spreadsheet grid, active cell ${colName(col)}${row + 1}`,
    };
  }

  render(renderer: IRenderer): void {
    drawRect(renderer, 0, 0, this.width, this.height, "#ffffff");

    const {
      rowHeaderWidth,
      columnHeaderHeight,
      colWidth,
      rowHeight,
      scrollX,
      scrollY,
    } = this.viewport;
    const range = this.viewport.visibleRange();
    const bodyWidth = Math.max(0, this.width - rowHeaderWidth);
    const bodyHeight = Math.max(0, this.height - columnHeaderHeight);

    renderer.save();
    renderer.clip(rowHeaderWidth, columnHeaderHeight, bodyWidth, bodyHeight);
    for (let row = range.rowStart; row <= range.rowEnd; row++) {
      const y = columnHeaderHeight + row * rowHeight - scrollY;
      for (let col = range.colStart; col <= range.colEnd; col++) {
        const x = rowHeaderWidth + col * colWidth - scrollX;
        const display = this.model.getDisplay(row, col);
        if (display)
          renderer.fillText(
            display,
            x + 6,
            y + 16,
            "13px Inter, sans-serif",
            "#1f2937",
          );
      }
    }
    drawGridLines(
      renderer,
      range.rowStart,
      range.rowEnd,
      range.colStart,
      range.colEnd,
      this.viewport,
    );
    const selectedRange = selectionPixelRect(this.viewport);
    drawRect(
      renderer,
      selectedRange.x,
      selectedRange.y,
      selectedRange.width,
      selectedRange.height,
      "rgba(26, 115, 232, 0.12)",
    );
    renderer.beginPath();
    renderer.roundRect(
      selectedRange.x,
      selectedRange.y,
      selectedRange.width,
      selectedRange.height,
      0,
    );
    renderer.stroke("#1a73e8", 2);
    const selected = this.viewport.cellRect(this.viewport.selected);
    renderer.beginPath();
    renderer.roundRect(
      selected.x,
      selected.y,
      selected.width,
      selected.height,
      0,
    );
    renderer.stroke("#1a73e8", 2);
    renderer.restore();

    drawRect(renderer, 0, 0, rowHeaderWidth, columnHeaderHeight, "#f8fafc");
    drawRect(
      renderer,
      rowHeaderWidth,
      0,
      bodyWidth,
      columnHeaderHeight,
      "#f8fafc",
    );
    drawRect(
      renderer,
      0,
      columnHeaderHeight,
      rowHeaderWidth,
      bodyHeight,
      "#f8fafc",
    );

    for (let col = range.colStart; col <= range.colEnd; col++) {
      const x = rowHeaderWidth + col * colWidth - scrollX;
      renderer.fillText(
        colName(col),
        x + 6,
        18,
        "600 12px Inter, sans-serif",
        "#475569",
      );
    }
    for (let row = range.rowStart; row <= range.rowEnd; row++) {
      const y = columnHeaderHeight + row * rowHeight - scrollY;
      renderer.fillText(
        String(row + 1),
        6,
        y + 16,
        "600 12px Inter, sans-serif",
        "#475569",
      );
    }

    renderer.beginPath();
    renderer.moveTo(rowHeaderWidth, 0);
    renderer.lineTo(rowHeaderWidth, this.height);
    renderer.moveTo(0, columnHeaderHeight);
    renderer.lineTo(this.width, columnHeaderHeight);
    renderer.stroke("#cbd5e1", 1);
  }
}

function drawRect(
  renderer: IRenderer,
  x: number,
  y: number,
  width: number,
  height: number,
  color: string,
): void {
  renderer.beginPath();
  renderer.roundRect(x, y, width, height, 0);
  renderer.fill(color);
}

function drawGridLines(
  renderer: IRenderer,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
  viewport: SheetViewport,
): void {
  const {
    rowHeaderWidth,
    columnHeaderHeight,
    colWidth,
    rowHeight,
    scrollX,
    scrollY,
  } = viewport;
  renderer.beginPath();
  for (let col = colStart; col <= colEnd + 1; col++) {
    const x = rowHeaderWidth + col * colWidth - scrollX;
    renderer.moveTo(x, columnHeaderHeight);
    renderer.lineTo(x, viewport.height);
  }
  for (let row = rowStart; row <= rowEnd + 1; row++) {
    const y = columnHeaderHeight + row * rowHeight - scrollY;
    renderer.moveTo(rowHeaderWidth, y);
    renderer.lineTo(viewport.width, y);
  }
  renderer.stroke("#e2e8f0", 1);
}
