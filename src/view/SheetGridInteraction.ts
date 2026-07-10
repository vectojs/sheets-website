import type { Rect } from "@vectojs/sheets-core";
import { type CellPosition, SheetViewport } from "./SheetViewport";

export interface SheetGridEvents {
  onCellPointerDown?: (cell: CellPosition, extend: boolean) => void;
  onCellPointerMove?: (cell: CellPosition) => void;
  onCellPointerUp?: () => void;
  onScroll?: (deltaX: number, deltaY: number) => void;
  onAxisResize?: (axis: "row" | "column", index: number, size: number) => void;
  onFill?: (target: Rect) => void;
  onGestureChange?: () => void;
}

export interface AxisResizeTarget {
  axis: "row" | "column";
  index: number;
  size: number;
}

export interface ResizePreview extends AxisResizeTarget {
  nextSize: number;
}

const FILL_HANDLE_SIZE = 8;
const MIN_ROW_SIZE = 12;
const MIN_COLUMN_SIZE = 32;

/**
 * Deep pointer-state module for the grid. The Entity forwards local numeric
 * events; this module owns gesture precedence, preview state, and commit timing.
 */
export class SheetGridInteraction {
  private selectionDragging = false;
  private resizeDrag:
    (AxisResizeTarget & { startCoordinate: number; nextSize: number }) | null =
    null;
  private fillDrag: { target: Rect } | null = null;

  constructor(
    private readonly viewport: SheetViewport,
    private readonly events: SheetGridEvents,
  ) {}

  get resizePreview(): ResizePreview | null {
    if (!this.resizeDrag) return null;
    const { axis, index, size, nextSize } = this.resizeDrag;
    return { axis, index, size, nextSize };
  }

  get fillPreview(): Rect | null {
    return this.fillDrag?.target ?? null;
  }

  pointerDown(localX: number, localY: number, extend = false): void {
    const resize = headerResizeTargetAt(this.viewport, localX, localY);
    if (resize) {
      this.resizeDrag = {
        ...resize,
        startCoordinate: resize.axis === "row" ? localY : localX,
        nextSize: resize.size,
      };
      this.events.onGestureChange?.();
      return;
    }
    if (pointInRect(localX, localY, fillHandleRect(this.viewport))) {
      this.fillDrag = { target: this.viewport.selectionRange() };
      this.events.onGestureChange?.();
      return;
    }
    const cell = this.viewport.cellAt(localX, localY);
    if (!cell) return;
    this.selectionDragging = true;
    this.events.onCellPointerDown?.(cell, extend);
  }

  pointerMove(localX: number, localY: number): void {
    if (this.resizeDrag) {
      const coordinate = this.resizeDrag.axis === "row" ? localY : localX;
      const minimum =
        this.resizeDrag.axis === "row" ? MIN_ROW_SIZE : MIN_COLUMN_SIZE;
      this.resizeDrag.nextSize = Math.max(
        minimum,
        Math.round(
          this.resizeDrag.size + coordinate - this.resizeDrag.startCoordinate,
        ),
      );
      this.events.onGestureChange?.();
      return;
    }
    if (this.fillDrag) {
      const cell = this.viewport.cellAt(localX, localY);
      if (cell) this.fillDrag.target = fillTargetRange(this.viewport, cell);
      this.events.onGestureChange?.();
      return;
    }
    if (!this.selectionDragging) return;
    const cell = this.viewport.cellAt(localX, localY);
    if (cell) this.events.onCellPointerMove?.(cell);
  }

  pointerUp(): void {
    if (this.resizeDrag) {
      const drag = this.resizeDrag;
      this.resizeDrag = null;
      this.events.onAxisResize?.(drag.axis, drag.index, drag.nextSize);
      this.events.onGestureChange?.();
      return;
    }
    if (this.fillDrag) {
      const drag = this.fillDrag;
      this.fillDrag = null;
      this.events.onFill?.(drag.target);
      this.events.onGestureChange?.();
      return;
    }
    if (!this.selectionDragging) return;
    this.selectionDragging = false;
    this.events.onCellPointerUp?.();
  }

  pointerLeave(): void {
    if (!this.selectionDragging) return;
    this.selectionDragging = false;
    this.events.onCellPointerUp?.();
  }
}

export function rangePixelRect(
  viewport: SheetViewport,
  range: Rect,
): { x: number; y: number; width: number; height: number } {
  const topLeft = viewport.cellRect({ row: range.r1, col: range.c1 });
  const bottomRight = viewport.cellRect({ row: range.r2, col: range.c2 });
  return {
    x: topLeft.x,
    y: topLeft.y,
    width: bottomRight.x + bottomRight.width - topLeft.x,
    height: bottomRight.y + bottomRight.height - topLeft.y,
  };
}

export function selectionPixelRect(viewport: SheetViewport): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  return rangePixelRect(viewport, viewport.selectionRange());
}

export function fillHandleRect(viewport: SheetViewport): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  const selection = selectionPixelRect(viewport);
  return {
    x: selection.x + selection.width - FILL_HANDLE_SIZE / 2,
    y: selection.y + selection.height - FILL_HANDLE_SIZE / 2,
    width: FILL_HANDLE_SIZE,
    height: FILL_HANDLE_SIZE,
  };
}

/** Resolve header-edge hit zones from the same variable geometry used to draw them. */
export function headerResizeTargetAt(
  viewport: SheetViewport,
  localX: number,
  localY: number,
  tolerance = 5,
): AxisResizeTarget | null {
  if (
    localX >= 0 &&
    localX < viewport.rowHeaderWidth &&
    localY >= viewport.columnHeaderHeight
  ) {
    const cell = viewport.cellAt(viewport.rowHeaderWidth, localY);
    if (!cell) return null;
    const rect = viewport.cellRect(cell);
    if (Math.abs(localY - rect.y) <= tolerance && cell.row > 0)
      return {
        axis: "row",
        index: cell.row - 1,
        size: viewport.rowSizeAt(cell.row - 1),
      };
    if (Math.abs(localY - (rect.y + rect.height)) <= tolerance)
      return { axis: "row", index: cell.row, size: rect.height };
  }
  if (
    localY >= 0 &&
    localY < viewport.columnHeaderHeight &&
    localX >= viewport.rowHeaderWidth
  ) {
    const cell = viewport.cellAt(localX, viewport.columnHeaderHeight);
    if (!cell) return null;
    const rect = viewport.cellRect(cell);
    if (Math.abs(localX - rect.x) <= tolerance && cell.col > 0)
      return {
        axis: "column",
        index: cell.col - 1,
        size: viewport.columnSizeAt(cell.col - 1),
      };
    if (Math.abs(localX - (rect.x + rect.width)) <= tolerance)
      return { axis: "column", index: cell.col, size: rect.width };
  }
  return null;
}

function fillTargetRange(viewport: SheetViewport, cell: CellPosition): Rect {
  const source = viewport.selectionRange();
  const rowDistance =
    cell.row < source.r1
      ? source.r1 - cell.row
      : Math.max(0, cell.row - source.r2);
  const columnDistance =
    cell.col < source.c1
      ? source.c1 - cell.col
      : Math.max(0, cell.col - source.c2);
  if (rowDistance >= columnDistance)
    return {
      r1: Math.min(source.r1, cell.row),
      c1: source.c1,
      r2: Math.max(source.r2, cell.row),
      c2: source.c2,
    };
  return {
    r1: source.r1,
    c1: Math.min(source.c1, cell.col),
    r2: source.r2,
    c2: Math.max(source.c2, cell.col),
  };
}

function pointInRect(
  x: number,
  y: number,
  rect: { x: number; y: number; width: number; height: number },
): boolean {
  return (
    x >= rect.x &&
    x <= rect.x + rect.width &&
    y >= rect.y &&
    y <= rect.y + rect.height
  );
}
