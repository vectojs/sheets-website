import { Entity, type A11yAttributes, type IRenderer } from "@vectojs/core";
import { colName, SheetModel } from "@vectojs/sheets-core";
import { measureText } from "@vectojs/ui";
import {
  fillHandleRect,
  rangePixelRect,
  selectionPixelRect,
  SheetGridInteraction,
  type SheetGridEvents,
} from "./SheetGridInteraction";
import { SheetViewport } from "./SheetViewport";

export {
  fillHandleRect,
  headerResizeTargetAt,
  selectionPixelRect,
} from "./SheetGridInteraction";
export type { SheetGridEvents } from "./SheetGridInteraction";

/** Number of cells that the current grid frame will inspect and render. */
export function visibleCellCount(viewport: SheetViewport): number {
  const range = viewport.visibleRange();
  if (range.rowEnd < range.rowStart || range.colEnd < range.colStart) return 0;
  return (
    (range.rowEnd - range.rowStart + 1) * (range.colEnd - range.colStart + 1)
  );
}

/**
 * A single canvas entity for the sheet surface. It renders only the rows and
 * columns intersecting the viewport; the 10,000 by 100 document never becomes
 * a matching entity tree or a DOM cell grid.
 */
export class SheetGridEntity extends Entity {
  private readonly interaction: SheetGridInteraction;

  constructor(
    readonly model: SheetModel,
    readonly viewport: SheetViewport,
    private readonly events: SheetGridEvents = {},
  ) {
    super();
    this.interaction = new SheetGridInteraction(viewport, events);
    this.interactive = true;
    this.on(
      "pointerdown",
      (event: { localX?: number; localY?: number; shiftKey?: boolean }) => {
        if (event.localX === undefined || event.localY === undefined) return;
        this.interaction.pointerDown(
          event.localX,
          event.localY,
          event.shiftKey ?? false,
        );
      },
    );
    this.on("pointermove", (event: { localX?: number; localY?: number }) => {
      if (event.localX === undefined || event.localY === undefined) return;
      this.interaction.pointerMove(event.localX, event.localY);
    });
    this.on("pointerup", () => this.interaction.pointerUp());
    this.on("pointerleave", () => {
      // Pointer capture on the projected grid keeps normal drags alive after
      // leaving its bounds. This reset covers synthetic/non-captured events.
      this.interaction.pointerLeave();
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

    const { rowHeaderWidth, columnHeaderHeight } = this.viewport;
    const range = this.viewport.visibleRange();
    const bodyWidth = Math.max(0, this.width - rowHeaderWidth);
    const bodyHeight = Math.max(0, this.height - columnHeaderHeight);

    renderer.save();
    renderer.clip(rowHeaderWidth, columnHeaderHeight, bodyWidth, bodyHeight);
    for (let row = range.rowStart; row <= range.rowEnd; row++) {
      for (let col = range.colStart; col <= range.colEnd; col++) {
        const { x, y, width, height } = this.viewport.cellRect({ row, col });
        const format = this.model.getFormat(row, col);
        if (format.background)
          drawRect(renderer, x, y, width, height, format.background);
        const display = this.model.getDisplay(row, col);
        if (display) {
          const font = `${format.italic ? "italic " : ""}${format.bold ? "700 " : ""}13px Inter, sans-serif`;
          const textWidth = measureText(display, font);
          const textX =
            format.horizontalAlign === "right"
              ? x + width - textWidth - 6
              : format.horizontalAlign === "center"
                ? x + (width - textWidth) / 2
                : x + 6;
          renderer.fillText(
            display,
            textX,
            y + Math.min(16, height - 6),
            font,
            format.foreground ?? "#1f2937",
          );
        }
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
    if (this.interaction.fillPreview) {
      const target = rangePixelRect(
        this.viewport,
        this.interaction.fillPreview,
      );
      drawRect(
        renderer,
        target.x,
        target.y,
        target.width,
        target.height,
        "rgba(26, 115, 232, 0.08)",
      );
      renderer.beginPath();
      renderer.roundRect(target.x, target.y, target.width, target.height, 0);
      renderer.stroke("#1a73e8", 1);
    }
    const handle = fillHandleRect(this.viewport);
    drawRect(
      renderer,
      handle.x,
      handle.y,
      handle.width,
      handle.height,
      "#1a73e8",
    );
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

    const resizePreview = this.interaction.resizePreview;
    if (resizePreview) {
      const cell =
        resizePreview.axis === "row"
          ? { row: resizePreview.index, col: range.colStart }
          : { row: range.rowStart, col: resizePreview.index };
      const rect = this.viewport.cellRect(cell);
      renderer.beginPath();
      if (resizePreview.axis === "row") {
        const y = rect.y + resizePreview.nextSize;
        renderer.moveTo(0, y);
        renderer.lineTo(this.width, y);
      } else {
        const x = rect.x + resizePreview.nextSize;
        renderer.moveTo(x, 0);
        renderer.lineTo(x, this.height);
      }
      renderer.stroke("#1a73e8", 2);
    }

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
      const x = this.viewport.cellRect({ row: range.rowStart, col }).x;
      renderer.fillText(
        colName(col),
        x + 6,
        18,
        "600 12px Inter, sans-serif",
        "#475569",
      );
    }
    for (let row = range.rowStart; row <= range.rowEnd; row++) {
      const y = this.viewport.cellRect({ row, col: range.colStart }).y;
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
  const { rowHeaderWidth, columnHeaderHeight } = viewport;
  renderer.beginPath();
  for (let col = colStart; col <= colEnd; col++) {
    const x = viewport.cellRect({ row: rowStart, col }).x;
    renderer.moveTo(x, columnHeaderHeight);
    renderer.lineTo(x, viewport.height);
  }
  const lastColumn = viewport.cellRect({ row: rowStart, col: colEnd });
  renderer.moveTo(lastColumn.x + lastColumn.width, columnHeaderHeight);
  renderer.lineTo(lastColumn.x + lastColumn.width, viewport.height);
  for (let row = rowStart; row <= rowEnd; row++) {
    const y = viewport.cellRect({ row, col: colStart }).y;
    renderer.moveTo(rowHeaderWidth, y);
    renderer.lineTo(viewport.width, y);
  }
  const lastRow = viewport.cellRect({ row: rowEnd, col: colStart });
  renderer.moveTo(rowHeaderWidth, lastRow.y + lastRow.height);
  renderer.lineTo(viewport.width, lastRow.y + lastRow.height);
  renderer.stroke("#e2e8f0", 1);
}
