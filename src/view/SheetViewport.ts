export interface CellPosition {
  row: number;
  col: number;
}

export interface VisibleRange {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export interface SheetViewportOptions {
  rows: number;
  cols: number;
  rowHeight: number;
  colWidth: number;
  rowHeaderWidth?: number;
  columnHeaderHeight?: number;
}

/**
 * Pure viewport state and geometry for a sparse sheet. It deliberately owns no
 * entities: rendering code consumes the visible range while interaction code
 * uses the same coordinate conversion, keeping selection and pixels aligned.
 */
export class SheetViewport {
  readonly rows: number;
  readonly cols: number;
  readonly rowHeight: number;
  readonly colWidth: number;
  readonly rowHeaderWidth: number;
  readonly columnHeaderHeight: number;

  width = 0;
  height = 0;
  scrollX = 0;
  scrollY = 0;
  selected: CellPosition = { row: 0, col: 0 };
  private anchor: CellPosition = { row: 0, col: 0 };

  constructor(options: SheetViewportOptions) {
    this.rows = options.rows;
    this.cols = options.cols;
    this.rowHeight = options.rowHeight;
    this.colWidth = options.colWidth;
    this.rowHeaderWidth = options.rowHeaderWidth ?? 40;
    this.columnHeaderHeight = options.columnHeaderHeight ?? 28;
  }

  resize(width: number, height: number): void {
    this.width = Math.max(0, width);
    this.height = Math.max(0, height);
    this.scrollTo(this.scrollX, this.scrollY);
  }

  scrollBy(deltaX: number, deltaY: number): void {
    this.scrollTo(this.scrollX + deltaX, this.scrollY + deltaY);
  }

  scrollTo(x: number, y: number): void {
    this.scrollX = clamp(x, 0, this.maxScrollX());
    this.scrollY = clamp(y, 0, this.maxScrollY());
  }

  moveSelection(rowDelta: number, colDelta: number, extend = false): void {
    const target = {
      row: clamp(this.selected.row + rowDelta, 0, this.rows - 1),
      col: clamp(this.selected.col + colDelta, 0, this.cols - 1),
    };
    if (extend) this.extendSelection(target);
    else this.select(target);
  }

  select(position: CellPosition): void {
    const target = {
      row: clamp(position.row, 0, this.rows - 1),
      col: clamp(position.col, 0, this.cols - 1),
    };
    this.selected = target;
    this.anchor = target;
  }

  extendSelection(position: CellPosition): void {
    this.selected = {
      row: clamp(position.row, 0, this.rows - 1),
      col: clamp(position.col, 0, this.cols - 1),
    };
  }

  selectionRange(): Rect {
    return normalizeRect(this.anchor, this.selected);
  }

  /** Adjust the scroll offset just enough to place a cell inside the body. */
  ensureVisible(position: CellPosition): void {
    const cell = {
      row: clamp(position.row, 0, this.rows - 1),
      col: clamp(position.col, 0, this.cols - 1),
    };
    const left = cell.col * this.colWidth;
    const right = left + this.colWidth;
    const top = cell.row * this.rowHeight;
    const bottom = top + this.rowHeight;
    let nextX = this.scrollX;
    let nextY = this.scrollY;
    if (left < nextX) nextX = left;
    else if (right > nextX + this.bodyWidth()) nextX = right - this.bodyWidth();
    if (top < nextY) nextY = top;
    else if (bottom > nextY + this.bodyHeight())
      nextY = bottom - this.bodyHeight();
    this.scrollTo(nextX, nextY);
  }

  visibleRange(): VisibleRange {
    const bodyWidth = this.bodyWidth();
    const bodyHeight = this.bodyHeight();
    if (bodyWidth <= 0 || bodyHeight <= 0) {
      return { rowStart: 0, rowEnd: -1, colStart: 0, colEnd: -1 };
    }
    return {
      rowStart: Math.floor(this.scrollY / this.rowHeight),
      rowEnd: Math.min(
        this.rows - 1,
        Math.floor((this.scrollY + bodyHeight - 1) / this.rowHeight),
      ),
      colStart: Math.floor(this.scrollX / this.colWidth),
      colEnd: Math.min(
        this.cols - 1,
        Math.floor((this.scrollX + bodyWidth - 1) / this.colWidth),
      ),
    };
  }

  /** Number of complete rows exposed by the current canvas body. */
  pageRows(): number {
    return Math.max(1, Math.floor(this.bodyHeight() / this.rowHeight));
  }

  cellAt(localX: number, localY: number): CellPosition | null {
    if (localX < this.rowHeaderWidth || localY < this.columnHeaderHeight)
      return null;
    if (localX >= this.width || localY >= this.height) return null;
    const col = Math.floor(
      (localX - this.rowHeaderWidth + this.scrollX) / this.colWidth,
    );
    const row = Math.floor(
      (localY - this.columnHeaderHeight + this.scrollY) / this.rowHeight,
    );
    return row >= 0 && row < this.rows && col >= 0 && col < this.cols
      ? { row, col }
      : null;
  }

  cellRect(position: CellPosition): {
    x: number;
    y: number;
    width: number;
    height: number;
  } {
    return {
      x: this.rowHeaderWidth + position.col * this.colWidth - this.scrollX,
      y: this.columnHeaderHeight + position.row * this.rowHeight - this.scrollY,
      width: this.colWidth,
      height: this.rowHeight,
    };
  }

  private bodyWidth(): number {
    return Math.max(0, this.width - this.rowHeaderWidth);
  }

  private bodyHeight(): number {
    return Math.max(0, this.height - this.columnHeaderHeight);
  }

  private maxScrollX(): number {
    return Math.max(0, this.cols * this.colWidth - this.bodyWidth());
  }

  private maxScrollY(): number {
    return Math.max(0, this.rows * this.rowHeight - this.bodyHeight());
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}
import { normalizeRect, type Rect } from "@vectojs/sheets-core";
