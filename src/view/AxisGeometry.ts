export interface AxisMetricSource {
  axisLength: number;
  default: number;
  get(index: number): number;
  entries(): readonly { index: number; size: number }[];
}

/**
 * Read-only sparse-axis geometry. Prefix deltas let the viewport calculate
 * positions from logical pixels without turning a 10,000-row sheet into a
 * dense geometry array or a matching entity tree.
 */
export class AxisGeometry {
  private entries: readonly { index: number; size: number }[] = [];
  private prefixDeltas: number[] = [];

  constructor(private readonly source: AxisMetricSource) {
    this.refresh();
  }

  get length(): number {
    return this.source.axisLength;
  }

  get totalSize(): number {
    return this.offsetOf(this.length);
  }

  refresh(): void {
    this.entries = this.source.entries();
    this.prefixDeltas = [];
    let delta = 0;
    for (const entry of this.entries) {
      delta += entry.size - this.source.default;
      this.prefixDeltas.push(delta);
    }
  }

  sizeAt(index: number): number {
    return this.source.get(clamp(index, 0, this.length - 1));
  }

  /** Pixel offset from the leading edge to the index, allowing `length`. */
  offsetOf(index: number): number {
    const bounded = clamp(index, 0, this.length);
    return bounded * this.source.default + this.deltaBefore(bounded);
  }

  /** Axis index containing a logical pixel, clamped to the axis bounds. */
  indexAt(offset: number): number {
    if (offset <= 0) return 0;
    if (offset >= this.totalSize) return this.length - 1;
    let low = 0;
    let high = this.length - 1;
    while (low <= high) {
      const middle = Math.floor((low + high) / 2);
      const start = this.offsetOf(middle);
      const end = start + this.sizeAt(middle);
      if (offset < start) high = middle - 1;
      else if (offset >= end) low = middle + 1;
      else return middle;
    }
    return clamp(low, 0, this.length - 1);
  }

  private deltaBefore(index: number): number {
    let low = 0;
    let high = this.entries.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.entries[middle].index < index) low = middle + 1;
      else high = middle;
    }
    return low === 0 ? 0 : this.prefixDeltas[low - 1];
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}
