import { describe, expect, it } from "bun:test";
import { AxisGeometry } from "../src/view/AxisGeometry";

function metrics(overrides: ReadonlyMap<number, number>) {
  return {
    axisLength: 8,
    default: 24,
    get: (index: number) => overrides.get(index) ?? 24,
    entries: () =>
      [...overrides]
        .map(([index, size]) => ({ index, size }))
        .sort((left, right) => left.index - right.index),
  };
}

describe("AxisGeometry", () => {
  it("maps sparse sizes to offsets and indices without materializing an axis", () => {
    const geometry = new AxisGeometry(metrics(new Map([[3, 40]])));

    expect(geometry.sizeAt(3)).toBe(40);
    expect(geometry.offsetOf(0)).toBe(0);
    expect(geometry.offsetOf(3)).toBe(72);
    expect(geometry.offsetOf(4)).toBe(112);
    expect(geometry.totalSize).toBe(208);
    expect(geometry.indexAt(71)).toBe(2);
    expect(geometry.indexAt(72)).toBe(3);
    expect(geometry.indexAt(111.9)).toBe(3);
    expect(geometry.indexAt(112)).toBe(4);
  });

  it("clamps lookup positions and refreshes a changed sparse source", () => {
    const overrides = new Map<number, number>();
    const geometry = new AxisGeometry(metrics(overrides));

    expect(geometry.indexAt(-1)).toBe(0);
    expect(geometry.indexAt(999)).toBe(7);
    overrides.set(0, 48);
    geometry.refresh();
    expect(geometry.offsetOf(1)).toBe(48);
  });
});
