import { Entity, type A11yAttributes, type IRenderer } from "@vectojs/core";

const BUTTONS = [
  { id: "json", label: "JSON", x: 8, width: 52 },
  { id: "csv", label: "CSV", x: 64, width: 46 },
] as const;

/** Canvas export controls; actions are supplied by the application adapter. */
export class SheetToolbarEntity extends Entity {
  constructor(private readonly onExport: (format: "json" | "csv") => void) {
    super();
    this.interactive = true;
    this.on("pointerdown", (event: { localX?: number; localY?: number }) => {
      if (event.localX === undefined || event.localY === undefined) return;
      const button = BUTTONS.find(
        (candidate) =>
          event.localX! >= candidate.x &&
          event.localX! < candidate.x + candidate.width &&
          event.localY! >= 8 &&
          event.localY! < 40,
      );
      if (button) this.onExport(button.id);
    });
  }

  resize(width: number, height: number): void {
    this.width = width;
    this.height = height;
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
    return { role: "toolbar", label: "Spreadsheet export toolbar" };
  }

  render(renderer: IRenderer): void {
    renderer.beginPath();
    renderer.roundRect(0, 0, this.width, this.height, 0);
    renderer.fill("#f8fafc");
    for (const button of BUTTONS) {
      renderer.beginPath();
      renderer.roundRect(button.x, 8, button.width, 32, 5);
      renderer.fill("#ffffff");
      renderer.fillText(
        button.label,
        button.x + 10,
        28,
        "600 12px Inter, sans-serif",
        "#334155",
      );
    }
  }
}
