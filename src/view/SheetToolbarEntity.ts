import { Entity, type A11yAttributes, type IRenderer } from "@vectojs/core";

export type SheetToolbarAction =
  | "export-json"
  | "export-csv"
  | "import-xlsx"
  | "export-xlsx"
  | "insert-row"
  | "delete-row"
  | "insert-column"
  | "delete-column";

interface ToolbarButton {
  id: SheetToolbarAction;
  label: string;
  x: number;
  width: number;
}

const WIDE_BUTTONS: readonly ToolbarButton[] = [
  { id: "export-json", label: "JSON", x: 8, width: 48 },
  { id: "export-csv", label: "CSV", x: 60, width: 44 },
  { id: "insert-row", label: "+R", x: 112, width: 40 },
  { id: "delete-row", label: "−R", x: 156, width: 40 },
  { id: "insert-column", label: "+C", x: 200, width: 40 },
  { id: "delete-column", label: "−C", x: 244, width: 40 },
  { id: "import-xlsx", label: "Import", x: 292, width: 60 },
  { id: "export-xlsx", label: "XLSX", x: 356, width: 58 },
] as const;

const COMPACT_BUTTONS: readonly ToolbarButton[] = [
  { id: "insert-row", label: "+R", x: 8, width: 40 },
  { id: "delete-row", label: "−R", x: 52, width: 40 },
  { id: "insert-column", label: "+C", x: 96, width: 40 },
  { id: "delete-column", label: "−C", x: 140, width: 40 },
] as const;

/** Canvas structural and export controls supplied by the application adapter. */
export class SheetToolbarEntity extends Entity {
  private compact = false;
  private status = "";

  constructor(private readonly onAction: (action: SheetToolbarAction) => void) {
    super();
    this.interactive = true;
    this.on("pointerdown", (event: { localX?: number; localY?: number }) => {
      if (event.localX === undefined || event.localY === undefined) return;
      const button = this.buttons().find(
        (candidate) =>
          event.localX! >= candidate.x &&
          event.localX! < candidate.x + candidate.width &&
          event.localY! >= 8 &&
          event.localY! < 40,
      );
      if (button) this.onAction(button.id);
    });
  }

  resize(width: number, height: number): void {
    this.width = width;
    this.height = height;
  }

  /** Small viewports keep structural controls and defer low-priority exports. */
  setCompact(compact: boolean): void {
    this.compact = compact;
  }

  setStatus(status: string): void {
    this.status = status;
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
    return {
      role: "toolbar",
      label: this.status
        ? `Spreadsheet structure and export toolbar. ${this.status}`
        : "Spreadsheet structure and export toolbar",
    };
  }

  render(renderer: IRenderer): void {
    renderer.beginPath();
    renderer.roundRect(0, 0, this.width, this.height, 0);
    renderer.fill("#f8fafc");
    for (const button of this.buttons()) {
      renderer.beginPath();
      renderer.roundRect(button.x, 8, button.width, 32, 5);
      renderer.fill("#ffffff");
      renderer.fillText(
        button.label,
        button.x + 8,
        28,
        "600 12px Inter, sans-serif",
        "#334155",
      );
    }
    if (this.status) {
      const maximumCharacters = Math.max(
        1,
        Math.floor((this.width - 16) / 6.5),
      );
      const text =
        this.status.length > maximumCharacters
          ? `${this.status.slice(0, Math.max(0, maximumCharacters - 1))}…`
          : this.status;
      renderer.fillText(
        text,
        8,
        56,
        "500 11px Inter, sans-serif",
        this.status.startsWith("Import failed") ? "#b91c1c" : "#475569",
      );
    }
  }

  private buttons(): readonly ToolbarButton[] {
    return this.compact ? COMPACT_BUTTONS : WIDE_BUTTONS;
  }
}
