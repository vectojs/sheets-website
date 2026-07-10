import { describe, expect, it } from "bun:test";
import { createBrowserXlsxFileBridge } from "../src/view/xlsxFileBridge";

type Listener = () => void;

class FakeElement {
  hidden = false;
  type = "";
  accept = "";
  download = "";
  href = "";
  files: FileList | null = null;
  clicked = false;
  removed = false;
  private readonly listeners = new Map<string, Listener>();

  addEventListener(type: string, listener: Listener): void {
    this.listeners.set(type, listener);
  }

  click(): void {
    this.clicked = true;
  }

  remove(): void {
    this.removed = true;
  }

  emit(type: string): void {
    this.listeners.get(type)?.();
  }
}

function createFakeDocument(): {
  document: Document;
  inputs: FakeElement[];
  anchors: FakeElement[];
} {
  const inputs: FakeElement[] = [];
  const anchors: FakeElement[] = [];
  const body = { append: () => undefined };
  const document = {
    body,
    createElement: (tag: string) => {
      const element = new FakeElement();
      if (tag === "input") inputs.push(element);
      if (tag === "a") anchors.push(element);
      return element;
    },
  } as unknown as Document;
  return { document, inputs, anchors };
}

describe("createBrowserXlsxFileBridge", () => {
  it("uses a transient hidden file input and returns its bytes", async () => {
    const fake = createFakeDocument();
    const bridge = createBrowserXlsxFileBridge(fake.document, {
      createObjectURL: () => "blob:unused",
      revokeObjectURL: () => undefined,
    });

    const pending = bridge.chooseXlsx();
    const input = fake.inputs[0]!;
    input.files = [
      {
        arrayBuffer: async () => Uint8Array.from([1, 2, 3]).buffer,
      },
    ] as unknown as FileList;
    input.emit("change");

    await expect(pending).resolves.toEqual(Uint8Array.from([1, 2, 3]));
    expect(input.type).toBe("file");
    expect(input.accept).toContain(".xlsx");
    expect(input.hidden).toBe(true);
    expect(input.clicked).toBe(true);
    expect(input.removed).toBe(true);
  });

  it("downloads a deterministic filename and revokes its object URL", () => {
    const fake = createFakeDocument();
    const revoked: string[] = [];
    const bridge = createBrowserXlsxFileBridge(fake.document, {
      createObjectURL: () => "blob:numera",
      revokeObjectURL: (url) => revoked.push(url),
    });

    bridge.downloadXlsx(Uint8Array.from([1, 2]), "numera-workbook.xlsx");

    const anchor = fake.anchors[0]!;
    expect(anchor.download).toBe("numera-workbook.xlsx");
    expect(anchor.href).toBe("blob:numera");
    expect(anchor.clicked).toBe(true);
    expect(anchor.removed).toBe(true);
    expect(revoked).toEqual(["blob:numera"]);
  });
});
