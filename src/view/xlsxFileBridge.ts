export interface XlsxFileBridge {
  chooseXlsx(): Promise<Uint8Array | null>;
  downloadXlsx(bytes: Uint8Array, filename: string): void;
}

export interface ObjectUrlApi {
  createObjectURL(value: Blob): string;
  revokeObjectURL(url: string): void;
}

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * The only browser DOM adapter in the file flow. It creates short-lived native
 * controls for platform file selection/download and never participates in UI
 * layout; Canvas entities remain the visible application surface.
 */
export function createBrowserXlsxFileBridge(
  document: Document,
  objectUrls: ObjectUrlApi = URL,
): XlsxFileBridge {
  return {
    chooseXlsx: () => chooseXlsx(document),
    downloadXlsx: (bytes, filename) =>
      downloadXlsx(document, objectUrls, bytes, filename),
  };
}

function chooseXlsx(document: Document): Promise<Uint8Array | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = `.xlsx,${XLSX_MIME}`;
    input.hidden = true;
    input.tabIndex = -1;

    let settled = false;
    const finish = (value: Uint8Array | null): void => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(value);
    };
    input.addEventListener("cancel", () => finish(null), { once: true });
    input.addEventListener(
      "change",
      () => {
        const file = input.files?.[0];
        if (!file) {
          finish(null);
          return;
        }
        void file
          .arrayBuffer()
          .then((buffer) => finish(new Uint8Array(buffer)))
          .catch((error: unknown) => {
            if (settled) return;
            settled = true;
            input.remove();
            reject(error);
          });
      },
      { once: true },
    );
    document.body.append(input);
    input.click();
  });
}

function downloadXlsx(
  document: Document,
  objectUrls: ObjectUrlApi,
  bytes: Uint8Array,
  filename: string,
): void {
  const payload = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(payload).set(bytes);
  const href = objectUrls.createObjectURL(
    new Blob([payload], { type: XLSX_MIME }),
  );
  const anchor = document.createElement("a");
  anchor.href = href;
  anchor.download = filename;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  objectUrls.revokeObjectURL(href);
}
