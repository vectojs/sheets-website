import { Scene } from "@vectojs/core";
import {
  attachDevtools,
  auditScene,
  type EventTraceEntry,
} from "@vectojs/devtools";
import { createDemoModel, Workbook } from "@vectojs/numera-core";
import { NumeraApp } from "./view/NumeraApp";
import { measureSceneContainer } from "./view/sceneSizing";
import { persistWorkbook, restoreWorkbook } from "./view/workbookPersistence";

declare global {
  interface Window {
    __app?: {
      scene: Scene;
      model: ReturnType<typeof createDemoModel>;
      workbook: Workbook;
      app: NumeraApp;
      audit: () => ReturnType<typeof auditScene>;
      debugTrace?: () => readonly EventTraceEntry[];
    };
  }
}

const canvas = document.querySelector<HTMLCanvasElement>("#numera-canvas");
if (!canvas) throw new Error("Numera requires #numera-canvas");
const container = document.querySelector<HTMLElement>("#numera-root");
if (!container) throw new Error("Numera requires #numera-root");

const scene = new Scene(canvas, { disableWindowResize: true });
scene.renderMode = "onDemand";
const createDemoWorkbook = (): Workbook => {
  const demo = createDemoModel();
  const workbook = new Workbook({
    name: "Revenue",
    rows: demo.rows,
    cols: demo.cols,
  });
  for (const cell of demo.getCellsInRange({
    r1: 0,
    c1: 0,
    r2: demo.rows - 1,
    c2: demo.cols - 1,
  }))
    workbook.activeSheet.model.setCell(cell.row, cell.col, cell.raw);
  const notes = workbook.addSheet("Notes");
  notes.model.setCell(0, 0, "Use the + tab to add a worksheet.");
  return workbook;
};
const workbook = restoreWorkbook(window.localStorage, createDemoWorkbook);
const app = new NumeraApp(scene, workbook, {
  onWorkbookChanged: (next) => persistWorkbook(window.localStorage, next),
});

const resize = (): void => {
  const { width, height } = measureSceneContainer(container);
  app.resize(width, height);
};
const observer = new ResizeObserver(resize);
observer.observe(container);
resize();
scene.start();

window.__app = {
  scene,
  get model() {
    return app.model;
  },
  get workbook() {
    return app.workbook;
  },
  app,
  audit: () => auditScene(scene),
};

if (new URLSearchParams(window.location.search).has("debug")) {
  const devtools = attachDevtools(scene, {
    refreshInterval: 0,
    traceEvents: true,
  });
  window.__app.debugTrace = () => devtools.trace?.entries ?? [];
}

window.addEventListener(
  "beforeunload",
  () => {
    persistWorkbook(window.localStorage, app.workbook);
    observer.disconnect();
    app.destroy();
    scene.destroy();
  },
  { once: true },
);
