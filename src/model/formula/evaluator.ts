import { rectContains, type Rect } from "../cellRef";
import type { Node } from "./parser";
import { FUNCTIONS } from "./functions";

/** Error values are first-class and propagate through operators. */
export interface CellError {
  error: "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#CYCLE!" | "#ERROR!";
}

/** `null` = empty cell (coerces to 0 in arithmetic, skipped by aggregates). */
export type Value = number | string | boolean | null | CellError;

export const isErr = (v: Value): v is CellError =>
  typeof v === "object" && v !== null && "error" in v;

export const err = (code: CellError["error"]): CellError => ({ error: code });

export interface EvalContext {
  /** Resolved value of a cell (never a formula string). May throw CycleError. */
  getCell(row: number, col: number): Value;
}

export class CycleError extends Error {}

function toNumber(v: Value): number | CellError {
  if (v === null) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") {
    if (v.trim() === "") return 0;
    const n = Number(v);
    return Number.isNaN(n) ? err("#VALUE!") : n;
  }
  return v; // already an error
}

/** Flatten an argument node into values; ranges expand to all their cells. */
export function flatten(node: Node, ctx: EvalContext): Value[] {
  if (node.t === "range") {
    const out: Value[] = [];
    forEachInRect(node.rect, (r, c) => out.push(ctx.getCell(r, c)));
    return out;
  }
  return [evaluate(node, ctx)];
}

function forEachInRect(
  rect: Rect,
  fn: (row: number, col: number) => void,
): void {
  for (let r = rect.r1; r <= rect.r2; r++) {
    for (let c = rect.c1; c <= rect.c2; c++) fn(r, c);
  }
}

export function evaluate(node: Node, ctx: EvalContext): Value {
  switch (node.t) {
    case "num":
      return node.v;
    case "str":
      return node.v;
    case "ref":
      return ctx.getCell(node.row, node.col);
    case "range":
      // A bare range outside a function argument has no scalar meaning here.
      return err("#VALUE!");
    case "neg": {
      const v = toNumber(evaluate(node.e, ctx));
      return isErr(v) ? v : -v;
    }
    case "pct": {
      const v = toNumber(evaluate(node.e, ctx));
      return isErr(v) ? v : v / 100;
    }
    case "bin": {
      const l = evaluate(node.l, ctx);
      if (isErr(l)) return l;
      const r = evaluate(node.r, ctx);
      if (isErr(r)) return r;
      if (node.op === "&") return toText(l) + toText(r);
      if (node.op === "+" && (typeof l === "string" || typeof r === "string")) {
        // String + string concatenates only when both sides are non-numeric
        // strings? Sheets errors; numbers still add, otherwise #VALUE!.
        const ln = toNumber(l);
        const rn = toNumber(r);
        if (isErr(ln) || isErr(rn)) return err("#VALUE!");
        return ln + rn;
      }
      const ln = toNumber(l);
      if (isErr(ln)) return ln;
      const rn = toNumber(r);
      if (isErr(rn)) return rn;
      switch (node.op) {
        case "+":
          return ln + rn;
        case "-":
          return ln - rn;
        case "*":
          return ln * rn;
        case "/":
          return rn === 0 ? err("#DIV/0!") : ln / rn;
        case "^":
          return ln ** rn;
      }
    }
    case "call": {
      const fn = FUNCTIONS[node.name];
      if (!fn) return err("#NAME?");
      return fn(node.args, ctx);
    }
  }
}

function toText(value: Value): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/** Collect every cell/range a parsed formula references (for the dep graph). */
export function collectRefs(
  node: Node,
  scalars: Set<string>,
  ranges: Rect[],
): void {
  switch (node.t) {
    case "ref":
      scalars.add(`${node.row}:${node.col}`);
      return;
    case "range":
      ranges.push(node.rect);
      return;
    case "bin":
      collectRefs(node.l, scalars, ranges);
      collectRefs(node.r, scalars, ranges);
      return;
    case "neg":
    case "pct":
      collectRefs(node.e, scalars, ranges);
      return;
    case "call":
      for (const a of node.args) collectRefs(a, scalars, ranges);
      return;
    default:
      return;
  }
}

export { rectContains };
