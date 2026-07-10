import { describe, it, expect } from "bun:test";
import { parse, ParseError } from "../src/model/formula/parser";
import {
  evaluate,
  isErr,
  type Value,
  type EvalContext,
} from "../src/model/formula/evaluator";

/** Evaluate a formula against a tiny fixed sheet. */
function run(src: string, cells: Record<string, Value> = {}): Value {
  const ctx: EvalContext = {
    getCell: (r, c) => cells[`${r}:${c}`] ?? null,
  };
  return evaluate(parse(src), ctx);
}

describe("parser", () => {
  it("honors precedence and associativity", () => {
    expect(run("2+3*4")).toBe(14);
    expect(run("(2+3)*4")).toBe(20);
    expect(run("2^3^2")).toBe(512); // right-assoc: 2^(3^2)
    expect(run("-2^2")).toBe(4); // unary binds outside ^: (-2)^2
    expect(run("2*-3")).toBe(-6);
  });

  it("treats % as postfix percent, not modulo", () => {
    expect(run("50%")).toBe(0.5);
    expect(run("200*10%")).toBe(20);
  });

  it("parses strings with doubled-quote escapes", () => {
    expect(run('"he said ""hi"""')).toBe('he said "hi"');
  });

  it("concatenates scalars with the Sheets ampersand operator", () => {
    expect(run('"Q"&A1&2', { "0:0": "1" })).toBe("Q12");
    expect(run('A2&"x"', {})).toBe("x");
  });

  it("rejects malformed input with ParseError", () => {
    for (const bad of ["2+", "SUM(1", ")", "2 3", "@#!"]) {
      expect(() => parse(bad)).toThrow(ParseError);
    }
  });
});

describe("evaluator", () => {
  it("resolves refs and coerces empties to 0 in arithmetic", () => {
    expect(run("A1+B1", { "0:0": 7, "0:1": 3 })).toBe(10);
    expect(run("A1+5", {})).toBe(5); // empty ref → 0
  });

  it("propagates errors through operators", () => {
    const v = run("1/0+5");
    expect(isErr(v) && v.error).toBe("#DIV/0!");
  });

  it("returns #VALUE! for non-numeric strings in arithmetic", () => {
    const v = run("A1*2", { "0:0": "abc" });
    expect(isErr(v) && v.error).toBe("#VALUE!");
  });

  it("returns #NAME? for unknown functions", () => {
    const v = run("NOPE(1)");
    expect(isErr(v) && v.error).toBe("#NAME?");
  });
});

describe("functions", () => {
  const sheet: Record<string, Value> = {
    "0:0": 1, // A1
    "1:0": 2, // A2
    "2:0": 3, // A3
    "3:0": "x", // A4 — strings are skipped by aggregates
  };

  it("SUM/AVG/MIN/MAX/COUNT over ranges skip strings and empties", () => {
    expect(run("SUM(A1:A5)", sheet)).toBe(6);
    expect(run("AVG(A1:A3)", sheet)).toBe(2);
    expect(run("AVERAGE(A1:A3)", sheet)).toBe(2);
    expect(run("MIN(A1:A3)", sheet)).toBe(1);
    expect(run("MAX(A1:A3)", sheet)).toBe(3);
    expect(run("COUNT(A1:A5)", sheet)).toBe(3);
  });

  it("SUM accepts mixed scalar and range args", () => {
    expect(run("SUM(A1:A3, 10, A1)", sheet)).toBe(17);
  });

  it("IF branches on numeric truthiness and passes errors through", () => {
    expect(run('IF(A1, "yes", "no")', sheet)).toBe("yes");
    expect(run('IF(A9, "yes", "no")', sheet)).toBe("no"); // empty → falsy
    const v = run("IF(1/0, 1, 2)");
    expect(isErr(v) && v.error).toBe("#DIV/0!");
  });

  it("CONCAT joins values, skipping empties", () => {
    expect(run('CONCAT("a", A1, A9, "b")', sheet)).toBe("a1b");
  });

  it("AVG of no numbers is #DIV/0!", () => {
    const v = run("AVG(A9:A10)", {});
    expect(isErr(v) && v.error).toBe("#DIV/0!");
  });
});
