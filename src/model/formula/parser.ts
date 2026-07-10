import { parseA1, normalizeRect, type Rect } from "../cellRef";
import { lex, LexError, type Token } from "./lexer";

/**
 * Recursive-descent parser. Precedence (low→high): `+ -` < `* /` < `^`
 * (right-assoc) < unary `-` < postfix `%` < primary. `%` is Sheets-style postfix
 * percent (50% → 0.5), NOT modulo.
 */

export type Node =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "ref"; row: number; col: number }
  | { t: "range"; rect: Rect }
  | { t: "bin"; op: "+" | "-" | "*" | "/" | "^" | "&"; l: Node; r: Node }
  | { t: "neg"; e: Node }
  | { t: "pct"; e: Node }
  | { t: "call"; name: string; args: Node[] };

export class ParseError extends Error {}

export function parse(src: string): Node {
  let tokens: Token[];
  try {
    tokens = lex(src);
  } catch (e) {
    throw e instanceof LexError ? new ParseError(e.message) : e;
  }
  let pos = 0;

  const peek = (): Token => tokens[pos];
  const next = (): Token => tokens[pos++];
  const expectOp = (text: string): void => {
    const t = next();
    if (t.type !== "op" || t.text !== text) {
      throw new ParseError(`Expected "${text}" at ${t.pos}`);
    }
  };

  function addExpr(): Node {
    let l = mulExpr();
    while (
      peek().type === "op" &&
      (peek().text === "+" || peek().text === "-" || peek().text === "&")
    ) {
      const op = next().text as "+" | "-" | "&";
      l = { t: "bin", op, l, r: mulExpr() };
    }
    return l;
  }

  function mulExpr(): Node {
    let l = powExpr();
    while (
      peek().type === "op" &&
      (peek().text === "*" || peek().text === "/")
    ) {
      const op = next().text as "*" | "/";
      l = { t: "bin", op, l, r: powExpr() };
    }
    return l;
  }

  function signedExpr(): Node {
    if (peek().type === "op" && peek().text === "-") {
      next();
      return { t: "neg", e: signedExpr() };
    }
    return postfixExpr();
  }

  function powExpr(): Node {
    const base = signedExpr();
    if (peek().type === "op" && peek().text === "^") {
      next();
      return { t: "bin", op: "^", l: base, r: powExpr() }; // right-assoc
    }
    return base;
  }

  function postfixExpr(): Node {
    let e = primary();
    while (peek().type === "op" && peek().text === "%") {
      next();
      e = { t: "pct", e };
    }
    return e;
  }

  function primary(): Node {
    const t = next();
    if (t.type === "num") return { t: "num", v: Number(t.text) };
    if (t.type === "str") return { t: "str", v: t.text };
    if (t.type === "ref") {
      const a = parseA1(t.text);
      if (!a) throw new ParseError(`Bad reference "${t.text}" at ${t.pos}`);
      if (peek().type === "op" && peek().text === ":") {
        next();
        const t2 = next();
        const b = t2.type === "ref" ? parseA1(t2.text) : null;
        if (!b) throw new ParseError(`Bad range end at ${t2.pos}`);
        return { t: "range", rect: normalizeRect(a, b) };
      }
      return { t: "ref", row: a.row, col: a.col };
    }
    if (t.type === "ident") {
      expectOp("(");
      const args: Node[] = [];
      if (!(peek().type === "op" && peek().text === ")")) {
        args.push(addExpr());
        while (peek().type === "op" && peek().text === ",") {
          next();
          args.push(addExpr());
        }
      }
      expectOp(")");
      return { t: "call", name: t.text.toUpperCase(), args };
    }
    if (t.type === "op" && t.text === "(") {
      const e = addExpr();
      expectOp(")");
      return e;
    }
    throw new ParseError(
      `Unexpected token "${t.text || "end of formula"}" at ${t.pos}`,
    );
  }

  const root = addExpr();
  const last = peek();
  if (last.type !== "eof") {
    throw new ParseError(`Unexpected trailing "${last.text}" at ${last.pos}`);
  }
  return root;
}
