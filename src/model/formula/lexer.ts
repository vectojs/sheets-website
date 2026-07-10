/** Formula tokenizer. Input is the text AFTER the leading `=`. */

export type TokenType = "num" | "str" | "ref" | "ident" | "op" | "eof";

export interface Token {
  type: TokenType;
  text: string;
  pos: number;
}

const REF_RE = /^\$?[A-Z]{1,3}\$?\d+/;
const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*/;
const NUM_RE = /^\d+(\.\d+)?|^\.\d+/;
const OPS = new Set(["+", "-", "*", "/", "^", "%", "&", "(", ")", ",", ":"]);

export class LexError extends Error {}

export function lex(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === " " || ch === "\t") {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let out = "";
      while (j < src.length) {
        if (src[j] === '"' && src[j + 1] === '"') {
          out += '"'; // doubled quote escapes a quote, Sheets-style
          j += 2;
        } else if (src[j] === '"') {
          break;
        } else {
          out += src[j];
          j++;
        }
      }
      if (j >= src.length) throw new LexError(`Unterminated string at ${i}`);
      tokens.push({ type: "str", text: out, pos: i });
      i = j + 1;
      continue;
    }
    if (OPS.has(ch)) {
      tokens.push({ type: "op", text: ch, pos: i });
      i++;
      continue;
    }
    const rest = src.slice(i);
    const num = NUM_RE.exec(rest);
    if (num) {
      tokens.push({ type: "num", text: num[0], pos: i });
      i += num[0].length;
      continue;
    }
    const ref = REF_RE.exec(rest);
    // A token like "A1" is a ref only when not followed by more identifier
    // chars ("A1B" is an identifier, e.g. a function name).
    if (ref && !/[A-Za-z0-9_]/.test(rest[ref[0].length] ?? "")) {
      tokens.push({ type: "ref", text: ref[0].replaceAll("$", ""), pos: i });
      i += ref[0].length;
      continue;
    }
    const ident = IDENT_RE.exec(rest);
    if (ident) {
      tokens.push({ type: "ident", text: ident[0], pos: i });
      i += ident[0].length;
      continue;
    }
    throw new LexError(`Unexpected character "${ch}" at ${i}`);
  }
  tokens.push({ type: "eof", text: "", pos: src.length });
  return tokens;
}
