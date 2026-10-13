/**
 * Excel işleç önceliğiyle Pratt ayrıştırıcısı.
 * Öncelik (düşükten yükseğe): karşılaştırma, &, + -, * /, ^, %, tekli - +, boşluk (kesişim), : (aralık).
 * Excel'de tekli eksi ^'den önce uygulanır: =-2^2 → 4.
 */
import { CALL_FN, ParseError, type BinaryOp, type Node } from "./ast";
import { EN_SYNTAX, tokenize, type SyntaxLocale, type Token } from "./lexer";
import type { ErrorCode } from "./values";

const BINARY_PREC: Record<string, number> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
};
const PERCENT_PREC = 6;
const UNARY_PREC = 7;

export type FunctionNameResolver = (raw: string) => string | undefined;

/** "_xlfn." gibi önekleri atar ve büyük harfe çevirir. */
export function stripFunctionPrefix(raw: string): string {
  return raw.replace(/^(_xlfn\.|_xlws\.|_xlpm\.)+/i, "");
}

export interface ParseOptions {
  syntax?: SyntaxLocale;
  resolveFunction?: FunctionNameResolver;
}

class Parser {
  private i = 0;
  constructor(
    private readonly tokens: Token[],
    private readonly resolve: FunctionNameResolver,
  ) {}

  private peek(): Token | undefined {
    return this.tokens[this.i];
  }

  private next(): Token {
    const t = this.tokens[this.i++];
    if (!t) throw new ParseError("Beklenmeyen formül sonu", this.endPos());
    return t;
  }

  private endPos(): number {
    const last = this.tokens[this.tokens.length - 1];
    return last ? last.pos + last.text.length : 0;
  }

  parseFormula(): Node {
    if (this.tokens.length === 0) throw new ParseError("Boş formül", 0);
    const node = this.parseExpr(0);
    const t = this.peek();
    if (t) throw new ParseError(`Beklenmeyen "${t.text}"`, t.pos);
    return node;
  }

  private parseExpr(minPrec: number): Node {
    let left = this.parsePrefix();
    for (;;) {
      const t = this.peek();
      if (!t || t.type !== "op") break;
      if (t.text === "%") {
        if (PERCENT_PREC < minPrec) break;
        this.i++;
        left = { type: "postfix", op: "%", arg: left };
        continue;
      }
      const prec = BINARY_PREC[t.text];
      if (prec === undefined || prec < minPrec) break;
      this.i++;
      const right = this.parseExpr(prec + 1);
      left = { type: "binary", op: t.text as BinaryOp, left, right };
    }
    return left;
  }

  private parsePrefix(): Node {
    const t = this.peek();
    if (t && t.type === "op" && (t.text === "-" || t.text === "+")) {
      this.i++;
      const arg = this.parseExpr(UNARY_PREC);
      return { type: "unary", op: t.text as "-" | "+", arg };
    }
    if (t && t.type === "op" && t.text === "@") {
      this.i++;
      return this.parseExpr(UNARY_PREC);
    }
    return this.parseIntersect();
  }

  /** Kesişim işleci (boşluk): "A1:C3 B2:D4". */
  private parseIntersect(): Node {
    let left = this.parseRange();
    while (this.peek()?.type === "isect") {
      this.i++;
      const right = this.parseRange();
      left = { type: "binary", op: " ", left, right };
    }
    return left;
  }

  /** Aralık işleci (:) en yüksek önceliktedir. */
  private parseRange(): Node {
    let left = this.parsePrimary();
    for (;;) {
      const t = this.peek();
      if (t && t.type === "op" && t.text === ":") {
        this.i++;
        const right = this.parsePrimary();
        left = { type: "binary", op: ":", left, right };
      } else break;
    }
    return left;
  }

  private parsePrimary(): Node {
    const t = this.next();
    switch (t.type) {
      case "number":
        return { type: "num", value: t.value as number };
      case "string":
        return { type: "str", value: t.value as string };
      case "bool":
        return { type: "bool", value: t.value as boolean };
      case "error":
        return { type: "err", code: t.value as ErrorCode };
      case "ref":
        return { type: "ref", ref: t.ref! };
      case "structref":
        return { type: "structref", ref: t.sref! };
      case "name":
        return { type: "name", name: t.value as string };
      case "func": {
        let node = this.parseCall(t);
        // Doğrudan çağrı: LAMBDA(x;x*2)(4) — parantezli bağımsız değişken listesi işlevin ardından gelir
        while (this.peek()?.type === "lparen") {
          const args = this.parseArgs();
          node = { type: "func", name: CALL_FN, args: [node, ...args] };
        }
        return node;
      }
      case "lparen": {
        const first = this.parseExpr(0);
        const items = [first];
        while (this.peek()?.type === "sep") {
          this.i++;
          items.push(this.parseExpr(0));
        }
        const close = this.next();
        if (close.type !== "rparen") throw new ParseError('")" bekleniyordu', close.pos);
        return items.length > 1 ? { type: "union", items } : first;
      }
      case "lbrace":
        return this.parseArray();
      default:
        throw new ParseError(`Beklenmeyen "${t.text}"`, t.pos);
    }
  }

  private parseCall(t: Token): Node {
    const raw = stripFunctionPrefix(t.value as string);
    // Bilinmeyen ad (LAMBDA tanımlı ad ya da LET değişkeni olabilir) Türkçe büyük harfe çevrilir
    const name = this.resolve(raw) ?? raw.toLocaleUpperCase("tr-TR");
    return { type: "func", name, args: this.parseArgs() };
  }

  /** "(" ... ")" bağımsız değişken listesi (boş bağımsız değişkenler "missing"). */
  private parseArgs(): Node[] {
    const lp = this.next();
    if (lp.type !== "lparen") throw new ParseError('"(" bekleniyordu', lp.pos);
    const args: Node[] = [];
    if (this.peek()?.type === "rparen") {
      this.i++;
      return args;
    }
    for (;;) {
      const p = this.peek();
      if (!p) throw new ParseError('")" bekleniyordu', this.endPos());
      if (p.type === "sep" || p.type === "rparen") {
        args.push({ type: "missing" });
      } else {
        args.push(this.parseExpr(0));
      }
      const sep = this.next();
      if (sep.type === "rparen") break;
      if (sep.type !== "sep") throw new ParseError(`Beklenmeyen "${sep.text}"`, sep.pos);
    }
    return args;
  }

  private parseArray(): Node {
    const rows: Node[][] = [[]];
    for (;;) {
      let t = this.next();
      let sign = 1;
      while (t.type === "op" && (t.text === "-" || t.text === "+")) {
        if (t.text === "-") sign = -sign;
        t = this.next();
      }
      let item: Node;
      switch (t.type) {
        case "number":
          item = { type: "num", value: sign * (t.value as number) };
          break;
        case "string":
          item = { type: "str", value: t.value as string };
          break;
        case "bool":
          item = { type: "bool", value: t.value as boolean };
          break;
        case "error":
          item = { type: "err", code: t.value as ErrorCode };
          break;
        default:
          throw new ParseError("Dizi sabitinde geçersiz öğe", t.pos);
      }
      rows[rows.length - 1].push(item);
      const sep = this.next();
      if (sep.type === "rbrace") break;
      if (sep.type === "arrayCol" || sep.type === "sep") continue;
      if (sep.type === "arrayRow") {
        rows.push([]);
        continue;
      }
      throw new ParseError("Dizi sabitinde beklenmeyen ayırıcı", sep.pos);
    }
    const w = rows[0].length;
    if (rows.some((r) => r.length !== w)) throw new ParseError("Dizi sabiti satırları eşit uzunlukta olmalı", 0);
    return { type: "array", rows };
  }
}

const OPERAND_END = new Set(["ref", "structref", "name", "rparen"]);
const REF_START = new Set(["ref", "structref", "name", "func", "lparen"]);

/** Belirteç listesini AST'ye çevirir; boşluklar atlanır (iki başvuru arasındaki boşluk kesişim işlecidir). */
export function parseTokens(tokens: Token[], resolve: FunctionNameResolver = (n) => n.toUpperCase()): Node {
  const significant: Token[] = [];
  tokens.forEach((t, i) => {
    if (t.type !== "ws") {
      significant.push(t);
      return;
    }
    const prev = significant[significant.length - 1];
    const next = tokens.slice(i + 1).find((x) => x.type !== "ws");
    if (prev && next && OPERAND_END.has(prev.type) && REF_START.has(next.type)) significant.push({ ...t, type: "isect" });
  });
  const bad = significant.find((t) => t.type === "unknown");
  if (bad) throw new ParseError(`Geçersiz karakter "${bad.text}"`, bad.pos);
  return new Parser(significant, resolve).parseFormula();
}

/** Formül metnini ("=" olmadan) ayrıştırır. */
export function parse(src: string, options: ParseOptions = {}): Node {
  const text = src.startsWith("=") ? src.slice(1) : src;
  return parseTokens(tokenize(text, options.syntax ?? EN_SYNTAX), options.resolveFunction);
}
