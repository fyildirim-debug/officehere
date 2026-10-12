/**
 * Formül denetimi: etkileyen/etkilenen hücreler (Etkileyenleri İzle /
 * Etkilenenleri İzle) ve "Formülü Değerlendir" için adım adım
 * değerlendirme. Değerlendirme her adımda özgün alt ağacı motorun kendi
 * bağlamında hesaplar; gösterimde o alt ağaç sonucuyla değiştirilir.
 */
import type { Node } from "./ast";
import type { FormulaEngine } from "./engine";
import { cellKey, gk, gkLocal, gkSheet, keyCol, keyRow, safeParse } from "./engine-types";
import { createContext, evaluateNode } from "./evaluator";
import { trNameOf } from "./functions/registry";
import { localBoolName, localErrorName } from "./locale";
import { numberToText } from "./numfmt";
import { quoteSheetName } from "./references";
import { renderRefToken } from "./transform";
import { isArray, isError, isRef, type Scalar, type Value } from "./values";

export interface AuditRange {
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Formülün doğrudan başvurduğu hücre ve aralıklar (adlar üzerinden olanlar dahil). */
export function precedentsOf(engine: FormulaEngine, sheet: number, row: number, col: number): AuditRange[] {
  const info = engine.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula;
  if (!info) return [];
  const out: AuditRange[] = [];
  const seen = new Set<string>();
  const add = (r: AuditRange) => {
    const k = `${r.sheet}:${r.r1}:${r.c1}:${r.r2}:${r.c2}`;
    if (seen.has(k)) return;
    seen.add(k);
    out.push(r);
  };
  for (const g of info.cellDeps) {
    const l = gkLocal(g);
    add({ sheet: gkSheet(g), r1: keyRow(l), c1: keyCol(l), r2: keyRow(l), c2: keyCol(l) });
  }
  for (const rd of info.rangeDeps) add({ sheet: rd.sheet, r1: rd.r1, c1: rd.c1, r2: rd.r2, c2: rd.c2 });
  return out;
}

/** Hücreye doğrudan başvuran formül hücreleri. */
export function dependentsOf(engine: FormulaEngine, sheet: number, row: number, col: number): { sheet: number; row: number; col: number }[] {
  const out: { sheet: number; row: number; col: number }[] = [];
  const seen = new Set<number>();
  engine.deps.dependentsOf(gk(sheet, row, col), (f) => {
    if (seen.has(f)) return;
    seen.add(f);
    const l = gkLocal(f);
    out.push({ sheet: gkSheet(f), row: keyRow(l), col: keyCol(l) });
  });
  return out.sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col);
}

/* ------------------------- Formülü Değerlendir ------------------------- */

/** Gösterim ağacı: özgün düğüm + (hesaplandıysa) değeri. */
interface DNode {
  node: Node;
  kids: DNode[];
  value?: Value;
  /** Hesaplanmayacak (ör. ETOPLA'nın aralık bağımsız değişkeni). */
  opaque?: boolean;
}

/** Başvuru bağımsız değişkenlerini değer olarak değil başvuru olarak kullanan fonksiyonlar. */
const REF_FUNCS = new Set(["ROW", "COLUMN", "ROWS", "COLUMNS", "OFFSET", "INDEX", "INDIRECT", "ISREF", "CELL", "AREAS", "SUBTOTAL", "AGGREGATE", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS", "COUNTBLANK", "ISBLANK", "FORMULATEXT", "ISFORMULA", "SHEET", "SHEETS", "LOOKUP", "VLOOKUP", "HLOOKUP", "MATCH", "XLOOKUP", "XMATCH"]);

function build(node: Node, parentFunc?: string): DNode {
  const d: DNode = { node, kids: [] };
  switch (node.type) {
    case "func":
      d.kids = node.args.map((a) => build(a, node.name.toUpperCase()));
      break;
    case "unary":
    case "postfix":
      d.kids = [build(node.arg)];
      break;
    case "binary":
      if (node.op === ":") d.opaque = true;
      else d.kids = [build(node.left), build(node.right)];
      break;
    case "union":
      d.opaque = true;
      break;
    case "ref":
      if (parentFunc && REF_FUNCS.has(parentFunc)) d.opaque = true;
      else if (node.ref.kind !== "cell" && parentFunc) d.opaque = true;
      break;
  }
  return d;
}

function isLiteral(n: Node): boolean {
  return n.type === "num" || n.type === "str" || n.type === "bool" || n.type === "err" || n.type === "missing" || n.type === "array";
}

export interface EvalView {
  /** Gösterilecek metin ("=" ile). */
  text: string;
  /** Altı çizilecek (sıradaki adım) bölüm. */
  mark: [number, number] | null;
  /** Son sonuç hesaplandı mı. */
  done: boolean;
  /** Sıradaki adım, formül içeren bir hücreye başvuru mu (İçeri Adımla). */
  canStepIn: boolean;
}

/** Formülü Değerlendir iletişim kutusunun durum makinesi. */
export class FormulaStepper {
  private root: DNode;
  private sheetName: (id: number) => string;
  constructor(
    private engine: FormulaEngine,
    readonly sheet: number,
    readonly row: number,
    readonly col: number,
    formula: string,
  ) {
    const ast = safeParse(formula.startsWith("=") ? formula.slice(1) : formula);
    this.root = ast ? build(ast) : { node: { type: "err", code: "#NAME?" }, kids: [], value: undefined };
    this.sheetName = (id) => engine.sheetName(id) ?? "";
  }

  /** Değerlendirme sırasına göre sıradaki düğüm (sol→sağ, içten dışa). */
  private next(d: DNode = this.root): DNode | null {
    if (d.value !== undefined) return null;
    if (d.node.type === "func" && d.node.name.toUpperCase() === "IF" && d.kids.length >= 1) {
      // EĞER: koşul → seçilen dal → EĞER
      const c = this.next(d.kids[0]);
      if (c) return c;
      const cond = d.kids[0].value;
      const pick = cond === false || cond === 0 ? 2 : 1;
      const br = d.kids[pick];
      if (br) {
        const b = this.next(br);
        if (b) return b;
      }
      return d;
    }
    for (const k of d.kids) {
      const n = this.next(k);
      if (n) return n;
    }
    if (d.opaque || isLiteral(d.node)) return null;
    return d;
  }

  private evalNode(d: DNode): Value {
    const ctx = createContext({ host: this.engine, sheet: this.sheet, row: this.row, col: this.col, locale: this.engine.locale });
    const v = evaluateNode(d.node, ctx);
    // Tek hücrelik başvuru değeriyle gösterilir
    if (isRef(v) && v.r1 === v.r2 && v.c1 === v.c2) {
      const s = this.engine.getValue(v.sheet, v.r1, v.c1);
      return s === null ? 0 : s;
    }
    return v;
  }

  /** Bir adım ilerler (sıradaki alt ifadeyi ya da son sonucu hesaplar). */
  step(): void {
    const n = this.next();
    if (!n) {
      if (this.root.value === undefined) this.root.value = this.evalNode(this.root);
      return;
    }
    n.value = this.evalNode(n);
  }

  /** Sıradaki adımı (formül içeren hücreye başvuruyu) değeriyle tamamlar — Dışarı Adımla sonrası. */
  completeStep(): void {
    this.step();
  }

  get done(): boolean {
    return this.root.value !== undefined;
  }

  /** Sıradaki adımın başvurduğu hücre (içinde formül varsa). */
  stepInTarget(): { sheet: number; row: number; col: number; formula: string } | null {
    const n = this.next();
    if (!n || n.node.type !== "ref" || n.node.ref.kind !== "cell") return null;
    const t = n.node.ref;
    const sheet = t.sheet !== undefined ? this.engine.sheetIdByName(t.sheet) : this.sheet;
    if (sheet === undefined) return null;
    const f = this.engine.getFormula(sheet, t.a.row, t.a.col);
    return f ? { sheet, row: t.a.row, col: t.a.col, formula: f } : null;
  }

  view(): EvalView {
    const target = this.next();
    let mark: [number, number] | null = null;
    const text = "=" + this.print(this.root, 0, target, (s, e) => (mark = [s + 1, e + 1]));
    return { text, mark, done: this.done, canStepIn: !!this.stepInTarget() };
  }

  /** Bağlamın adresi: "Sayfa1!$B$2". */
  address(): string {
    return `${quoteSheetName(this.sheetName(this.sheet))}!$${colName(this.col)}$${this.row + 1}`;
  }

  private print(d: DNode, prec: number, target: DNode | null, onMark: (s: number, e: number) => void, offset = 0): string {
    if (d.value !== undefined) return valueText(d.value);
    const n = d.node;
    let out = "";
    const child = (k: DNode, p: number) => {
      const s = this.print(k, p, target, onMark, offset + out.length);
      out += s;
    };
    const myPrec = precOf(n);
    const paren = myPrec < prec;
    if (paren) out += "(";
    switch (n.type) {
      case "num":
        out += numberToText(n.value);
        break;
      case "str":
        out += `"${n.value.replace(/"/g, '""')}"`;
        break;
      case "bool":
        out += localBoolName(n.value);
        break;
      case "err":
        out += localErrorName(n.code);
        break;
      case "ref":
        out += renderRefToken(n.ref, "tr");
        break;
      case "name":
        out += n.name;
        break;
      case "missing":
        break;
      case "func":
        out += trNameOf(n.name.toUpperCase()) + "(";
        d.kids.forEach((k, i) => {
          if (i) out += ";";
          child(k, 0);
        });
        out += ")";
        break;
      case "unary":
        out += n.op;
        child(d.kids[0], myPrec);
        break;
      case "postfix":
        child(d.kids[0], myPrec);
        out += "%";
        break;
      case "binary":
        if (d.opaque) {
          out += printPlain(n);
          break;
        }
        child(d.kids[0], myPrec);
        out += n.op;
        child(d.kids[1], myPrec + 1);
        break;
      case "union":
      case "array":
        out += printPlain(n);
        break;
    }
    if (paren) out += ")";
    if (d === target) onMark(offset, offset + out.length);
    return out;
  }
}

function colName(c: number): string {
  let s = "";
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function precOf(n: Node): number {
  if (n.type === "binary") {
    switch (n.op) {
      case ":":
        return 10;
      case " ":
        return 9;
      case "^":
        return 6;
      case "*":
      case "/":
        return 5;
      case "+":
      case "-":
        return 4;
      case "&":
        return 3;
      default:
        return 2;
    }
  }
  if (n.type === "unary") return 8;
  if (n.type === "postfix") return 7;
  return 11;
}

function printPlain(n: Node): string {
  switch (n.type) {
    case "ref":
      return renderRefToken(n.ref, "tr");
    case "binary":
      return printPlain(n.left) + n.op + printPlain(n.right);
    case "union":
      return "(" + n.items.map(printPlain).join(";") + ")";
    case "array":
      return "{" + n.rows.map((r) => r.map(printPlain).join("\\")).join(";") + "}";
    case "num":
      return numberToText(n.value);
    case "str":
      return `"${n.value}"`;
    case "bool":
      return localBoolName(n.value);
    case "err":
      return localErrorName(n.code);
    case "name":
      return n.name;
    case "func":
      return trNameOf(n.name.toUpperCase()) + "(" + n.args.map(printPlain).join(";") + ")";
    case "unary":
      return n.op + printPlain(n.arg);
    case "postfix":
      return printPlain(n.arg) + "%";
    case "missing":
      return "";
  }
}

/** Değerin formül içinde gösterimi (Türkçe). */
export function valueText(v: Value): string {
  if (v === null) return '""';
  if (typeof v === "number") return numberToText(v);
  if (typeof v === "boolean") return localBoolName(v);
  if (typeof v === "string") return `"${v.replace(/"/g, '""')}"`;
  if (isError(v)) return localErrorName(v.code);
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return `$${colName(v.c1)}$${v.r1 + 1}`;
    return `$${colName(v.c1)}$${v.r1 + 1}:$${colName(v.c2)}$${v.r2 + 1}`;
  }
  if (isArray(v)) {
    const rows = v.data.slice(0, 20).map((r) => r.slice(0, 20).map((s: Scalar) => valueText(s)).join("\\"));
    return "{" + rows.join(";") + (v.data.length > 20 ? ";..." : "") + "}";
  }
  return "#DEĞER!";
}
