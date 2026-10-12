/**
 * Formül başvuru dönüşümleri (belirteç düzeyinde, boşlukları korur):
 * kopyalamada göreli kaydırma, satır/sütun ekleme-silme, sayfa yeniden
 * adlandırma/silme ve taşıma (kes-yapıştır).
 */
import { EN_SYNTAX, tokenize, type RefPart, type RefToken } from "./lexer";
import { MAX_COLS, MAX_ROWS, colToName, nameToCol, quoteSheetName } from "./references";
import { renameTableInRef, renderStructRef, type StructRef } from "./structref";
import type { Node } from "./ast";

/** Yapısal başvuru belirteçlerine işlev uygulayıp formülü yeniden yazar. */
function mapStructRefs(formula: string, fn: (r: StructRef) => StructRef): string {
  if (!formula.includes("[")) return formula;
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  let changed = false;
  const out = tokenize(body, EN_SYNTAX)
    .map((t) => {
      if (t.type !== "structref" || !t.sref) return t.text;
      const r = fn(t.sref);
      if (r === t.sref) return t.text;
      changed = true;
      return renderStructRef(r, "en");
    })
    .join("");
  return changed ? (hasEq ? "=" : "") + out : formula;
}

/** Tablo adını formüldeki yapısal başvurularda değiştirir. */
export function renameTableInFormula(formula: string, oldName: string, newName: string): string {
  return mapStructRefs(formula, (r) => renameTableInRef(r, oldName, newName));
}

/** Tablo sütunu yeniden adlandırılınca yapısal başvuruları günceller. */
export function renameTableColumnInFormula(formula: string, table: string, oldCol: string, newCol: string): string {
  const T = table.toLocaleUpperCase("tr-TR");
  const O = oldCol.toLocaleUpperCase("tr-TR");
  return mapStructRefs(formula, (r) => {
    if (!r.table || r.table.toLocaleUpperCase("tr-TR") !== T) return r;
    const c1 = r.col1 !== undefined && r.col1.toLocaleUpperCase("tr-TR") === O ? newCol : r.col1;
    const c2 = r.col2 !== undefined && r.col2.toLocaleUpperCase("tr-TR") === O ? newCol : r.col2;
    return c1 === r.col1 && c2 === r.col2 ? r : { ...r, col1: c1, col2: c2 };
  });
}

/** Tablosuz yapısal başvurulara ([@Adet]) tablo adını ekler (dosyada tablo adı zorunludur). */
export function qualifyStructRefs(formula: string, tableName: string): string {
  return mapStructRefs(formula, (r) => (r.table ? r : { ...r, table: tableName }));
}

export function renderPart(p: RefPart, kind: RefToken["kind"]): string {
  if (kind === "cols") return `${p.colAbs ? "$" : ""}${colToName(p.col)}`;
  if (kind === "rows") return `${p.rowAbs ? "$" : ""}${p.row + 1}`;
  return `${p.colAbs ? "$" : ""}${colToName(p.col)}${p.rowAbs ? "$" : ""}${p.row + 1}`;
}

export function renderRefToken(r: RefToken, locale: "tr" | "en" = "en"): string {
  const prefix = r.sheet !== undefined ? `${quoteSheetName(r.sheet)}!` : "";
  if (r.refError) return prefix + (locale === "tr" ? "#BAŞV!" : "#REF!");
  if (r.kind === "cell") return prefix + renderPart(r.a, "cell");
  return `${prefix}${renderPart(r.a, r.kind)}:${renderPart(r.b!, r.kind)}`;
}

/** Her başvuru belirtecine işlev uygulayıp formülü yeniden yazar. */
export function mapRefs(formula: string, fn: (r: RefToken) => RefToken | null): string {
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(body, EN_SYNTAX);
  let out = "";
  let changed = false;
  for (const t of tokens) {
    if (t.type === "ref" && t.ref && !t.ref.refError) {
      const r = fn(t.ref);
      if (r === t.ref) {
        out += t.text;
        continue;
      }
      changed = true;
      out += r === null ? renderRefToken({ ...t.ref, refError: true }) : renderRefToken(r);
    } else out += t.text;
  }
  if (!changed) return formula;
  return (hasEq ? "=" : "") + out;
}

function cloneRef(r: RefToken): RefToken {
  return { ...r, a: { ...r.a }, b: r.b ? { ...r.b } : undefined };
}

const SIMPLE_REF = /(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\w(.])/g;

/**
 * Hızlı yol: metin, sayfa adı, tablo, hata ve satır/sütun aralığı içermeyen
 * yalın formüllerde A1 başvurularını belirteçlere ayırmadan kaydırır.
 * Uygun değilse ya da taşma olursa null döner.
 */
function fastShift(formula: string, dRow: number, dCol: number): string | null {
  if (/["'!\[#{]/.test(formula) || /(^|[^A-Za-z$\d])\$?\d+\s*:|:\s*\$?[A-Za-z]{1,3}(?![\d$A-Za-z])/.test(formula) || /[a-z]/.test(formula.replace(/[A-Z_][A-Za-z0-9_.]*\(/g, ""))) return null;
  let bad = false;
  const out = formula.replace(SIMPLE_REF, (m, ca: string, c: string, ra: string, r: string) => {
    let col = nameToCol(c);
    let row = Number(r) - 1;
    if (!ca) col += dCol;
    if (!ra) row += dRow;
    if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
      bad = true;
      return m;
    }
    return `${ca}${colToName(col)}${ra}${row + 1}`;
  });
  return bad ? null : out;
}

/** Kopyalanan formüldeki göreli başvuruları kaydırır; taşan başvurular #REF! olur. */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  const fast = fastShift(formula, dRow, dCol);
  if (fast !== null) return fast;
  return mapRefs(formula, (ref) => {
    const r = cloneRef(ref);
    const parts = r.b ? [r.a, r.b] : [r.a];
    for (const p of parts) {
      if (r.kind !== "cols" && !p.rowAbs) {
        p.row += dRow;
        if (p.row < 0 || p.row >= MAX_ROWS) return null;
      }
      if (r.kind !== "rows" && !p.colAbs) {
        p.col += dCol;
        if (p.col < 0 || p.col >= MAX_COLS) return null;
      }
    }
    return r;
  });
}

export interface StructuralChange {
  /** Değişikliğin yapıldığı sayfa adı. */
  sheet: string;
  axis: "row" | "col";
  at: number;
  /** Pozitif: ekleme, negatif: silme. */
  count: number;
}

/** Tek eksen koordinatını ekleme/silmeye göre eşler; silinen için null. */
function mapIndex(i: number, at: number, count: number): number | null {
  if (count > 0) return i >= at ? i + count : i;
  const n = -count;
  if (i < at) return i;
  if (i < at + n) return null;
  return i - n;
}

/** Satır/sütun ekleme-silmede başvuruları günceller. formulaSheet: formülün bulunduğu sayfa. */
const upperCache = new Map<string, string>();
/** Sayfa adının büyük harf karşılığı (önbellekli; yerel ayarlı dönüşüm pahalıdır). */
function upperName(s: string): string {
  let u = upperCache.get(s);
  if (u === undefined) {
    if (upperCache.size > 1000) upperCache.clear();
    u = s.toLocaleUpperCase("tr-TR");
    upperCache.set(s, u);
  }
  return u;
}

export function adjustFormula(formula: string, formulaSheet: string, ch: StructuralChange): string {
  const target = upperName(ch.sheet);
  const own = upperName(formulaSheet);
  return mapRefs(formula, (ref) => {
    const refSheet = ref.sheet === undefined ? own : upperName(ref.sheet);
    if (refSheet !== target) return ref;
    if (ch.axis === "row" && ref.kind === "cols") return ref;
    if (ch.axis === "col" && ref.kind === "rows") return ref;
    const r = cloneRef(ref);
    const key = ch.axis === "row" ? "row" : "col";
    if (!r.b) {
      const m = mapIndex(r.a[key], ch.at, ch.count);
      if (m === null) return null;
      if (m === r.a[key]) return ref;
      r.a[key] = m;
      return r;
    }
    // Aralık: uç noktaları sıralı kabul et
    const lo = Math.min(r.a[key], r.b[key]);
    const hi = Math.max(r.a[key], r.b[key]);
    let nlo: number;
    let nhi: number;
    if (ch.count > 0) {
      nlo = lo >= ch.at ? lo + ch.count : lo;
      nhi = hi >= ch.at ? hi + ch.count : hi;
    } else {
      const n = -ch.count;
      const end = ch.at + n - 1;
      if (lo >= ch.at && hi <= end) return null;
      nlo = lo < ch.at ? lo : lo > end ? lo - n : ch.at;
      nhi = hi < ch.at ? hi : hi > end ? hi - n : ch.at - 1;
      if (nhi < nlo) return null;
    }
    const max = ch.axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
    nhi = Math.min(nhi, max);
    if (nlo > max) return null;
    if (nlo === lo && nhi === hi) return ref;
    const aIsLo = r.a[key] <= r.b[key];
    (aIsLo ? r.a : r.b)[key] = nlo;
    (aIsLo ? r.b : r.a)[key] = nhi;
    return r;
  });
}

/** Sayfa yeniden adlandırıldığında başvuruları günceller. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const o = oldName.toLocaleUpperCase("tr-TR");
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(body, EN_SYNTAX);
  let changed = false;
  const out = tokens
    .map((t) => {
      if (t.type === "ref" && t.ref?.sheet !== undefined && t.ref.sheet.toLocaleUpperCase("tr-TR") === o) {
        changed = true;
        return renderRefToken({ ...t.ref, sheet: newName });
      }
      return t.text;
    })
    .join("");
  return changed ? (hasEq ? "=" : "") + out : formula;
}

/** Silinen sayfaya yapılan başvuruları #REF! yapar. */
export function removeSheetFromFormula(formula: string, name: string): string {
  const o = name.toLocaleUpperCase("tr-TR");
  return mapRefs(formula, (ref) => (ref.sheet !== undefined && ref.sheet.toLocaleUpperCase("tr-TR") === o ? null : ref));
}

export interface MoveSpec {
  srcSheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  dRow: number;
  dCol: number;
  dstSheet: string;
}

/** Kes-yapıştırda taşınan hücrelere yapılan başvuruları yeni konuma çevirir. */
export function moveRefsInFormula(formula: string, formulaSheet: string, mv: MoveSpec): string {
  const src = mv.srcSheet.toLocaleUpperCase("tr-TR");
  return mapRefs(formula, (ref) => {
    if (ref.kind === "cols" || ref.kind === "rows") return ref;
    const refSheet = (ref.sheet ?? formulaSheet).toLocaleUpperCase("tr-TR");
    if (refSheet !== src) return ref;
    const parts = ref.b ? [ref.a, ref.b] : [ref.a];
    const inside = parts.every((p) => p.row >= mv.r1 && p.row <= mv.r2 && p.col >= mv.c1 && p.col <= mv.c2);
    if (!inside) return ref;
    const r = cloneRef(ref);
    for (const p of r.b ? [r.a, r.b] : [r.a]) {
      p.row += mv.dRow;
      p.col += mv.dCol;
      if (p.row < 0 || p.col < 0 || p.row >= MAX_ROWS || p.col >= MAX_COLS) return null;
    }
    if (mv.dstSheet.toLocaleUpperCase("tr-TR") !== src) r.sheet = mv.dstSheet;
    return r;
  });
}

/** Formüldeki başvuruların listesi (renkli başvuru vurguları için). */
export interface FormulaRefSpan {
  start: number;
  end: number;
  ref: RefToken;
}

export function findRefSpans(formula: string, syntax = EN_SYNTAX): FormulaRefSpan[] {
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  const off = hasEq ? 1 : 0;
  return tokenize(body, syntax)
    .filter((t) => t.type === "ref" && t.ref && !t.ref.refError)
    .map((t) => ({ start: t.pos + off, end: t.pos + off + t.text.length, ref: t.ref! }));
}

/** F4: başvurunun mutlak/göreli döngüsü ($A$1 → A$1 → $A1 → A1). */
export function cycleAbsolute(ref: RefToken): RefToken {
  const r = cloneRef(ref);
  const parts = r.b ? [r.a, r.b] : [r.a];
  const p0 = r.a;
  let next: [boolean, boolean];
  if (r.kind === "cols") next = p0.colAbs ? [false, false] : [false, true];
  else if (r.kind === "rows") next = p0.rowAbs ? [false, false] : [true, false];
  else if (p0.colAbs && p0.rowAbs) next = [true, false];
  else if (!p0.colAbs && p0.rowAbs) next = [false, true];
  else if (p0.colAbs && !p0.rowAbs) next = [false, false];
  else next = [true, true];
  for (const p of parts) {
    p.rowAbs = next[0];
    p.colAbs = next[1];
  }
  return r;
}

/**
 * Ayrıştırılmış formülü göreli olarak kaydırır (shiftFormula'nın AST karşılığı;
 * yeniden ayrıştırmadan kopyalama/sıralama için). Taşan başvuru #REF! olur.
 */
export function shiftAst(node: Node, dRow: number, dCol: number): Node {
  if (dRow === 0 && dCol === 0) return node;
  const shiftRef = (ref: RefToken): RefToken => {
    if (ref.refError) return ref;
    const r = cloneRef(ref);
    const parts = r.b ? [r.a, r.b] : [r.a];
    for (const p of parts) {
      if (r.kind !== "cols" && !p.rowAbs) {
        p.row += dRow;
        if (p.row < 0 || p.row >= MAX_ROWS) return { ...ref, refError: true };
      }
      if (r.kind !== "rows" && !p.colAbs) {
        p.col += dCol;
        if (p.col < 0 || p.col >= MAX_COLS) return { ...ref, refError: true };
      }
    }
    return r;
  };
  const rec = (n: Node): Node => {
    switch (n.type) {
      case "ref":
        return { type: "ref", ref: shiftRef(n.ref) };
      case "func":
        return { ...n, args: n.args.map(rec) };
      case "unary":
      case "postfix":
        return { ...n, arg: rec(n.arg) } as Node;
      case "binary":
        return { ...n, left: rec(n.left), right: rec(n.right) };
      case "union":
        return { ...n, items: n.items.map(rec) };
      default:
        return n;
    }
  };
  return rec(node);
}
