/**
 * Formül başvuru dönüşümleri (belirteç düzeyinde, boşlukları korur):
 * kopyalamada göreli kaydırma, satır/sütun ekleme-silme, sayfa yeniden
 * adlandırma/silme ve taşıma (kes-yapıştır).
 */
import { EN_SYNTAX, tokenize, type RefPart, type RefToken } from "./lexer";
import { MAX_COLS, MAX_ROWS, colToName, quoteSheetName } from "./references";

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

/** Kopyalanan formüldeki göreli başvuruları kaydırır; taşan başvurular #REF! olur. */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
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
export function adjustFormula(formula: string, formulaSheet: string, ch: StructuralChange): string {
  const target = ch.sheet.toLocaleUpperCase("tr-TR");
  return mapRefs(formula, (ref) => {
    const refSheet = (ref.sheet ?? formulaSheet).toLocaleUpperCase("tr-TR");
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
