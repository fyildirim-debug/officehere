/**
 * Fonksiyon gerçeklemeleri için ortak yardımcılar. Hata değerleri
 * FormulaError fırlatılarak bildirilir; değerlendirici yakalar.
 */
import type { FnContext } from "../context";
import { makeCriteria, toBool, toNumber, toText } from "../coerce";
import { Errors, FormulaError, isArray, isError, isLambda, isMultiRef, isRef, makeArray, type ArrayValue, type RefValue, type Scalar, type Value } from "../values";

/** Hata ise fırlatır. */
export function check<T>(v: T | FormulaError): T {
  if (v instanceof FormulaError) throw v;
  return v;
}

const HUGE = 2_000_000;

/** Başvuruyu kullanılan alana kırpar (boşsa null). */
export function clipRef(ref: RefValue, ctx: FnContext): RefValue | null {
  const ext = ctx.host.getUsedExtent(ref.sheet);
  const r2 = Math.min(ref.r2, ext.rows - 1);
  const c2 = Math.min(ref.c2, ext.cols - 1);
  if (r2 < ref.r1 || c2 < ref.c1) return null;
  return { ...ref, r2, c2 };
}

/** Başvurudaki değerleri 2B diziye çevirir. clip: kullanılan alana kırp. */
export function refValues(ref: RefValue, ctx: FnContext, clip = false): Scalar[][] {
  let r = ref;
  const area = (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1);
  if (clip || area > HUGE) {
    const c = clipRef(ref, ctx);
    if (!c) return [];
    r = c;
  }
  const out: Scalar[][] = [];
  const host = ctx.host;
  for (let i = r.r1; i <= r.r2; i++) {
    const row: Scalar[] = new Array(r.c2 - r.c1 + 1);
    for (let j = r.c1; j <= r.c2; j++) row[j - r.c1] = host.getCellValue(r.sheet, i, j);
    out.push(row);
  }
  return out;
}

/** Değeri skalere indirger (tek hücreli başvuru → değeri; dizi → ilk öğe). */
export function scalar(v: Value, ctx: FnContext): Scalar {
  if (v === null || typeof v !== "object" || isError(v)) return v as Scalar;
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return ctx.host.getCellValue(v.sheet, v.r1, v.c1);
    // Örtük kesişim: formül hücresiyle aynı satır/sütun
    if (v.c1 === v.c2 && ctx.sheet === v.sheet && ctx.row >= v.r1 && ctx.row <= v.r2) return ctx.host.getCellValue(v.sheet, ctx.row, v.c1);
    if (v.r1 === v.r2 && ctx.sheet === v.sheet && ctx.col >= v.c1 && ctx.col <= v.c2) return ctx.host.getCellValue(v.sheet, v.r1, ctx.col);
    return Errors.VALUE;
  }
  if (isArray(v)) return v.data[0]?.[0] ?? null;
  if (isMultiRef(v) || isLambda(v)) return Errors.VALUE;
  return Errors.VALUE;
}

/** Değeri 2B diziye çevirir. */
export function matrix(v: Value, ctx: FnContext, clip = false): Scalar[][] {
  if (isRef(v)) return refValues(v, ctx, clip);
  if (isArray(v)) return v.data;
  if (isMultiRef(v)) throw Errors.VALUE;
  if (isLambda(v)) throw Errors.VALUE;
  return [[v as Scalar]];
}

/** Değer dizi/aralık mı (birden çok öğe)? */
export function isMulti(v: Value): boolean {
  if (isRef(v)) return v.r1 !== v.r2 || v.c1 !== v.c2;
  if (isArray(v)) return v.data.length > 1 || (v.data[0]?.length ?? 0) > 1;
  return isMultiRef(v);
}

export function num(v: Value, ctx: FnContext): number {
  return check(toNumber(check(scalarNoErr(v, ctx)), ctx.locale));
}

function scalarNoErr(v: Value, ctx: FnContext): Scalar {
  const s = scalar(v, ctx);
  if (isError(s)) throw s;
  return s;
}

/** İsteğe bağlı sayısal bağımsız değişken. */
export function optNum(v: Value | undefined, def: number, ctx: FnContext): number {
  if (v === undefined || v === null) return def;
  return num(v, ctx);
}

export function int(v: Value, ctx: FnContext): number {
  return Math.trunc(num(v, ctx));
}

export function text(v: Value, ctx: FnContext): string {
  return check(toText(scalarNoErr(v, ctx), ctx.locale));
}

export function optText(v: Value | undefined, def: string, ctx: FnContext): string {
  if (v === undefined) return def;
  return text(v, ctx);
}

export function bool(v: Value, ctx: FnContext): boolean {
  return check(toBool(scalarNoErr(v, ctx)));
}

export function optBool(v: Value | undefined, def: boolean, ctx: FnContext): boolean {
  if (v === undefined || v === null) return def;
  return bool(v, ctx);
}

/** Bir değeri sırasıyla gezer; fromRef: öğe başvuru/diziden mi geldi. */
export function forEachValue(v: Value, ctx: FnContext, cb: (s: Scalar, fromRef: boolean) => void, clip = true): void {
  if (isRef(v)) {
    const c = clip ? clipRef(v, ctx) : v;
    if (!c) return;
    const host = ctx.host;
    for (let i = c.r1; i <= c.r2; i++) for (let j = c.c1; j <= c.c2; j++) cb(host.getCellValue(c.sheet, i, j), true);
    return;
  }
  if (isMultiRef(v)) {
    v.areas.forEach((a) => forEachValue(a, ctx, cb, clip));
    return;
  }
  if (isArray(v)) {
    for (const row of v.data) for (const s of row) cb(s, true);
    return;
  }
  if (isLambda(v)) throw Errors.VALUE;
  cb(v as Scalar, false);
}

export interface CollectOptions {
  /** Doğrudan verilen metinler sayıya çevrilir (TOPLA("3")). */
  coerceDirectText?: boolean;
  /** Başvurudaki metinleri 0, mantıksalları 1/0 say (ORTALAMAA). */
  refTextAsZero?: boolean;
  /** Gizli satırları atla (ALTTOPLAM 1xx). */
  skipHidden?: boolean;
  /** İç içe ALTTOPLAM hücrelerini atla. */
  skipSubtotals?: boolean;
  /** Hataları yok say (TOPLAMA). */
  ignoreErrors?: boolean;
}

/** TOPLA/ORTALAMA tarzı sayı toplama. */
export function collectNumbers(args: Value[], ctx: FnContext, opts: CollectOptions = {}): number[] {
  const out: number[] = [];
  for (const a of args) {
    if ((opts.skipHidden || opts.skipSubtotals) && isRef(a)) {
      const c = clipRef(a, ctx);
      if (!c) continue;
      for (let i = c.r1; i <= c.r2; i++) {
        if (opts.skipHidden && ctx.host.isRowHidden(c.sheet, i)) continue;
        for (let j = c.c1; j <= c.c2; j++) {
          if (opts.skipSubtotals && ctx.host.isSubtotalCell(c.sheet, i, j)) continue;
          addRefValue(ctx.host.getCellValue(c.sheet, i, j), out, opts);
        }
      }
      continue;
    }
    forEachValue(a, ctx, (s, fromRef) => {
      if (fromRef) addRefValue(s, out, opts);
      else {
        if (s === null) {
          out.push(0);
          return;
        }
        if (isError(s)) {
          if (opts.ignoreErrors) return;
          throw s;
        }
        if (typeof s === "string" && opts.coerceDirectText === false) return;
        out.push(check(toNumber(s, ctx.locale)));
      }
    });
  }
  return out;
}

function addRefValue(s: Scalar, out: number[], opts: CollectOptions) {
  if (typeof s === "number") out.push(s);
  else if (isError(s)) {
    if (!opts.ignoreErrors) throw s;
  } else if (opts.refTextAsZero) {
    if (typeof s === "string") out.push(0);
    else if (typeof s === "boolean") out.push(s ? 1 : 0);
  }
}

/** Tüm değerleri düz liste olarak verir (sayım fonksiyonları için). */
export function flatValues(args: Value[], ctx: FnContext, clip = true): Scalar[] {
  const out: Scalar[] = [];
  for (const a of args) forEachValue(a, ctx, (s) => out.push(s), clip);
  return out;
}

/** Ölçüt aralığı + ölçüt çiftleri için eşleşen konumlar (ÇOKEĞERSAY vb.). */
export function matchCriteriaRanges(pairs: [Value, Value][], ctx: FnContext): { ref: RefValue | null; rows: number; cols: number; mask: boolean[][] } {
  if (pairs.length === 0) throw Errors.VALUE;
  let rows = -1;
  let cols = -1;
  let mask: boolean[][] = [];
  let firstRef: RefValue | null = null;
  pairs.forEach(([range, critV], idx) => {
    const crit = makeCriteria(scalar(critV, ctx), ctx.locale);
    let data: Scalar[][];
    if (isRef(range)) {
      if (idx === 0) firstRef = range;
      // İlk aralığın kırpılmış şekline göre hizala
      const base = idx === 0 ? clipRef(range, ctx) : null;
      if (idx === 0) {
        data = base ? refValues(base, ctx) : [];
        rows = data.length;
        cols = rows ? data[0].length : 0;
      } else {
        const h = range.r2 - range.r1 + 1;
        const w = range.c2 - range.c1 + 1;
        const fr = firstRef!;
        if (h !== fr.r2 - fr.r1 + 1 || w !== fr.c2 - fr.c1 + 1) throw Errors.VALUE;
        data = [];
        for (let i = 0; i < rows; i++) {
          const row: Scalar[] = [];
          for (let j = 0; j < cols; j++) row.push(ctx.host.getCellValue(range.sheet, range.r1 + i, range.c1 + j));
          data.push(row);
        }
      }
    } else {
      data = matrix(range, ctx);
      if (idx === 0) {
        rows = data.length;
        cols = rows ? data[0].length : 0;
      } else if (data.length !== rows || (data[0]?.length ?? 0) !== cols) throw Errors.VALUE;
    }
    if (idx === 0) mask = data.map((r) => r.map((s) => crit(s)));
    else for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) mask[i][j] = mask[i][j] && crit(data[i][j]);
  });
  return { ref: firstRef, rows, cols, mask };
}

/** Aralığın i,j konumundaki değer (aralık dışında boş). */
export function valueAt(v: Value, i: number, j: number, ctx: FnContext): Scalar {
  if (isRef(v)) {
    const r = v.r1 + i;
    const c = v.c1 + j;
    return ctx.host.getCellValue(v.sheet, r, c);
  }
  if (isArray(v)) return v.data[i]?.[j] ?? null;
  return i === 0 && j === 0 ? (v as Scalar) : null;
}

export function arr(data: Scalar[][]): ArrayValue | Scalar {
  if (data.length === 1 && data[0].length === 1) return data[0][0];
  return makeArray(data);
}

/** Sütun vektörü olarak dizi. */
export function column(values: Scalar[]): ArrayValue | Scalar {
  return arr(values.map((v) => [v]));
}

export function dims(v: Value): { rows: number; cols: number } {
  if (isRef(v)) return { rows: v.r2 - v.r1 + 1, cols: v.c2 - v.c1 + 1 };
  if (isArray(v)) return { rows: v.data.length, cols: v.data[0]?.length ?? 0 };
  return { rows: 1, cols: 1 };
}

/** Sonlu sayı değilse #NUM!. */
export function finite(n: number): number {
  if (!Number.isFinite(n)) throw Errors.NUM;
  return n;
}

export function flatNumbersFromMatrix(m: Scalar[][]): number[] {
  const out: number[] = [];
  for (const r of m) for (const s of r) {
    if (isError(s)) throw s;
    if (typeof s === "number") out.push(s);
  }
  return out;
}
