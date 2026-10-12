/** Matematik fonksiyonlarının yardımcıları: Excel yuvarlama, faktöriyel, Roma rakamları, matris işlemleri, ALTTOPLAM/ETOPLA çekirdekleri. */
import type { FnContext } from "../context";
import { makeCriteria } from "../coerce";
import { Errors, isError, isRef, type Value } from "../values";
import { clipRef, forEachValue, matrix, scalar, valueAt } from "./helpers";
/** Excel YUVARLA: yarıyı sıfırdan uzağa yuvarlar, kayan nokta hatasını düzeltir. */
export function roundHalfAway(x: number, digits: number): number {
  if (!Number.isFinite(x)) return x;
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  if (d >= 0) {
    const f = Math.pow(10, d);
    const v = Number((ax * f).toPrecision(15));
    return (sign * Math.round(v)) / f;
  }
  const f = Math.pow(10, -d);
  const v = Number((ax / f).toPrecision(15));
  return sign * Math.round(v) * f;
}

export function roundDir(x: number, digits: number, up: boolean): number {
  const d = Math.trunc(digits);
  const f = Math.pow(10, Math.abs(d));
  const ax = Math.abs(x);
  const scaled = Number((d >= 0 ? ax * f : ax / f).toPrecision(15));
  const r = up ? Math.ceil(scaled) : Math.floor(scaled);
  const back = d >= 0 ? r / f : r * f;
  return x < 0 ? -back : back;
}

export function quot(x: number, s: number): number {
  const q = x / s;
  const r = Math.round(q);
  return Math.abs(q - r) < 1e-12 * Math.max(1, Math.abs(q)) ? r : q;
}

export function ceilingExcel(x: number, s: number): number {
  if (s === 0 || x === 0) return 0;
  if (x > 0 && s < 0) throw Errors.NUM;
  if (x < 0 && s < 0) return -Math.ceil(quot(-x, -s)) * -s;
  return Math.ceil(quot(x, s)) * s;
}

export function floorExcel(x: number, s: number): number {
  if (x === 0) return 0;
  if (s === 0) throw Errors.DIV0;
  if (x > 0 && s < 0) throw Errors.NUM;
  if (x < 0 && s < 0) return -Math.floor(quot(-x, -s)) * -s;
  return Math.floor(quot(x, s)) * s;
}

export function fixZero(n: number): number {
  return Object.is(n, -0) ? 0 : Number(n.toPrecision(15));
}

export function factorial(n: number): number {
  if (n < 0) throw Errors.NUM;
  const k = Math.floor(n);
  if (k > 170) throw Errors.NUM;
  let r = 1;
  for (let i = 2; i <= k; i++) r *= i;
  return r;
}

export function gcd2(a: number, b: number): number {
  a = Math.abs(a);
  b = Math.abs(b);
  while (b) [a, b] = [b, a % b];
  return a;
}

export function combin(n: number, k: number): number {
  n = Math.floor(n);
  k = Math.floor(k);
  if (n < 0 || k < 0 || k > n) throw Errors.NUM;
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
}

export const ROMAN: [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

export function toRoman(n: number): string {
  if (n < 0 || n > 3999) throw Errors.VALUE;
  let s = "";
  let v = Math.floor(n);
  for (const [k, r] of ROMAN) {
    while (v >= k) {
      s += r;
      v -= k;
    }
  }
  return s;
}

export function fromRoman(s: string): number {
  const t = s.trim().toUpperCase();
  if (t === "") return 0;
  const neg = t.startsWith("-");
  const body = neg ? t.slice(1) : t;
  const map: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  for (let i = 0; i < body.length; i++) {
    const v = map[body[i]];
    if (v === undefined) throw Errors.VALUE;
    const next = map[body[i + 1]] ?? 0;
    total += v < next ? -v : v;
  }
  return neg ? -total : total;
}

/** SUBTOTAL/AGGREGATE için başvurulardaki değerler. */
export function subtotalValues(args: Value[], ctx: FnContext, skipManualHidden: boolean, ignoreErrors = false): { nums: number[]; countA: number } {
  const nums: number[] = [];
  let countA = 0;
  for (const a of args) {
    if (isRef(a)) {
      const c = clipRef(a, ctx);
      if (!c) continue;
      for (let i = c.r1; i <= c.r2; i++) {
        if (ctx.host.isRowHidden(c.sheet, i) && (skipManualHidden || isFiltered(ctx, c.sheet, i))) continue;
        for (let j = c.c1; j <= c.c2; j++) {
          if (ctx.host.isSubtotalCell(c.sheet, i, j)) continue;
          const v = ctx.host.getCellValue(c.sheet, i, j);
          if (v !== null) countA++;
          if (typeof v === "number") nums.push(v);
          else if (isError(v) && !ignoreErrors) throw v;
        }
      }
    } else {
      forEachValue(a, ctx, (v) => {
        if (v !== null) countA++;
        if (typeof v === "number") nums.push(v);
        else if (isError(v) && !ignoreErrors) throw v;
      });
    }
  }
  return { nums, countA };
}

export function isFiltered(ctx: FnContext, sheet: number, row: number): boolean {
  const h = ctx.host as { isRowFiltered?: (s: number, r: number) => boolean };
  return h.isRowFiltered ? h.isRowFiltered(sheet, row) : true;
}

export function sumIfImpl(args: Value[], ctx: FnContext): number {
  const [range, critV, sumRange] = args;
  const crit = makeCriteria(scalar(critV, ctx), ctx.locale);
  let total = 0;
  if (isRef(range)) {
    const c = clipRef(range, ctx);
    if (!c) return 0;
    const sr = sumRange !== undefined ? sumRange : range;
    for (let i = c.r1; i <= c.r2; i++) {
      for (let j = c.c1; j <= c.c2; j++) {
        if (!crit(ctx.host.getCellValue(c.sheet, i, j))) continue;
        const v = valueAt(sr, i - range.r1, j - range.c1, ctx);
        if (typeof v === "number") total += v;
        else if (isError(v)) throw v;
      }
    }
    return total;
  }
  const data = matrix(range, ctx);
  data.forEach((row, i) =>
    row.forEach((v, j) => {
      if (!crit(v)) return;
      const s = sumRange !== undefined ? valueAt(sumRange, i, j, ctx) : v;
      if (typeof s === "number") total += s;
      else if (isError(s)) throw s;
    }),
  );
  return total;
}

export function sumProduct(args: Value[], ctx: FnContext): number {
  const mats = args.map((a) => matrix(a, ctx));
  const rows = mats[0].length;
  const cols = mats[0][0]?.length ?? 0;
  for (const m of mats) if (m.length !== rows || (m[0]?.length ?? 0) !== cols) throw Errors.VALUE;
  let total = 0;
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      let p = 1;
      for (const m of mats) {
        const v = m[i][j];
        if (isError(v)) throw v;
        p *= typeof v === "number" ? v : 0;
      }
      total += p;
    }
  }
  return total;
}

export function numMatrix(v: Value, ctx: FnContext): number[][] {
  return matrix(v, ctx).map((r) =>
    r.map((x) => {
      if (isError(x)) throw x;
      if (typeof x !== "number") throw Errors.VALUE;
      return x;
    }),
  );
}

export function determinant(m: number[][]): number {
  const n = m.length;
  const a = m.map((r) => [...r]);
  let det = 1;
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let k = i + 1; k < n; k++) if (Math.abs(a[k][i]) > Math.abs(a[p][i])) p = k;
    if (a[p][i] === 0) return 0;
    if (p !== i) {
      [a[p], a[i]] = [a[i], a[p]];
      det = -det;
    }
    det *= a[i][i];
    for (let k = i + 1; k < n; k++) {
      const f = a[k][i] / a[i][i];
      for (let j = i; j < n; j++) a[k][j] -= f * a[i][j];
    }
  }
  return det;
}

export function inverse(m: number[][]): number[][] {
  const n = m.length;
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let k = i + 1; k < n; k++) if (Math.abs(a[k][i]) > Math.abs(a[p][i])) p = k;
    if (Math.abs(a[p][i]) < 1e-15) throw Errors.NUM;
    [a[p], a[i]] = [a[i], a[p]];
    const piv = a[i][i];
    for (let j = 0; j < 2 * n; j++) a[i][j] /= piv;
    for (let k = 0; k < n; k++) {
      if (k === i) continue;
      const f = a[k][i];
      for (let j = 0; j < 2 * n; j++) a[k][j] -= f * a[i][j];
    }
  }
  return a.map((r) => r.slice(n));
}


export function pairSum(a: Value[], c: FnContext, f: (x: number, y: number) => number): number {
  const X = matrix(a[0], c).flat();
  const Y = matrix(a[1], c).flat();
  if (X.length !== Y.length) throw Errors.NA;
  let s = 0;
  for (let i = 0; i < X.length; i++) {
    const x = X[i];
    const y = Y[i];
    if (isError(x)) throw x;
    if (isError(y)) throw y;
    if (typeof x === "number" && typeof y === "number") s += f(x, y);
  }
  return s;
}
