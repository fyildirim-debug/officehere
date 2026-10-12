/** Matematik ve trigonometri fonksiyonları. */
import { toNumber } from "../coerce";
import type { FnContext } from "../context";
import { Errors, isError, isRef, makeArray, type Scalar, type Value } from "../values";
import * as agg from "./aggregate";
import { check, clipRef, collectNumbers, finite, forEachValue, int, matchCriteriaRanges, matrix, num, optNum, scalar, text, valueAt } from "./helpers";
import { fn, type FnDef } from "./registry";
import { makeCriteria } from "../coerce";

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

function roundDir(x: number, digits: number, up: boolean): number {
  const d = Math.trunc(digits);
  const f = Math.pow(10, Math.abs(d));
  const ax = Math.abs(x);
  const scaled = Number((d >= 0 ? ax * f : ax / f).toPrecision(15));
  const r = up ? Math.ceil(scaled) : Math.floor(scaled);
  const back = d >= 0 ? r / f : r * f;
  return x < 0 ? -back : back;
}

function quot(x: number, s: number): number {
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

function fixZero(n: number): number {
  return Object.is(n, -0) ? 0 : Number(n.toPrecision(15));
}

function factorial(n: number): number {
  if (n < 0) throw Errors.NUM;
  const k = Math.floor(n);
  if (k > 170) throw Errors.NUM;
  let r = 1;
  for (let i = 2; i <= k; i++) r *= i;
  return r;
}

function gcd2(a: number, b: number): number {
  a = Math.abs(a);
  b = Math.abs(b);
  while (b) [a, b] = [b, a % b];
  return a;
}

function combin(n: number, k: number): number {
  n = Math.floor(n);
  k = Math.floor(k);
  if (n < 0 || k < 0 || k > n) throw Errors.NUM;
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
}

const ROMAN: [number, string][] = [
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

function toRoman(n: number): string {
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

function fromRoman(s: string): number {
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
function subtotalValues(args: Value[], ctx: FnContext, skipManualHidden: boolean, ignoreErrors = false): { nums: number[]; countA: number } {
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

function isFiltered(ctx: FnContext, sheet: number, row: number): boolean {
  const h = ctx.host as { isRowFiltered?: (s: number, r: number) => boolean };
  return h.isRowFiltered ? h.isRowFiltered(sheet, row) : true;
}

function sumIfImpl(args: Value[], ctx: FnContext): number {
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

function sumProduct(args: Value[], ctx: FnContext): number {
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

function numMatrix(v: Value, ctx: FnContext): number[][] {
  return matrix(v, ctx).map((r) =>
    r.map((x) => {
      if (isError(x)) throw x;
      if (typeof x !== "number") throw Errors.VALUE;
      return x;
    }),
  );
}

function determinant(m: number[][]): number {
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

function inverse(m: number[][]): number[][] {
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

const n1 = (f: (x: number) => number) => (a: Value[], c: FnContext) => finite(f(num(a[0], c)));

export const mathFunctions: FnDef[] = [
  fn("SUM", "TOPLA", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerdeki tüm sayıları toplar.", 1, 255, (a, c) => agg.sum(collectNumbers(a, c))),
  fn("SUMIF", "ETOPLA", "math", "aralık; ölçüt; [toplam_aralığı]", "Verilen koşula uyan hücreleri toplar.", 2, 3, (a, c) => sumIfImpl(a, c)),
  fn(
    "SUMIFS",
    "ÇOKETOPLA",
    "math",
    "toplam_aralığı; ölçüt_aralığı1; ölçüt1; ...",
    "Birden çok koşula uyan hücreleri toplar.",
    3,
    255,
    (a, c) => {
      if ((a.length - 1) % 2) throw Errors.VALUE;
      const pairs: [Value, Value][] = [];
      for (let i = 1; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
      const { mask, rows, cols } = matchCriteriaRanges(pairs, c);
      let total = 0;
      for (let i = 0; i < rows; i++)
        for (let j = 0; j < cols; j++) {
          if (!mask[i][j]) continue;
          const v = valueAt(a[0], i, j, c);
          if (typeof v === "number") total += v;
          else if (isError(v)) throw v;
        }
      return total;
    },
    { xlfn: false },
  ),
  fn("SUMPRODUCT", "TOPLA.ÇARPIM", "math", "dizi1; [dizi2]; ...", "Karşılık gelen dizi öğelerinin çarpımlarının toplamını verir.", 1, 255, sumProduct),
  fn("SUMSQ", "TOPKARE", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerin karelerinin toplamını verir.", 1, 255, (a, c) => agg.sum(collectNumbers(a, c).map((x) => x * x))),
  fn("SUMX2MY2", "TOPX2EY2", "math", "dizi_x; dizi_y", "İki dizideki karşılık gelen değerlerin kareleri farkının toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => x * x - y * y)),
  fn("SUMX2PY2", "TOPX2AY2", "math", "dizi_x; dizi_y", "İki dizideki karşılık gelen değerlerin kareleri toplamının toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => x * x + y * y)),
  fn("SUMXMY2", "TOPXEY2", "math", "dizi_x; dizi_y", "İki dizideki değerlerin farklarının karelerinin toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => (x - y) ** 2)),
  fn("PRODUCT", "ÇARPIM", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerin çarpımını verir.", 1, 255, (a, c) => agg.product(collectNumbers(a, c))),
  fn("ROUND", "YUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı belirtilen basamak sayısına yuvarlar.", 2, 2, (a, c) => roundHalfAway(num(a[0], c), num(a[1], c)), { lift: true }),
  fn("ROUNDUP", "YUKARIYUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı sıfırdan uzağa yuvarlar.", 2, 2, (a, c) => roundDir(num(a[0], c), num(a[1], c), true), { lift: true }),
  fn("ROUNDDOWN", "AŞAĞIYUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı sıfıra doğru yuvarlar.", 2, 2, (a, c) => roundDir(num(a[0], c), num(a[1], c), false), { lift: true }),
  fn("INT", "TAMSAYI", "math", "sayı", "Sayıyı en yakın alt tamsayıya yuvarlar.", 1, 1, n1(Math.floor), { lift: true }),
  fn("TRUNC", "NSAT", "math", "sayı; [sayı_rakamlar]", "Sayının ondalık kısmını atar.", 1, 2, (a, c) => roundDir(num(a[0], c), optNum(a[1], 0, c), false), { lift: true }),
  fn(
    "MOD",
    "MOD",
    "math",
    "sayı; bölen",
    "Bölme işleminden kalanı verir.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const d = num(a[1], c);
      if (d === 0) throw Errors.DIV0;
      return fixZero(x - d * Math.floor(x / d));
    },
    { lift: true },
  ),
  fn("ABS", "MUTLAK", "math", "sayı", "Sayının mutlak değerini verir.", 1, 1, n1(Math.abs), { lift: true }),
  fn(
    "POWER",
    "KUVVET",
    "math",
    "sayı; üs",
    "Sayının üssünü verir.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const y = num(a[1], c);
      if (x === 0 && y === 0) throw Errors.NUM;
      if (x === 0 && y < 0) throw Errors.DIV0;
      return finite(Math.pow(x, y));
    },
    { lift: true },
  ),
  fn(
    "SQRT",
    "KAREKÖK",
    "math",
    "sayı",
    "Sayının karekökünü verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x < 0) throw Errors.NUM;
      return Math.sqrt(x);
    },
    { lift: true },
  ),
  fn("SQRTPI", "KAREKÖKPİ", "math", "sayı", "(sayı * Pi) değerinin karekökünü verir.", 1, 1, (a, c) => {
    const x = num(a[0], c);
    if (x < 0) throw Errors.NUM;
    return Math.sqrt(x * Math.PI);
  }, { lift: true }),
  fn("EXP", "ÜS", "math", "sayı", "e sayısının verilen kuvvetini verir.", 1, 1, n1(Math.exp), { lift: true }),
  fn(
    "LN",
    "LN",
    "math",
    "sayı",
    "Sayının doğal logaritmasını verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x <= 0) throw Errors.NUM;
      return Math.log(x);
    },
    { lift: true },
  ),
  fn(
    "LOG",
    "LOG",
    "math",
    "sayı; [taban]",
    "Sayının belirtilen tabandaki logaritmasını verir.",
    1,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const b = optNum(a[1], 10, c);
      if (x <= 0 || b <= 0) throw Errors.NUM;
      if (b === 1) throw Errors.DIV0;
      return Math.log(x) / Math.log(b);
    },
    { lift: true },
  ),
  fn(
    "LOG10",
    "LOG10",
    "math",
    "sayı",
    "Sayının 10 tabanında logaritmasını verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x <= 0) throw Errors.NUM;
      return Math.log10(x);
    },
    { lift: true },
  ),
  fn("PI", "Pİ", "math", "", "Pi sayısını (3,14159265358979) verir.", 0, 0, () => Math.PI),
  fn("RAND", "S_SAYI_ÜRET", "math", "", "0 ile 1 arasında rastgele bir sayı verir.", 0, 0, (_a, c) => c.host.random(), { volatile: true }),
  fn(
    "RANDBETWEEN",
    "RASTGELEARADA",
    "math",
    "alt; üst",
    "Belirtilen sayılar arasında rastgele bir tamsayı verir.",
    2,
    2,
    (a, c) => {
      const lo = Math.ceil(num(a[0], c));
      const hi = Math.floor(num(a[1], c));
      if (lo > hi) throw Errors.NUM;
      return lo + Math.floor(c.host.random() * (hi - lo + 1));
    },
    { volatile: true },
  ),
  fn(
    "RANDARRAY",
    "RASTGDİZİ",
    "math",
    "[satırlar]; [sütunlar]; [min]; [maks]; [tamsayı]",
    "Rastgele sayılardan oluşan bir dizi verir.",
    0,
    5,
    (a, c) => {
      const rows = Math.trunc(optNum(a[0], 1, c));
      const cols = Math.trunc(optNum(a[1], 1, c));
      const lo = optNum(a[2], 0, c);
      const hi = optNum(a[3], 1, c);
      const whole = a[4] !== undefined && a[4] !== null ? !!num(a[4], c) : false;
      if (rows < 1 || cols < 1 || lo > hi) throw Errors.VALUE;
      const data: Scalar[][] = [];
      for (let i = 0; i < rows; i++) {
        const r: Scalar[] = [];
        for (let j = 0; j < cols; j++) {
          const x = c.host.random();
          r.push(whole ? Math.ceil(lo) + Math.floor(x * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + x * (hi - lo));
        }
        data.push(r);
      }
      return makeArray(data);
    },
    { volatile: true, xlfn: true },
  ),
  fn("CEILING", "TAVANAYUVARLA", "math", "sayı; anlam", "Sayıyı anlamın en yakın katına yukarı yuvarlar.", 2, 2, (a, c) => fixZero(ceilingExcel(num(a[0], c), num(a[1], c))), { lift: true }),
  fn(
    "CEILING.MATH",
    "TAVANAYUVARLA.MATEMATİK",
    "math",
    "sayı; [anlam]; [mod]",
    "Sayıyı en yakın tamsayıya ya da anlamın en yakın katına yukarı yuvarlar.",
    1,
    3,
    (a, c) => {
      const x = num(a[0], c);
      const s = Math.abs(optNum(a[1], 1, c));
      const mode = optNum(a[2], 0, c);
      if (s === 0) return 0;
      if (x >= 0 || mode === 0) return fixZero(Math.ceil(quot(x, s)) * s);
      return fixZero(-Math.ceil(quot(-x, s)) * s);
    },
    { lift: true, xlfn: true },
  ),
  fn("CEILING.PRECISE", "TAVANAYUVARLA.DUYARLI", "math", "sayı; [anlam]", "Sayıyı anlamın en yakın katına yukarı yuvarlar.", 1, 2, (a, c) => {
    const x = num(a[0], c);
    const s = Math.abs(optNum(a[1], 1, c));
    return s === 0 ? 0 : fixZero(Math.ceil(quot(x, s)) * s);
  }, { lift: true, xlfn: true }),
  fn("FLOOR", "TABANAYUVARLA", "math", "sayı; anlam", "Sayıyı anlamın en yakın katına aşağı yuvarlar.", 2, 2, (a, c) => fixZero(floorExcel(num(a[0], c), num(a[1], c))), { lift: true }),
  fn(
    "FLOOR.MATH",
    "TABANAYUVARLA.MATEMATİK",
    "math",
    "sayı; [anlam]; [mod]",
    "Sayıyı en yakın tamsayıya ya da anlamın en yakın katına aşağı yuvarlar.",
    1,
    3,
    (a, c) => {
      const x = num(a[0], c);
      const s = Math.abs(optNum(a[1], 1, c));
      const mode = optNum(a[2], 0, c);
      if (s === 0) return 0;
      if (x >= 0 || mode === 0) return fixZero(Math.floor(quot(x, s)) * s);
      return fixZero(-Math.floor(quot(-x, s)) * s);
    },
    { lift: true, xlfn: true },
  ),
  fn("FLOOR.PRECISE", "TABANAYUVARLA.DUYARLI", "math", "sayı; [anlam]", "Sayıyı anlamın en yakın katına aşağı yuvarlar.", 1, 2, (a, c) => {
    const x = num(a[0], c);
    const s = Math.abs(optNum(a[1], 1, c));
    return s === 0 ? 0 : fixZero(Math.floor(quot(x, s)) * s);
  }, { lift: true, xlfn: true }),
  fn("SIGN", "İŞARET", "math", "sayı", "Sayının işaretini verir.", 1, 1, n1((x) => (x > 0 ? 1 : x < 0 ? -1 : 0)), { lift: true }),
  fn("FACT", "ÇARPINIM", "math", "sayı", "Sayının faktöriyelini verir.", 1, 1, (a, c) => factorial(num(a[0], c)), { lift: true }),
  fn("FACTDOUBLE", "ÇİFTFAKTÖR", "math", "sayı", "Sayının çift faktöriyelini verir.", 1, 1, (a, c) => {
    const n = Math.floor(num(a[0], c));
    if (n < -1) throw Errors.NUM;
    let r = 1;
    for (let i = n; i > 1; i -= 2) r *= i;
    return r;
  }, { lift: true }),
  fn("GCD", "OBEB", "math", "sayı1; [sayı2]; ...", "En büyük ortak böleni verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    return xs.reduce((g, x) => gcd2(g, x), 0);
  }),
  fn("LCM", "OKEK", "math", "sayı1; [sayı2]; ...", "En küçük ortak katı verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    if (xs.some((x) => x === 0)) return 0;
    return xs.reduce((l, x) => (l * x) / gcd2(l, x), 1);
  }),
  fn("EVEN", "ÇİFT", "math", "sayı", "Sayıyı en yakın çift tamsayıya yukarı yuvarlar.", 1, 1, n1((x) => {
    const r = Math.ceil(Math.abs(x) / 2) * 2;
    return x < 0 ? -r : r;
  }), { lift: true }),
  fn("ODD", "TEK", "math", "sayı", "Sayıyı en yakın tek tamsayıya yukarı yuvarlar.", 1, 1, n1((x) => {
    let r = Math.ceil(Math.abs(x));
    if (r % 2 === 0) r += 1;
    return x < 0 ? -r : r;
  }), { lift: true }),
  fn(
    "MROUND",
    "KYUVARLA",
    "math",
    "sayı; katsayı",
    "Sayıyı istenen katsayıya yuvarlar.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const m = num(a[1], c);
      if (m === 0) return 0;
      if (x * m < 0) throw Errors.NUM;
      return fixZero(roundHalfAway(x / m, 0) * m);
    },
    { lift: true },
  ),
  fn(
    "QUOTIENT",
    "BÖLÜM",
    "math",
    "pay; payda",
    "Bölmenin tamsayı kısmını verir.",
    2,
    2,
    (a, c) => {
      const d = num(a[1], c);
      if (d === 0) throw Errors.DIV0;
      return Math.trunc(num(a[0], c) / d);
    },
    { lift: true },
  ),
  fn("SUBTOTAL", "ALTTOPLAM", "math", "işlev_sayısı; başv1; ...", "Bir listedeki ya da veritabanındaki alt toplamı verir.", 2, 255, (a, c) => {
    const code = int(a[0], c);
    const skipHidden = code > 100;
    const k = skipHidden ? code - 100 : code;
    if (k < 1 || k > 11) throw Errors.VALUE;
    const { nums, countA } = subtotalValues(a.slice(1), c, skipHidden);
    return agg.subtotalBy(k, nums, countA);
  }),
  fn("AGGREGATE", "TOPLAMA", "math", "işlev_sayısı; seçenekler; başv1; ...", "Bir listedeki toplamı, gizli satırları ve hataları yok sayma seçenekleriyle verir.", 3, 255, (a, c) => {
    const code = int(a[0], c);
    const opt = int(a[1], c);
    const ignoreErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
    const skipHidden = opt === 1 || opt === 3 || opt === 5 || opt === 7;
    if (code >= 14 && code <= 19) {
      const { nums } = subtotalValues([a[2]], c, skipHidden, ignoreErrors);
      const k = num(a[3], c);
      switch (code) {
        case 14:
          return agg.kth(nums, k, true);
        case 15:
          return agg.kth(nums, k, false);
        case 16:
          return agg.percentileInc(nums, k);
        case 17:
          return agg.percentileInc(nums, k / 4);
        case 18:
          return agg.percentileExc(nums, k);
        default:
          return agg.percentileExc(nums, k / 4);
      }
    }
    const { nums, countA } = subtotalValues(a.slice(2), c, skipHidden, ignoreErrors);
    return agg.subtotalBy(code, nums, countA);
  }, { xlfn: true }),
  fn("RADIANS", "RADYAN", "math", "açı", "Dereceyi radyana çevirir.", 1, 1, n1((x) => (x * Math.PI) / 180), { lift: true }),
  fn("DEGREES", "DERECE", "math", "açı", "Radyanı dereceye çevirir.", 1, 1, n1((x) => (x * 180) / Math.PI), { lift: true }),
  fn("SIN", "SİN", "math", "sayı", "Açının sinüsünü verir.", 1, 1, n1(Math.sin), { lift: true }),
  fn("COS", "COS", "math", "sayı", "Açının kosinüsünü verir.", 1, 1, n1(Math.cos), { lift: true }),
  fn("TAN", "TAN", "math", "sayı", "Açının tanjantını verir.", 1, 1, n1(Math.tan), { lift: true }),
  fn("ASIN", "ASİN", "math", "sayı", "Sayının arksinüsünü verir.", 1, 1, n1((x) => {
    if (x < -1 || x > 1) throw Errors.NUM;
    return Math.asin(x);
  }), { lift: true }),
  fn("ACOS", "ACOS", "math", "sayı", "Sayının arkkosinüsünü verir.", 1, 1, n1((x) => {
    if (x < -1 || x > 1) throw Errors.NUM;
    return Math.acos(x);
  }), { lift: true }),
  fn("ATAN", "ATAN", "math", "sayı", "Sayının arktanjantını verir.", 1, 1, n1(Math.atan), { lift: true }),
  fn("ATAN2", "ATAN2", "math", "x_sayısı; y_sayısı", "x ve y koordinatlarının arktanjantını verir.", 2, 2, (a, c) => {
    const x = num(a[0], c);
    const y = num(a[1], c);
    if (x === 0 && y === 0) throw Errors.DIV0;
    return Math.atan2(y, x);
  }, { lift: true }),
  fn("SINH", "SİNH", "math", "sayı", "Hiperbolik sinüsü verir.", 1, 1, n1(Math.sinh), { lift: true }),
  fn("COSH", "COSH", "math", "sayı", "Hiperbolik kosinüsü verir.", 1, 1, n1(Math.cosh), { lift: true }),
  fn("TANH", "TANH", "math", "sayı", "Hiperbolik tanjantı verir.", 1, 1, n1(Math.tanh), { lift: true }),
  fn("ASINH", "ASİNH", "math", "sayı", "Ters hiperbolik sinüsü verir.", 1, 1, n1(Math.asinh), { lift: true }),
  fn("ACOSH", "ACOSH", "math", "sayı", "Ters hiperbolik kosinüsü verir.", 1, 1, n1((x) => {
    if (x < 1) throw Errors.NUM;
    return Math.acosh(x);
  }), { lift: true }),
  fn("ATANH", "ATANH", "math", "sayı", "Ters hiperbolik tanjantı verir.", 1, 1, n1((x) => {
    if (x <= -1 || x >= 1) throw Errors.NUM;
    return Math.atanh(x);
  }), { lift: true }),
  fn("COT", "COT", "math", "sayı", "Açının kotanjantını verir.", 1, 1, n1((x) => {
    if (x === 0) throw Errors.DIV0;
    return 1 / Math.tan(x);
  }), { lift: true, xlfn: true }),
  fn("CSC", "CSC", "math", "sayı", "Açının kosekantını verir.", 1, 1, n1((x) => {
    if (x === 0) throw Errors.DIV0;
    return 1 / Math.sin(x);
  }), { lift: true, xlfn: true }),
  fn("SEC", "SEC", "math", "sayı", "Açının sekantını verir.", 1, 1, n1((x) => 1 / Math.cos(x)), { lift: true, xlfn: true }),
  fn("COMBIN", "KOMBİNASYON", "math", "sayı; sayı_seçilen", "Verilen öğe sayısı için kombinasyon sayısını verir.", 2, 2, (a, c) => combin(num(a[0], c), num(a[1], c)), { lift: true }),
  fn("COMBINA", "KOMBİNASYONA", "math", "sayı; sayı_seçilen", "Yinelemeli kombinasyon sayısını verir.", 2, 2, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const k = Math.floor(num(a[1], c));
    if (n < 0 || k < 0) throw Errors.NUM;
    return n === 0 && k === 0 ? 1 : combin(n + k - 1, k);
  }, { lift: true, xlfn: true }),
  fn("PERMUT", "PERMÜTASYON", "math", "sayı; sayı_seçilen", "Permütasyon sayısını verir.", 2, 2, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const k = Math.floor(num(a[1], c));
    if (n < 0 || k < 0 || k > n) throw Errors.NUM;
    let r = 1;
    for (let i = 0; i < k; i++) r *= n - i;
    return r;
  }, { lift: true }),
  fn("MULTINOMIAL", "ÇOKTERİMLİ", "math", "sayı1; [sayı2]; ...", "Sayılar kümesinin çok terimlisini verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    return Math.round(factorial(agg.sum(xs)) / xs.reduce((p, x) => p * factorial(x), 1));
  }),
  fn("ROMAN", "ROMEN", "math", "sayı; [biçim]", "Arap rakamını Romen rakamına çevirir.", 1, 2, (a, c) => toRoman(num(a[0], c)), { lift: true }),
  fn("ARABIC", "ARAP", "math", "metin", "Romen rakamını Arap rakamına çevirir.", 1, 1, (a, c) => fromRoman(text(a[0], c)), { lift: true, xlfn: true }),
  fn("BASE", "TABAN", "math", "sayı; taban; [min_uzunluk]", "Sayıyı verilen tabanda metne çevirir.", 2, 3, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const b = Math.floor(num(a[1], c));
    if (n < 0 || b < 2 || b > 36) throw Errors.NUM;
    return n.toString(b).toUpperCase().padStart(Math.floor(optNum(a[2], 0, c)), "0");
  }, { lift: true, xlfn: true }),
  fn("DECIMAL", "ONDALIK", "math", "sayı; taban", "Verilen tabandaki metni ondalık sayıya çevirir.", 2, 2, (a, c) => {
    const s = text(a[0], c);
    const b = Math.floor(num(a[1], c));
    if (b < 2 || b > 36) throw Errors.NUM;
    const n = parseInt(s, b);
    if (Number.isNaN(n)) throw Errors.NUM;
    return n;
  }, { lift: true, xlfn: true }),
  fn("MMULT", "DÇARP", "math", "dizi1; dizi2", "İki dizinin matris çarpımını verir.", 2, 2, (a, c) => {
    const A = numMatrix(a[0], c);
    const B = numMatrix(a[1], c);
    if (A[0].length !== B.length) throw Errors.VALUE;
    return makeArray(A.map((row) => B[0].map((_, j) => row.reduce((s, x, k) => s + x * B[k][j], 0))));
  }),
  fn("MDETERM", "DETERMİNANT", "math", "dizi", "Dizinin matris determinantını verir.", 1, 1, (a, c) => {
    const m = numMatrix(a[0], c);
    if (m.length !== m[0].length) throw Errors.VALUE;
    return determinant(m);
  }),
  fn("MINVERSE", "DİZEY_TERS", "math", "dizi", "Dizinin ters matrisini verir.", 1, 1, (a, c) => {
    const m = numMatrix(a[0], c);
    if (m.length !== m[0].length) throw Errors.VALUE;
    return makeArray(inverse(m));
  }),
  fn("MUNIT", "BİRİMMATRİS", "math", "boyut", "Belirtilen boyutta birim matris verir.", 1, 1, (a, c) => {
    const n = int(a[0], c);
    if (n < 1) throw Errors.VALUE;
    return makeArray(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
  }, { xlfn: true }),
];

function pairSum(a: Value[], c: FnContext, f: (x: number, y: number) => number): number {
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

void check;
void toNumber;
