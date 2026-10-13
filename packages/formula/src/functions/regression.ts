/**
 * Çoklu doğrusal regresyon işlevleri: DOT (LINEST), LOT (LOGEST), EĞİLİM (TREND),
 * BÜYÜME (GROWTH). En küçük kareler çözümü sütun seçmeli Gauss eleme ile yapılır;
 * diğerleriyle doğrusal bağımlı değişkenin katsayısı ve standart hatası 0 olur.
 */
import type { FnContext } from "../context";
import { Errors, makeArray, type Scalar, type Value } from "../values";
import { optBool } from "./helpers";
import { numMatrix } from "./math-helpers";
import { fn, type FnDef } from "./registry";

interface Fit {
  /** Değişken katsayıları (x sırasıyla; atılanlar 0). */
  coef: number[];
  b: number;
  /** Standart hatalar (katsayılar, sonra sabit). */
  se: number[];
  seB: number;
  r2: number;
  sey: number;
  f: number;
  df: number;
  ssreg: number;
  ssresid: number;
  k: number;
}

/** y ve x'leri gözlem satırlarına çevirir: y n öğeli, X n × p. */
function observations(yv: Value, xv: Value | undefined, ctx: FnContext, logY: boolean): { y: number[]; X: number[][]; byRow: boolean; shape: [number, number] } {
  const Y = numMatrix(yv, ctx);
  const rows = Y.length;
  const cols = Y[0]?.length ?? 0;
  if (!rows || !cols) throw Errors.VALUE;
  const byRow = rows === 1 && cols > 1;
  const y = Y.flat();
  if (logY) {
    if (y.some((v) => v <= 0)) throw Errors.NUM;
    for (let i = 0; i < y.length; i++) y[i] = Math.log(y[i]);
  }
  const n = y.length;
  let X: number[][];
  if (xv === undefined || xv === null) X = y.map((_, i) => [i + 1]);
  else {
    const M = numMatrix(xv, ctx);
    if (rows > 1 && cols > 1) {
      // y bir matris: x aynı boyutta tek değişken olmalı
      if (M.length !== rows || M[0].length !== cols) throw Errors.REF;
      X = M.flat().map((v) => [v]);
    } else if (byRow) {
      if (M[0]?.length !== n) throw Errors.REF;
      X = Array.from({ length: n }, (_, i) => M.map((r) => r[i]));
    } else {
      if (M.length !== n) throw Errors.REF;
      X = M.map((r) => r.slice());
    }
  }
  return { y, X, byRow, shape: [rows, cols] };
}

function fit(y: number[], X: number[][], useConst: boolean): Fit {
  const n = y.length;
  const p = X[0].length;
  // Tasarım matrisi: sabit varsa ilk sütun 1
  const A = X.map((r) => (useConst ? [1, ...r] : r.slice()));
  const q = A[0].length;
  // Normal denklemler: (AᵀA) β = Aᵀy
  const N = Array.from({ length: q }, (_, i) => Array.from({ length: q }, (_, j) => A.reduce((s, r) => s + r[i] * r[j], 0)));
  const v = Array.from({ length: q }, (_, i) => A.reduce((s, r, k) => s + r[i] * y[k], 0));
  // Gauss-Jordan; tersini de çıkar (standart hatalar için), bağımlı sütunları at
  const M = N.map((r, i) => [...r, ...Array.from({ length: q }, (_, j) => (i === j ? 1 : 0)), v[i]]);
  const dropped = new Array(q).fill(false);
  const scale = Math.max(1e-300, ...N.map((r, i) => Math.abs(r[i])));
  for (let c = 0; c < q; c++) {
    // Kısmi sıralı pivot: c'den sonraki satırlar arasında
    let piv = c;
    for (let r = c + 1; r < q; r++) if (Math.abs(M[r][c]) > Math.abs(M[piv][c])) piv = r;
    if (Math.abs(M[piv][c]) <= 1e-12 * scale) {
      dropped[c] = true;
      continue;
    }
    [M[c], M[piv]] = [M[piv], M[c]];
    const d = M[c][c];
    for (let j = 0; j < M[c].length; j++) M[c][j] /= d;
    for (let r = 0; r < q; r++) {
      if (r === c || M[r][c] === 0) continue;
      const f = M[r][c];
      for (let j = 0; j < M[r].length; j++) M[r][j] -= f * M[c][j];
    }
  }
  const beta = M.map((r, i) => (dropped[i] ? 0 : r[2 * q]));
  const inv = (i: number) => (dropped[i] ? 0 : M[i][q + i]);
  const yhat = A.map((r) => r.reduce((s, x, j) => s + x * beta[j], 0));
  const ybar = y.reduce((s, x) => s + x, 0) / n;
  let ssresid = 0;
  let ssreg = 0;
  for (let i = 0; i < n; i++) {
    ssresid += (y[i] - yhat[i]) ** 2;
    ssreg += useConst ? (yhat[i] - ybar) ** 2 : yhat[i] ** 2;
  }
  const k = dropped.filter((d, i) => !d && (!useConst || i > 0)).length;
  const df = n - k - (useConst ? 1 : 0);
  const sey2 = df > 0 ? ssresid / df : NaN;
  const se = Array.from({ length: q }, (_, i) => (dropped[i] ? 0 : Math.sqrt(Math.max(0, inv(i) * sey2))));
  const off = useConst ? 1 : 0;
  return {
    coef: beta.slice(off, off + p),
    b: useConst ? beta[0] : 0,
    se: se.slice(off, off + p),
    seB: useConst ? se[0] : NaN,
    r2: ssreg + ssresid === 0 ? 1 : ssreg / (ssreg + ssresid),
    sey: Math.sqrt(sey2),
    f: df > 0 && ssresid > 0 ? ssreg / k / (ssresid / df) : NaN,
    df,
    ssreg,
    ssresid,
    k,
  };
}

function num(x: number): Scalar {
  return Number.isFinite(x) ? x : Errors.NUM;
}

/** DOT/LOT çıktısı: ilk satır katsayılar (ters sırada) ve sabit, istenirse 4 istatistik satırı. */
function estOutput(f: Fit, stats: boolean, exp: boolean): Value {
  const tr = (x: number) => (exp ? Math.exp(x) : x);
  const p = f.coef.length;
  const first: Scalar[] = [...f.coef.map(tr).reverse(), tr(f.b)];
  if (!stats) return p === 0 ? first[0] : makeArray([first]);
  const w = p + 1;
  const pad = (vals: Scalar[]) => [...vals, ...new Array(w - vals.length).fill(Errors.NA)];
  return makeArray([
    first,
    [...f.se.slice().reverse(), Number.isNaN(f.seB) ? Errors.NA : f.seB],
    pad([num(f.r2), num(f.sey)]),
    pad([num(f.f), f.df]),
    pad([f.ssreg, f.ssresid]),
  ]);
}

/** EĞİLİM/BÜYÜME: yeni x'ler için tahminler (yeni x verilmezse bilinen x'ler). */
function predict(a: Value[], ctx: FnContext, exp: boolean): Value {
  const obs = observations(a[0], a[1] ?? undefined, ctx, exp);
  const useConst = optBool(a[3], true, ctx);
  const f = fit(obs.y, obs.X, useConst);
  const p = obs.X[0].length;
  let newX: number[][];
  let outShape: [number, number];
  if (a[2] === undefined || a[2] === null) {
    // Yeni x verilmezse bilinen x'ler; sonuç bilinen y'lerin biçiminde
    newX = obs.X;
    outShape = obs.shape;
  } else {
    const M = numMatrix(a[2], ctx);
    if (p === 1) {
      newX = M.flat().map((v) => [v]);
      outShape = [M.length, M[0].length];
    } else if (obs.byRow) {
      if (M.length !== p) throw Errors.REF;
      newX = Array.from({ length: M[0].length }, (_, i) => M.map((r) => r[i]));
      outShape = [1, newX.length];
    } else {
      if (M[0].length !== p) throw Errors.REF;
      newX = M;
      outShape = [M.length, 1];
    }
  }
  const vals = newX.map((r) => {
    const yv = f.b + r.reduce((s, x, j) => s + x * f.coef[j], 0);
    return exp ? Math.exp(yv) : yv;
  });
  const [rows, cols] = outShape;
  if (rows * cols !== vals.length) throw Errors.REF;
  const out: Scalar[][] = Array.from({ length: rows }, (_, i) => vals.slice(i * cols, i * cols + cols));
  return rows === 1 && cols === 1 ? out[0][0] : makeArray(out);
}

export const regressionFunctions: FnDef[] = [
  fn("LINEST", "DOT", "stats", "bilinen_y'ler; [bilinen_x'ler]; [sabit]; [konum]", "En küçük kareler yöntemiyle doğrusal eğilimi tanımlayan istatistikleri verir.", 1, 4, (a, c) => {
    const obs = observations(a[0], a[1] ?? undefined, c, false);
    return estOutput(fit(obs.y, obs.X, optBool(a[2], true, c)), optBool(a[3], false, c), false);
  }),
  fn("LOGEST", "LOT", "stats", "bilinen_y'ler; [bilinen_x'ler]; [sabit]; [konum]", "Verilere uyan üstel eğrinin istatistiklerini verir.", 1, 4, (a, c) => {
    const obs = observations(a[0], a[1] ?? undefined, c, true);
    return estOutput(fit(obs.y, obs.X, optBool(a[2], true, c)), optBool(a[3], false, c), true);
  }),
  fn("TREND", "EĞİLİM", "stats", "bilinen_y'ler; [bilinen_x'ler]; [yeni_x'ler]; [sabit]", "Doğrusal eğilim boyunca değerleri verir.", 1, 4, (a, c) => predict(a, c, false)),
  fn("GROWTH", "BÜYÜME", "stats", "bilinen_y'ler; [bilinen_x'ler]; [yeni_x'ler]; [sabit]", "Üstel eğilim boyunca değerleri verir.", 1, 4, (a, c) => predict(a, c, true)),
];
