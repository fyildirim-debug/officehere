/**
 * Olasılık dağılımları ve testleri: gama, beta, t, ki-kare, F, lognormal,
 * Weibull, hipergeometrik, negatif binom; hata fonksiyonu; bazı finansal ekler.
 * Sayısal yöntemler: Lanczos lnΓ, eksik gama (seri + sürekli kesir),
 * eksik beta (Lentz sürekli kesri), ters fonksiyonlar için ikiye bölme + Newton.
 */
import type { FnContext } from "../context";
import { Errors, type Value } from "../values";
import { collectNumbers, matrix, num, optNum } from "./helpers";
import { fn, type FnDef } from "./registry";
import { normSDist, normSInv } from "./stats";
import { lnGammaHP } from "./special";

export { lnGammaHP };

/* ------------------------------ Temel sayısal ------------------------------ */

/** Γ(x) (negatif tam sayılar dışında). */
export function gammaFn(x: number): number {
  if (x <= 0 && Number.isInteger(x)) throw Errors.NUM;
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gammaFn(1 - x));
  if (x > 171.6) throw Errors.NUM;
  return Math.exp(lnGammaHP(x));
}

/** Düzenlenmiş alt eksik gama P(a, x). */
export function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (x < a + 1) {
    let ap = a;
    let sum = 1 / a;
    let del = sum;
    for (let n = 0; n < 1000; n++) {
      ap++;
      del *= x / ap;
      sum += del;
      if (Math.abs(del) < Math.abs(sum) * 1e-16) break;
    }
    return sum * Math.exp(-x + a * Math.log(x) - lnGammaHP(a));
  }
  return 1 - gammaQcf(a, x);
}

/** Düzenlenmiş üst eksik gama Q(a, x) = 1 − P(a, x) (kuyrukta hassas). */
export function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  return x < a + 1 ? 1 - gammaP(a, x) : gammaQcf(a, x);
}

function gammaQcf(a: number, x: number): number {
  const FPMIN = 1e-300;
  let b = x + 1 - a;
  let c = 1 / FPMIN;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 1000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGammaHP(a)) * h;
}

/** Düzenlenmiş eksik beta I_x(a, b). */
export function betaI(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(lnGammaHP(a + b) - lnGammaHP(a) - lnGammaHP(b) + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a;
  return 1 - (bt * betacf(1 - x, b, a)) / b;
}

function betacf(x: number, a: number, b: number): number {
  const FPMIN = 1e-300;
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 1000; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return h;
}

/** Artan bir F için F(x) = p çözümü (ikiye bölme, sonra ince ayar). */
export function invert(F: (x: number) => number, p: number, lo: number, hi: number): number {
  let a = lo;
  let b = hi;
  // Üst sınırı genişlet
  for (let i = 0; i < 200 && F(b) < p; i++) {
    a = b;
    b *= 2;
  }
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2;
    if (F(m) < p) a = m;
    else b = m;
    if (Math.abs(b - a) <= 1e-15 * Math.max(1, Math.abs(m))) break;
  }
  return (a + b) / 2;
}

/** Yüksek doğruluklu erf (W. J. Cody rasyonel yaklaşımı üzerinden normSDist). */
export function erfHP(x: number): number {
  if (x === 0) return 0;
  if (Math.abs(x) < 0.5) {
    // Taylor serisi: küçük x için doğrudan
    let sum = x;
    let term = x;
    const x2 = x * x;
    for (let n = 1; n < 60; n++) {
      term *= -x2 / n;
      const add = term / (2 * n + 1);
      sum += add;
      if (Math.abs(add) < 1e-17 * Math.abs(sum)) break;
    }
    return (2 / Math.sqrt(Math.PI)) * sum;
  }
  // erf(x) = 2Φ(x√2) − 1
  return 2 * normSDist(x * Math.SQRT2) - 1;
}

/* ------------------------------ Dağılım çekirdekleri ------------------------------ */

/** P(T > |t|) tek kuyruk. */
const tTail = (t: number, df: number) => 0.5 * betaI(df / (df + t * t), df / 2, 0.5);
const tCdf = (t: number, df: number) => (t >= 0 ? 1 - tTail(t, df) : tTail(t, df));
/** P(T > t) sağ kuyruk. */
const tRight = (t: number, df: number) => (t >= 0 ? tTail(t, df) : 1 - tTail(t, df));
const tPdf = (t: number, df: number) => Math.exp(lnGammaHP((df + 1) / 2) - lnGammaHP(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log(1 + (t * t) / df));
const chiCdf = (x: number, k: number) => gammaP(k / 2, x / 2);
const chiPdf = (x: number, k: number) => (x <= 0 ? 0 : Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.log(2) - lnGammaHP(k / 2)));
const fCdf = (x: number, d1: number, d2: number) => (x <= 0 ? 0 : betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2));
const fRight = (x: number, d1: number, d2: number) => (x <= 0 ? 1 : betaI(d2 / (d2 + d1 * x), d2 / 2, d1 / 2));
const fPdf = (x: number, d1: number, d2: number) =>
  x <= 0 ? 0 : Math.exp(0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - (lnGammaHP(d1 / 2) + lnGammaHP(d2 / 2) - lnGammaHP((d1 + d2) / 2)));
const gammaPdf = (x: number, a: number, b: number) => (x < 0 ? 0 : x === 0 ? (a === 1 ? 1 / b : a < 1 ? Infinity : 0) : Math.exp((a - 1) * Math.log(x) - x / b - a * Math.log(b) - lnGammaHP(a)));
const lnChoose = (n: number, k: number) => lnGammaHP(n + 1) - lnGammaHP(k + 1) - lnGammaHP(n - k + 1);

function tInv(p: number, df: number): number {
  if (p <= 0 || p >= 1) throw Errors.NUM;
  if (p === 0.5) return 0;
  if (p < 0.5) return -tInv(1 - p, df);
  return invert((t) => tCdf(t, df), p, 0, 10);
}

function chk(cond: boolean): void {
  if (cond) throw Errors.NUM;
}

const n = num;
const cum = (v: Value, c: FnContext) => !!num(v, c);

function betaArgs(a: Value[], c: FnContext, off: number): { A: number; B: number } {
  const A = optNum(a[off], 0, c);
  const B = optNum(a[off + 1], 1, c);
  chk(A >= B);
  return { A, B };
}

function betaDist(x: number, al: number, be: number, cumulative: boolean, A: number, B: number): number {
  chk(al <= 0 || be <= 0 || x < A || x > B);
  const z = (x - A) / (B - A);
  if (cumulative) return betaI(z, al, be);
  if (z === 0 || z === 1) {
    if ((z === 0 && al < 1) || (z === 1 && be < 1)) throw Errors.NUM;
  }
  return Math.exp((al - 1) * Math.log(z) + (be - 1) * Math.log(1 - z) - (lnGammaHP(al) + lnGammaHP(be) - lnGammaHP(al + be))) / (B - A);
}

function hypgeom(k: number, ns: number, M: number, N: number, cumulative: boolean): number {
  chk(k < 0 || ns <= 0 || M <= 0 || N <= 0 || ns > N || M > N || k > ns || k > M || ns - k > N - M);
  const pmf = (i: number) => Math.exp(lnChoose(M, i) + lnChoose(N - M, ns - i) - lnChoose(N, ns));
  if (!cumulative) return pmf(k);
  let s = 0;
  for (let i = Math.max(0, ns - (N - M)); i <= k; i++) s += pmf(i);
  return Math.min(1, s);
}

function negbinom(f: number, s: number, p: number, cumulative: boolean): number {
  chk(f < 0 || s < 1 || p < 0 || p > 1);
  const pmf = (k: number) => Math.exp(lnChoose(k + s - 1, s - 1) + s * Math.log(p) + k * Math.log(1 - p));
  if (!cumulative) return pmf(f);
  return betaI(p, s, f + 1);
}

function weibull(x: number, al: number, be: number, cumulative: boolean): number {
  chk(x < 0 || al <= 0 || be <= 0);
  const r = Math.pow(x / be, al);
  return cumulative ? 1 - Math.exp(-r) : (al / Math.pow(be, al)) * Math.pow(x, al - 1) * Math.exp(-r);
}

function lognorm(x: number, m: number, s: number, cumulative: boolean): number {
  chk(x <= 0 || s <= 0);
  const z = (Math.log(x) - m) / s;
  return cumulative ? normSDist(z) : Math.exp(-0.5 * z * z) / (x * s * Math.sqrt(2 * Math.PI));
}

/** İki örnek dizisi (testler için). */
function samples(a: Value, c: FnContext): number[] {
  return collectNumbers([a], c);
}

function mean(x: number[]): number {
  return x.reduce((s, v) => s + v, 0) / x.length;
}

function variance(x: number[]): number {
  const m = mean(x);
  return x.reduce((s, v) => s + (v - m) ** 2, 0) / (x.length - 1);
}

function tTest(a: Value[], c: FnContext): number {
  const tails = Math.trunc(n(a[2], c));
  const type = Math.trunc(n(a[3], c));
  chk(tails !== 1 && tails !== 2);
  chk(type < 1 || type > 3);
  let t: number;
  let df: number;
  if (type === 1) {
    const A = matrix(a[0], c).flat();
    const B = matrix(a[1], c).flat();
    if (A.length !== B.length) throw Errors.NA;
    const d: number[] = [];
    A.forEach((v, i) => {
      if (typeof v === "number" && typeof B[i] === "number") d.push(v - (B[i] as number));
    });
    if (d.length < 2) throw Errors.DIV0;
    t = mean(d) / Math.sqrt(variance(d) / d.length);
    df = d.length - 1;
  } else {
    const x = samples(a[0], c);
    const y = samples(a[1], c);
    if (x.length < 2 || y.length < 2) throw Errors.DIV0;
    const v1 = variance(x);
    const v2 = variance(y);
    if (type === 2) {
      df = x.length + y.length - 2;
      const sp = ((x.length - 1) * v1 + (y.length - 1) * v2) / df;
      t = (mean(x) - mean(y)) / Math.sqrt(sp * (1 / x.length + 1 / y.length));
    } else {
      const s1 = v1 / x.length;
      const s2 = v2 / y.length;
      t = (mean(x) - mean(y)) / Math.sqrt(s1 + s2);
      df = (s1 + s2) ** 2 / (s1 ** 2 / (x.length - 1) + s2 ** 2 / (y.length - 1));
    }
  }
  // Fark/örnek varyansı sıfırsa t tanımsızdır (Excel: #SAYI/0!)
  if (!Number.isFinite(t)) throw Errors.DIV0;
  const one = tTail(t, df);
  return tails === 1 ? one : 2 * one;
}

function fTest(a: Value[], c: FnContext): number {
  const x = samples(a[0], c);
  const y = samples(a[1], c);
  if (x.length < 2 || y.length < 2) throw Errors.DIV0;
  const v1 = variance(x);
  const v2 = variance(y);
  if (v1 === 0 || v2 === 0) throw Errors.DIV0;
  const f = v1 / v2;
  const p = fCdf(f, x.length - 1, y.length - 1);
  return 2 * Math.min(p, 1 - p);
}

function chiTest(a: Value[], c: FnContext): number {
  const obs = matrix(a[0], c);
  const exp = matrix(a[1], c);
  const r = obs.length;
  const k = obs[0]?.length ?? 0;
  if (r !== exp.length || k !== (exp[0]?.length ?? 0)) throw Errors.NA;
  let chi = 0;
  for (let i = 0; i < r; i++)
    for (let j = 0; j < k; j++) {
      const o = obs[i][j];
      const e = exp[i][j];
      if (typeof o !== "number" || typeof e !== "number") continue;
      if (e === 0) throw Errors.DIV0;
      chi += (o - e) ** 2 / e;
    }
  const df = r > 1 && k > 1 ? (r - 1) * (k - 1) : r * k - 1;
  if (df < 1) throw Errors.NA;
  return gammaQ(df / 2, chi / 2);
}

function zTest(a: Value[], c: FnContext): number {
  const x = samples(a[0], c);
  if (!x.length) throw Errors.NA;
  const mu = n(a[1], c);
  const sigma = a[2] === undefined || a[2] === null ? Math.sqrt(variance(x)) : n(a[2], c);
  if (!(sigma > 0)) throw Errors.DIV0;
  return 1 - normSDist((mean(x) - mu) / (sigma / Math.sqrt(x.length)));
}

/* ------------------------------ Tanımlar ------------------------------ */

const L = { lift: true as const };
const LX = { lift: true as const, xlfn: true };

export const distributionFunctions: FnDef[] = [
  fn("GAMMA", "GAMA", "stats", "x", "Gama fonksiyonu değerini verir.", 1, 1, (a, c) => gammaFn(n(a[0], c)), LX),
  fn("GAMMALN.PRECISE", "GAMALN.DUYARLI", "stats", "x", "Gama fonksiyonunun doğal logaritmasını verir.", 1, 1, (a, c) => {
    const x = n(a[0], c);
    chk(x <= 0);
    return lnGammaHP(x);
  }, LX),
  fn("GAMMA.DIST", "GAMA.DAĞ", "stats", "x; alfa; beta; kümülatif", "Gama dağılımını verir.", 4, 4, (a, c) => gammaDistImpl(a, c), LX),
  fn("GAMMADIST", "GAMADAĞ", "stats", "x; alfa; beta; kümülatif", "Gama dağılımını verir.", 4, 4, (a, c) => gammaDistImpl(a, c), L),
  fn("GAMMA.INV", "GAMA.TERS", "stats", "olasılık; alfa; beta", "Gama kümülatif dağılımının tersini verir.", 3, 3, (a, c) => gammaInvImpl(a, c), LX),
  fn("GAMMAINV", "GAMATERS", "stats", "olasılık; alfa; beta", "Gama kümülatif dağılımının tersini verir.", 3, 3, (a, c) => gammaInvImpl(a, c), L),
  fn("BETA.DIST", "BETA.DAĞ", "stats", "x; alfa; beta; kümülatif; [A]; [B]", "Beta dağılımını verir.", 4, 6, (a, c) => {
    const { A, B } = betaArgs(a, c, 4);
    return betaDist(n(a[0], c), n(a[1], c), n(a[2], c), cum(a[3], c), A, B);
  }, LX),
  fn("BETADIST", "BETADAĞ", "stats", "x; alfa; beta; [A]; [B]", "Beta kümülatif dağılımını verir.", 3, 5, (a, c) => {
    const { A, B } = betaArgs(a, c, 3);
    return betaDist(n(a[0], c), n(a[1], c), n(a[2], c), true, A, B);
  }, L),
  fn("BETA.INV", "BETA.TERS", "stats", "olasılık; alfa; beta; [A]; [B]", "Beta kümülatif dağılımının tersini verir.", 3, 5, (a, c) => betaInvImpl(a, c), LX),
  fn("BETAINV", "BETATERS", "stats", "olasılık; alfa; beta; [A]; [B]", "Beta kümülatif dağılımının tersini verir.", 3, 5, (a, c) => betaInvImpl(a, c), L),
  fn("BINOMDIST", "BİNOMDAĞ", "stats", "başarı_sayısı; denemeler; başarı_olasılığı; kümülatif", "Binom dağılımı olasılığını verir.", 4, 4, (a, c) => binomImpl(a, c), L),
  fn("BINOM.INV", "BİNOM.TERS", "stats", "denemeler; başarı_olasılığı; alfa", "Kümülatif binom dağılımının ölçüt değerine eşit ya da büyük olan en küçük değeri verir.", 3, 3, (a, c) => critBinom(a, c), LX),
  fn("CRITBINOM", "KRİTİKBİNOM", "stats", "denemeler; başarı_olasılığı; alfa", "Kümülatif binom dağılımının ölçüt değerine eşit ya da büyük olan en küçük değeri verir.", 3, 3, (a, c) => critBinom(a, c), L),
  fn("POISSON", "POISSON", "stats", "x; ortalama; kümülatif", "Poisson dağılımını verir.", 3, 3, (a, c) => {
    const x = Math.floor(n(a[0], c));
    const m = n(a[1], c);
    chk(x < 0 || m < 0);
    if (cum(a[2], c)) return m === 0 ? 1 : gammaQ(x + 1, m);
    return m === 0 ? (x === 0 ? 1 : 0) : Math.exp(-m + x * Math.log(m) - lnGammaHP(x + 1));
  }, L),
  fn("EXPONDIST", "ÜSTELDAĞ", "stats", "x; lambda; kümülatif", "Üstel dağılımı verir.", 3, 3, (a, c) => {
    const x = n(a[0], c);
    const l = n(a[1], c);
    chk(x < 0 || l <= 0);
    return cum(a[2], c) ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x);
  }, L),
  fn("T.DIST", "T.DAĞ", "stats", "x; serbestlik_derecesi; kümülatif", "Student t dağılımını verir.", 3, 3, (a, c) => {
    const x = n(a[0], c);
    const df = Math.floor(n(a[1], c));
    chk(df < 1);
    return cum(a[2], c) ? tCdf(x, df) : tPdf(x, df);
  }, LX),
  fn("T.DIST.2T", "T.DAĞ.2K", "stats", "x; serbestlik_derecesi", "İki kuyruklu Student t dağılımını verir.", 2, 2, (a, c) => {
    const x = n(a[0], c);
    const df = Math.floor(n(a[1], c));
    chk(x < 0 || df < 1);
    return 2 * tTail(x, df);
  }, LX),
  fn("T.DIST.RT", "T.DAĞ.SAĞK", "stats", "x; serbestlik_derecesi", "Sağ kuyruklu Student t dağılımını verir.", 2, 2, (a, c) => {
    const df = Math.floor(n(a[1], c));
    chk(df < 1);
    return tRight(n(a[0], c), df);
  }, LX),
  fn("TDIST", "TDAĞ", "stats", "x; serbestlik_derecesi; yazı_türleri", "Student t dağılımını verir.", 3, 3, (a, c) => {
    const x = n(a[0], c);
    const df = Math.floor(n(a[1], c));
    const tails = Math.floor(n(a[2], c));
    chk(x < 0 || df < 1 || (tails !== 1 && tails !== 2));
    return tails * tTail(x, df);
  }, L),
  fn("T.INV", "T.TERS", "stats", "olasılık; serbestlik_derecesi", "Student t dağılımının sol kuyruklu tersini verir.", 2, 2, (a, c) => {
    const df = Math.floor(n(a[1], c));
    chk(df < 1);
    return tInv(n(a[0], c), df);
  }, LX),
  fn("T.INV.2T", "T.TERS.2K", "stats", "olasılık; serbestlik_derecesi", "Student t dağılımının iki kuyruklu tersini verir.", 2, 2, (a, c) => tInv2(a, c), LX),
  fn("TINV", "TTERS", "stats", "olasılık; serbestlik_derecesi", "Student t dağılımının iki kuyruklu tersini verir.", 2, 2, (a, c) => tInv2(a, c), L),
  fn("CHISQ.DIST", "KİKARE.DAĞ", "stats", "x; serbestlik_derecesi; kümülatif", "Ki-kare dağılımını verir.", 3, 3, (a, c) => {
    const x = n(a[0], c);
    const k = Math.floor(n(a[1], c));
    chk(x < 0 || k < 1);
    return cum(a[2], c) ? chiCdf(x, k) : chiPdf(x, k);
  }, LX),
  fn("CHISQ.DIST.RT", "KİKARE.DAĞ.SAĞK", "stats", "x; serbestlik_derecesi", "Ki-kare dağılımının sağ kuyruk olasılığını verir.", 2, 2, (a, c) => chiRt(a, c), LX),
  fn("CHIDIST", "KİKAREDAĞ", "stats", "x; serbestlik_derecesi", "Ki-kare dağılımının sağ kuyruk olasılığını verir.", 2, 2, (a, c) => chiRt(a, c), L),
  fn("CHISQ.INV", "KİKARE.TERS", "stats", "olasılık; serbestlik_derecesi", "Ki-kare dağılımının sol kuyruklu tersini verir.", 2, 2, (a, c) => {
    const p = n(a[0], c);
    const k = Math.floor(n(a[1], c));
    chk(p < 0 || p >= 1 || k < 1);
    return p === 0 ? 0 : invert((x) => chiCdf(x, k), p, 0, Math.max(10, k * 3));
  }, LX),
  fn("CHISQ.INV.RT", "KİKARE.TERS.SAĞK", "stats", "olasılık; serbestlik_derecesi", "Ki-kare dağılımının sağ kuyruklu tersini verir.", 2, 2, (a, c) => chiInvRt(a, c), LX),
  fn("CHIINV", "KİKARETERS", "stats", "olasılık; serbestlik_derecesi", "Ki-kare dağılımının sağ kuyruklu tersini verir.", 2, 2, (a, c) => chiInvRt(a, c), L),
  fn("CHISQ.TEST", "KİKARE.TEST", "stats", "etkin_aralık; beklenen_aralık", "Bağımsızlık testinin sonucunu verir.", 2, 2, (a, c) => chiTest(a, c), { xlfn: true }),
  fn("CHITEST", "KİKARETEST", "stats", "etkin_aralık; beklenen_aralık", "Bağımsızlık testinin sonucunu verir.", 2, 2, (a, c) => chiTest(a, c)),
  fn("F.DIST", "F.DAĞ", "stats", "x; serbestlik_derecesi1; serbestlik_derecesi2; kümülatif", "F olasılık dağılımını verir.", 4, 4, (a, c) => {
    const x = n(a[0], c);
    const d1 = Math.floor(n(a[1], c));
    const d2 = Math.floor(n(a[2], c));
    chk(x < 0 || d1 < 1 || d2 < 1);
    return cum(a[3], c) ? fCdf(x, d1, d2) : fPdf(x, d1, d2);
  }, LX),
  fn("F.DIST.RT", "F.DAĞ.SAĞK", "stats", "x; serbestlik_derecesi1; serbestlik_derecesi2", "Sağ kuyruklu F olasılık dağılımını verir.", 3, 3, (a, c) => fRt(a, c), LX),
  fn("FDIST", "FDAĞ", "stats", "x; serbestlik_derecesi1; serbestlik_derecesi2", "Sağ kuyruklu F olasılık dağılımını verir.", 3, 3, (a, c) => fRt(a, c), L),
  fn("F.INV", "F.TERS", "stats", "olasılık; serbestlik_derecesi1; serbestlik_derecesi2", "F olasılık dağılımının tersini verir.", 3, 3, (a, c) => {
    const p = n(a[0], c);
    const d1 = Math.floor(n(a[1], c));
    const d2 = Math.floor(n(a[2], c));
    chk(p < 0 || p >= 1 || d1 < 1 || d2 < 1);
    return p === 0 ? 0 : invert((x) => fCdf(x, d1, d2), p, 0, 10);
  }, LX),
  fn("F.INV.RT", "F.TERS.SAĞK", "stats", "olasılık; serbestlik_derecesi1; serbestlik_derecesi2", "Sağ kuyruklu F dağılımının tersini verir.", 3, 3, (a, c) => fInvRt(a, c), LX),
  fn("FINV", "FTERS", "stats", "olasılık; serbestlik_derecesi1; serbestlik_derecesi2", "Sağ kuyruklu F dağılımının tersini verir.", 3, 3, (a, c) => fInvRt(a, c), L),
  fn("F.TEST", "F.TEST", "stats", "dizi1; dizi2", "F testinin sonucunu verir.", 2, 2, (a, c) => fTest(a, c), { xlfn: true }),
  fn("FTEST", "FTEST", "stats", "dizi1; dizi2", "F testinin sonucunu verir.", 2, 2, (a, c) => fTest(a, c)),
  fn("T.TEST", "T.TEST", "stats", "dizi1; dizi2; yazı_türleri; tür", "Student t testiyle ilişkili olasılığı verir.", 4, 4, (a, c) => tTest(a, c), { xlfn: true }),
  fn("TTEST", "TTEST", "stats", "dizi1; dizi2; yazı_türleri; tür", "Student t testiyle ilişkili olasılığı verir.", 4, 4, (a, c) => tTest(a, c)),
  fn("Z.TEST", "Z.TEST", "stats", "dizi; x; [sigma]", "z testinin tek kuyruklu olasılık değerini verir.", 2, 3, (a, c) => zTest(a, c), { xlfn: true }),
  fn("ZTEST", "ZTEST", "stats", "dizi; x; [sigma]", "z testinin tek kuyruklu olasılık değerini verir.", 2, 3, (a, c) => zTest(a, c)),
  fn("LOGNORM.DIST", "LOGNORM.DAĞ", "stats", "x; ortalama; standart_sapma; kümülatif", "Lognormal dağılımı verir.", 4, 4, (a, c) => lognorm(n(a[0], c), n(a[1], c), n(a[2], c), cum(a[3], c)), LX),
  fn("LOGNORMDIST", "LOGNORMDAĞ", "stats", "x; ortalama; standart_sapma", "Kümülatif lognormal dağılımı verir.", 3, 3, (a, c) => lognorm(n(a[0], c), n(a[1], c), n(a[2], c), true), L),
  fn("LOGNORM.INV", "LOGNORM.TERS", "stats", "olasılık; ortalama; standart_sapma", "Lognormal kümülatif dağılımın tersini verir.", 3, 3, (a, c) => logInv(a, c), LX),
  fn("LOGINV", "LOGTERS", "stats", "olasılık; ortalama; standart_sapma", "Lognormal kümülatif dağılımın tersini verir.", 3, 3, (a, c) => logInv(a, c), L),
  fn("WEIBULL", "WEIBULL", "stats", "x; alfa; beta; kümülatif", "Weibull dağılımını verir.", 4, 4, (a, c) => weibull(n(a[0], c), n(a[1], c), n(a[2], c), cum(a[3], c)), L),
  fn("WEIBULL.DIST", "WEIBULL.DAĞ", "stats", "x; alfa; beta; kümülatif", "Weibull dağılımını verir.", 4, 4, (a, c) => weibull(n(a[0], c), n(a[1], c), n(a[2], c), cum(a[3], c)), LX),
  fn("HYPGEOM.DIST", "HİPERGEOM.DAĞ", "stats", "örnek_başarıları; örnek_sayısı; popülasyon_başarıları; popülasyon_sayısı; kümülatif", "Hipergeometrik dağılımı verir.", 5, 5, (a, c) =>
    hypgeom(Math.floor(n(a[0], c)), Math.floor(n(a[1], c)), Math.floor(n(a[2], c)), Math.floor(n(a[3], c)), cum(a[4], c)), LX),
  fn("HYPGEOMDIST", "HİPERGEOMDAĞ", "stats", "örnek_başarıları; örnek_sayısı; popülasyon_başarıları; popülasyon_sayısı", "Hipergeometrik dağılımı verir.", 4, 4, (a, c) =>
    hypgeom(Math.floor(n(a[0], c)), Math.floor(n(a[1], c)), Math.floor(n(a[2], c)), Math.floor(n(a[3], c)), false), L),
  fn("NEGBINOM.DIST", "NEGBİNOM.DAĞ", "stats", "başarısızlık_sayısı; başarı_sayısı; başarı_olasılığı; kümülatif", "Negatif binom dağılımını verir.", 4, 4, (a, c) =>
    negbinom(Math.floor(n(a[0], c)), Math.floor(n(a[1], c)), n(a[2], c), cum(a[3], c)), LX),
  fn("NEGBINOMDIST", "NEGBİNOMDAĞ", "stats", "başarısızlık_sayısı; başarı_sayısı; başarı_olasılığı", "Negatif binom dağılımını verir.", 3, 3, (a, c) =>
    negbinom(Math.floor(n(a[0], c)), Math.floor(n(a[1], c)), n(a[2], c), false), L),
  fn("CONFIDENCE", "GÜVENİRLİK", "stats", "alfa; standart_sapma; boyut", "Popülasyon ortalaması için güven aralığını verir.", 3, 3, (a, c) => {
    const alpha = n(a[0], c);
    const s = n(a[1], c);
    const size = Math.floor(n(a[2], c));
    chk(alpha <= 0 || alpha >= 1 || s <= 0 || size < 1);
    return (normSInv(1 - alpha / 2) * s) / Math.sqrt(size);
  }, L),
  fn("CONFIDENCE.T", "GÜVENİLİRLİK.T", "stats", "alfa; standart_sapma; boyut", "Student t dağılımıyla güven aralığını verir.", 3, 3, (a, c) => {
    const alpha = n(a[0], c);
    const s = n(a[1], c);
    const size = Math.floor(n(a[2], c));
    chk(alpha <= 0 || alpha >= 1 || s <= 0 || size < 1);
    if (size === 1) throw Errors.DIV0;
    return (tInv(1 - alpha / 2, size - 1) * s) / Math.sqrt(size);
  }, LX),
  fn("FISHER", "FISHER", "stats", "x", "Fisher dönüşümünü verir.", 1, 1, (a, c) => {
    const x = n(a[0], c);
    chk(x <= -1 || x >= 1);
    return 0.5 * Math.log((1 + x) / (1 - x));
  }, L),
  fn("FISHERINV", "FISHERTERS", "stats", "y", "Fisher dönüşümünün tersini verir.", 1, 1, (a, c) => Math.tanh(n(a[0], c)), L),
  fn("PHI", "PHI", "stats", "x", "Standart normal dağılımın yoğunluk değerini verir.", 1, 1, (a, c) => Math.exp(-0.5 * n(a[0], c) ** 2) / Math.sqrt(2 * Math.PI), LX),
  fn("GAUSS", "GAUSS", "stats", "x", "Standart normal kümülatif dağılımdan 0,5 küçük değeri verir.", 1, 1, (a, c) => normSDist(n(a[0], c)) - 0.5, LX),
  fn("PERMUTATIONA", "PERMÜTASYONA", "stats", "sayı; sayı_seçilen", "Yinelemeli permütasyon sayısını verir.", 2, 2, (a, c) => {
    const nn = Math.trunc(n(a[0], c));
    const k = Math.trunc(n(a[1], c));
    chk(nn < 0 || k < 0);
    return Math.pow(nn, k);
  }, LX),
  fn("PROB", "OLASILIK", "stats", "x_aralığı; olasılık_aralığı; alt_sınır; [üst_sınır]", "Aralıktaki değerlerin iki sınır arasında olma olasılığını verir.", 3, 4, (a, c) => {
    const xs = matrix(a[0], c).flat();
    const ps = matrix(a[1], c).flat();
    if (xs.length !== ps.length) throw Errors.NA;
    const lo = n(a[2], c);
    const hi = a[3] === undefined || a[3] === null ? lo : n(a[3], c);
    let total = 0;
    let s = 0;
    xs.forEach((x, i) => {
      const p = ps[i];
      if (typeof p !== "number") return;
      chk(p < 0 || p > 1);
      total += p;
      if (typeof x === "number" && x >= lo && x <= hi) s += p;
    });
    if (Math.abs(total - 1) > 1e-7) throw Errors.NUM;
    return s;
  }),
  fn("ERF", "HATAİŞLEV", "engineering", "alt_sınır; [üst_sınır]", "Hata işlevini verir.", 1, 2, (a, c) => {
    const lo = n(a[0], c);
    return a[1] === undefined || a[1] === null ? erfHP(lo) : erfHP(n(a[1], c)) - erfHP(lo);
  }, L),
  fn("ERF.PRECISE", "HATAİŞLEV.DUYARLI", "engineering", "x", "Hata işlevini verir.", 1, 1, (a, c) => erfHP(n(a[0], c)), LX),
  fn("ERFC", "TÜMHATAİŞLEV", "engineering", "x", "Tümleyen hata işlevini verir.", 1, 1, (a, c) => 1 - erfHP(n(a[0], c)), L),
  fn("ERFC.PRECISE", "TÜMHATAİŞLEV.DUYARLI", "engineering", "x", "Tümleyen hata işlevini verir.", 1, 1, (a, c) => 1 - erfHP(n(a[0], c)), LX),
  // Finansal ekler
  fn("FVSCHEDULE", "GDPROGRAM", "financial", "anapara; program", "Bir dizi bileşik faiz uygulandıktan sonra anaparanın gelecekteki değerini verir.", 2, 2, (a, c) => {
    let v = n(a[0], c);
    for (const r of matrix(a[1], c).flat()) {
      if (r === null) continue;
      if (typeof r !== "number") throw Errors.VALUE;
      v *= 1 + r;
    }
    return v;
  }),
  fn("PDURATION", "PSÜRE", "financial", "oran; bd; gd", "Yatırımın belirli bir değere ulaşması için gereken dönem sayısını verir.", 3, 3, (a, c) => {
    const r = n(a[0], c);
    const pv = n(a[1], c);
    const fv = n(a[2], c);
    chk(r <= 0 || pv <= 0 || fv <= 0);
    return (Math.log(fv) - Math.log(pv)) / Math.log(1 + r);
  }, LX),
  fn("RRI", "GERÇEKLEŞENYATIRIMGETİRİSİ", "financial", "dönem_sayısı; bd; gd", "Yatırımın büyümesi için eşdeğer faiz oranını verir.", 3, 3, (a, c) => {
    const np = n(a[0], c);
    const pv = n(a[1], c);
    const fv = n(a[2], c);
    chk(np <= 0 || pv === 0);
    return Math.pow(fv / pv, 1 / np) - 1;
  }, LX),
  fn("ISPMT", "ISPMT", "financial", "oran; dönem; dönem_sayısı; bd", "Belirli bir dönemde ödenen faizi verir (eşit anapara ödemeli).", 4, 4, (a, c) => {
    const r = n(a[0], c);
    const per = n(a[1], c);
    const np = n(a[2], c);
    const pv = n(a[3], c);
    if (np === 0) throw Errors.DIV0;
    return pv * r * (per / np - 1);
  }, L),
  fn("SERIESSUM", "SERİTOPLA", "math", "x; n; m; katsayılar", "Bir kuvvet serisinin toplamını verir.", 4, 4, (a, c) => {
    const x = n(a[0], c);
    const n0 = n(a[1], c);
    const m = n(a[2], c);
    let s = 0;
    matrix(a[3], c)
      .flat()
      .forEach((k, i) => {
        if (typeof k !== "number") throw Errors.VALUE;
        s += k * Math.pow(x, n0 + i * m);
      });
    return s;
  }),
];

function gammaDistImpl(a: Value[], c: FnContext): number {
  const x = n(a[0], c);
  const al = n(a[1], c);
  const be = n(a[2], c);
  chk(x < 0 || al <= 0 || be <= 0);
  return cum(a[3], c) ? gammaP(al, x / be) : gammaPdf(x, al, be);
}

function gammaInvImpl(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const al = n(a[1], c);
  const be = n(a[2], c);
  chk(p < 0 || p >= 1 || al <= 0 || be <= 0);
  if (p === 0) return 0;
  return invert((x) => gammaP(al, x / be), p, 0, Math.max(1, al * be * 4));
}

function betaInvImpl(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const al = n(a[1], c);
  const be = n(a[2], c);
  const { A, B } = betaArgs(a, c, 3);
  chk(p <= 0 || p > 1 || al <= 0 || be <= 0);
  let lo = 0;
  let hi = 1;
  for (let i = 0; i < 200; i++) {
    const m = (lo + hi) / 2;
    if (betaI(m, al, be) < p) lo = m;
    else hi = m;
    if (hi - lo < 1e-16) break;
  }
  return A + ((lo + hi) / 2) * (B - A);
}

function binomImpl(a: Value[], c: FnContext): number {
  const k = Math.floor(n(a[0], c));
  const nn = Math.floor(n(a[1], c));
  const p = n(a[2], c);
  chk(k < 0 || k > nn || p < 0 || p > 1);
  if (cum(a[3], c)) return k >= nn ? 1 : p === 0 ? 1 : p === 1 ? 0 : 1 - betaI(p, k + 1, nn - k);
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === nn ? 1 : 0;
  return Math.exp(lnChoose(nn, k) + k * Math.log(p) + (nn - k) * Math.log(1 - p));
}

function critBinom(a: Value[], c: FnContext): number {
  const nn = Math.floor(n(a[0], c));
  const p = n(a[1], c);
  const alpha = n(a[2], c);
  chk(nn < 0 || p < 0 || p > 1 || alpha <= 0 || alpha >= 1);
  let s = 0;
  for (let k = 0; k <= nn; k++) {
    s += p === 0 ? (k === 0 ? 1 : 0) : p === 1 ? (k === nn ? 1 : 0) : Math.exp(lnChoose(nn, k) + k * Math.log(p) + (nn - k) * Math.log(1 - p));
    if (s >= alpha - 1e-12) return k;
  }
  return nn;
}

function tInv2(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const df = Math.floor(n(a[1], c));
  chk(p <= 0 || p > 1 || df < 1);
  return tInv(1 - p / 2, df);
}

function chiRt(a: Value[], c: FnContext): number {
  const x = n(a[0], c);
  const k = Math.floor(n(a[1], c));
  chk(x < 0 || k < 1);
  return gammaQ(k / 2, x / 2);
}

function chiInvRt(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const k = Math.floor(n(a[1], c));
  chk(p <= 0 || p > 1 || k < 1);
  // Q(x) azalan: −Q artan
  return invert((x) => -gammaQ(k / 2, x / 2), -p, 0, Math.max(10, k * 3));
}

function fRt(a: Value[], c: FnContext): number {
  const x = n(a[0], c);
  const d1 = Math.floor(n(a[1], c));
  const d2 = Math.floor(n(a[2], c));
  chk(x < 0 || d1 < 1 || d2 < 1);
  return fRight(x, d1, d2);
}

function fInvRt(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const d1 = Math.floor(n(a[1], c));
  const d2 = Math.floor(n(a[2], c));
  chk(p <= 0 || p > 1 || d1 < 1 || d2 < 1);
  return invert((x) => -fRight(x, d1, d2), -p, 0, 10);
}

function logInv(a: Value[], c: FnContext): number {
  const p = n(a[0], c);
  const m = n(a[1], c);
  const s = n(a[2], c);
  chk(p <= 0 || p >= 1 || s <= 0);
  return Math.exp(m + s * normSInv(p));
}
