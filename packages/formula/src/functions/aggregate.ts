/** Ortak istatistik hesapları (ALTTOPLAM, TOPLAMA ve istatistik fonksiyonları için). */
import { Errors } from "../values";

export function sum(xs: number[]): number {
  let s = 0;
  // Kahan toplamı: kayan nokta hatasını azaltır
  let c = 0;
  for (const x of xs) {
    const y = x - c;
    const t = s + y;
    c = t - s - y;
    s = t;
  }
  return s;
}

export function average(xs: number[]): number {
  if (xs.length === 0) throw Errors.DIV0;
  return sum(xs) / xs.length;
}

export function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n === 0 || (sample && n < 2)) throw Errors.DIV0;
  const m = sum(xs) / n;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s / (sample ? n - 1 : n);
}

export function stdev(xs: number[], sample: boolean): number {
  return Math.sqrt(variance(xs, sample));
}

export function product(xs: number[]): number {
  if (xs.length === 0) return 0;
  let p = 1;
  for (const x of xs) p *= x;
  return p;
}

export function max(xs: number[]): number {
  if (xs.length === 0) return 0;
  let m = -Infinity;
  for (const x of xs) if (x > m) m = x;
  return m;
}

export function min(xs: number[]): number {
  if (xs.length === 0) return 0;
  let m = Infinity;
  for (const x of xs) if (x < m) m = x;
  return m;
}

export function median(xs: number[]): number {
  if (xs.length === 0) throw Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** En sık görülen değerler (ilk görülme sırasına göre). */
export function modes(xs: number[]): number[] {
  const counts = new Map<number, number>();
  let best = 1;
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > best) best = c;
  }
  if (best < 2) throw Errors.NA;
  const out: number[] = [];
  for (const [k, c] of counts) if (c === best) out.push(k);
  return out;
}

export function percentileInc(xs: number[], k: number): number {
  if (xs.length === 0 || k < 0 || k > 1) throw Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const h = (s.length - 1) * k;
  const lo = Math.floor(h);
  const hi = Math.ceil(h);
  return s[lo] + (h - lo) * (s[hi] - s[lo]);
}

export function percentileExc(xs: number[], k: number): number {
  const n = xs.length;
  if (n === 0 || k <= 0 || k >= 1) throw Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const h = (n + 1) * k - 1;
  if (h < 0 || h > n - 1) throw Errors.NUM;
  const lo = Math.floor(h);
  const hi = Math.ceil(h);
  return s[lo] + (h - lo) * (s[hi] - s[lo]);
}

export function kth(xs: number[], k: number, largest: boolean): number {
  const n = Math.ceil(k);
  if (n < 1 || n > xs.length) throw Errors.NUM;
  const s = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return s[n - 1];
}

/** ALTTOPLAM/TOPLAMA fonksiyon numarasına göre hesap. */
export function subtotalBy(code: number, xs: number[], countA: number): number {
  switch (code) {
    case 1:
      return average(xs);
    case 2:
      return xs.length;
    case 3:
      return countA;
    case 4:
      return max(xs);
    case 5:
      return min(xs);
    case 6:
      return product(xs);
    case 7:
      return stdev(xs, true);
    case 8:
      return stdev(xs, false);
    case 9:
      return sum(xs);
    case 10:
      return variance(xs, true);
    case 11:
      return variance(xs, false);
    case 12:
      return median(xs);
    case 13:
      return modes(xs)[0];
    default:
      throw Errors.VALUE;
  }
}

export function covariance(xs: number[], ys: number[], sample: boolean): number {
  const n = xs.length;
  if (n === 0 || n !== ys.length || (sample && n < 2)) throw Errors.DIV0;
  const mx = sum(xs) / n;
  const my = sum(ys) / n;
  let s = 0;
  for (let i = 0; i < n; i++) s += (xs[i] - mx) * (ys[i] - my);
  return s / (sample ? n - 1 : n);
}

export function correl(xs: number[], ys: number[]): number {
  const n = xs.length;
  if (n === 0 || n !== ys.length) throw Errors.DIV0;
  const mx = sum(xs) / n;
  const my = sum(ys) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
    syy += (ys[i] - my) ** 2;
  }
  if (sxx === 0 || syy === 0) throw Errors.DIV0;
  return sxy / Math.sqrt(sxx * syy);
}
