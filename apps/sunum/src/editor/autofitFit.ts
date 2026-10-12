/**
 * "Taşan metni küçült" (normAutofit) hesabı. Sonuç yalnız metne ve kutuya
 * bağlıdır, o anki ölçeğe bağlı DEĞİLDİR: aynı metin + kutu her zaman aynı
 * adımı verir. Böylece ölç → uygula → yeniden ölç döngüsü salınamaz.
 */

/** PowerPoint'in kullandığı biçimde (yazı tipi ölçeği, satır aralığı azaltma) adımlar; büyükten küçüğe. */
export const SHRINK_STEPS: readonly { scale: number; red: number }[] = [
  { scale: 1, red: 0 },
  { scale: 1, red: 0.1 },
  { scale: 0.925, red: 0.1 },
  { scale: 0.85, red: 0.2 },
  { scale: 0.775, red: 0.2 },
  { scale: 0.7, red: 0.2 },
  { scale: 0.625, red: 0.2 },
  { scale: 0.55, red: 0.2 },
  { scale: 0.475, red: 0.2 },
  { scale: 0.4, red: 0.2 },
  { scale: 0.325, red: 0.2 },
  { scale: 0.25, red: 0.2 },
];

/**
 * Kutuya sığan en büyük adımın dizini (ikili arama). `heightAt(i)` i. adımda
 * metnin yüksekliğidir; yükseklik adımla birlikte (genelde) azalır. Hiçbiri
 * sığmazsa en küçük adım döner. En çok ⌈log2(n)⌉ + 2 ölçüm yapılır.
 */
export function pickShrinkStep(heightAt: (i: number) => number, boxH: number, n = SHRINK_STEPS.length): number {
  const fits = (i: number) => heightAt(i) <= boxH + 1;
  if (fits(0)) return 0;
  if (!fits(n - 1)) return n - 1;
  // fits(lo) yanlış, fits(hi) doğru.
  let lo = 0;
  let hi = n - 1;
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (fits(mid)) hi = mid;
    else lo = mid;
  }
  return hi;
}

/** Gövdedeki mevcut (fontScale, lnSpcReduction) değerine karşılık gelen adım. */
export function currentStep(fontScale: number | undefined, red: number | undefined): number {
  const s = fontScale ?? 1;
  const r = red ?? 0;
  const i = SHRINK_STEPS.findIndex((x) => Math.abs(x.scale - s) < 1e-3 && Math.abs(x.red - r) < 1e-3);
  return i;
}

/** Nesne kimliğine dayalı kısa anahtar (önbellek için). */
const ids = new WeakMap<object, number>();
let seq = 0;
export function objKey(o: object | undefined | null): number {
  if (!o) return 0;
  let v = ids.get(o);
  if (v === undefined) {
    v = ++seq;
    ids.set(o, v);
  }
  return v;
}

/**
 * Kısa sürede aynı şekil için tekrar tekrar sessiz güncelleme yapılırsa
 * (beklenmeyen bir ölçüm kararsızlığı) döngüyü keser.
 */
export class CommitGuard {
  private log = new Map<string, number[]>();
  constructor(
    private readonly max = 6,
    private readonly windowMs = 1000,
  ) {}

  allow(id: string, now = Date.now()): boolean {
    const list = (this.log.get(id) ?? []).filter((t) => now - t < this.windowMs);
    if (list.length >= this.max) {
      this.log.set(id, list);
      return false;
    }
    list.push(now);
    this.log.set(id, list);
    return true;
  }
}
