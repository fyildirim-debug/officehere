/** Özel fonksiyonlar: yüksek doğruluklu ln Γ (Stirling serisi + öteleme). */

const STIRLING = [1 / 12, -1 / 360, 1 / 1260, -1 / 1680, 1 / 1188, -691 / 360360, 1 / 156, -3617 / 122400];

/** ln Γ(x), x > 0; bağıl hata ~1e-15. */
export function lnGammaHP(x: number): number {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lnGammaHP(1 - x);
  // x ≥ 15 olana dek öteleme: Γ(x) = Γ(x + n) / (x (x+1) … (x+n−1))
  let shift = 0;
  while (x < 15) {
    shift += Math.log(x);
    x += 1;
  }
  const x2 = x * x;
  let ser = 0;
  let p = x;
  for (const k of STIRLING) {
    ser += k / p;
    p *= x2;
  }
  return (x - 0.5) * Math.log(x) - x + 0.5 * Math.log(2 * Math.PI) + ser - shift;
}
