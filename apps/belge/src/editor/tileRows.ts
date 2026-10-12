// Birden Çok Sayfa: sayfaların satırlara dağıtılması. Satır başı ancak gövde düzeyinde bir sayfa
// kesmesiyle (blok/satır) başlayabilir; tablo satırı ya da çok sütunlu bölüm içindeki sayfa
// sınırları satır başı olamaz. Bu durumda satır erken biter (boş yuvalar dolguyla tamamlanır)
// ya da gerekirse uzar (tüm satırlar en uzun satır kadar sütun alır).

/**
 * Satır başlangıç sayfaları (ilk değer her zaman 0).
 * @param count sayfa sayısı
 * @param perRow hedef satır başına sayfa
 * @param canStart k. sayfa yeni bir satır başlatabilir mi
 */
export function planTileRows(count: number, perRow: number, canStart: (k: number) => boolean): number[] {
  const starts = [0];
  if (perRow < 1) return starts;
  let s = 0;
  while (s + perRow < count) {
    let e = -1;
    for (let k = s + perRow; k > s; k--)
      if (canStart(k)) {
        e = k;
        break;
      }
    if (e < 0)
      for (let k = s + perRow + 1; k < count; k++)
        if (canStart(k)) {
          e = k;
          break;
        }
    if (e < 0) break;
    starts.push(e);
    s = e;
  }
  return starts;
}

/** Satır başlangıçlarından en uzun satırın sayfa sayısı. */
export function tileColumns(starts: number[], count: number): number {
  let max = 1;
  for (let i = 0; i < starts.length; i++) max = Math.max(max, (starts[i + 1] ?? count) - starts[i]);
  return max;
}

/** Sayfanın satır ve sütunu. */
export function tileCell(starts: number[], i: number): { row: number; col: number } {
  let lo = 0;
  let hi = starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid] <= i) lo = mid;
    else hi = mid - 1;
  }
  return { row: lo, col: i - (starts[lo] ?? 0) };
}

/** Kesme olmadan eşit satırlar (ölçüm öncesi varsayılan). */
export function evenTileRows(count: number, perRow: number): number[] {
  const out: number[] = [];
  for (let k = 0; k < Math.max(1, count); k += Math.max(1, perRow)) out.push(k);
  return out;
}
