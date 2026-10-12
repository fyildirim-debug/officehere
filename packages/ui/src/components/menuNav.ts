/**
 * Menü, liste ve galeri gezintisi için saf yardımcılar (DOM'a bağlı değil; birim testli).
 * Klavye işleyicileri (menuKeyboard.ts, RibbonCombo, Gallery) bunları kullanır.
 */

/**
 * current'tan delta yönünde ilk etkin indeks. current aralık dışındaysa (ör. -1) baştan/sondan
 * başlanır. wrap açıksa uçlarda döner; değilse uca gelince -1 döner.
 */
export function nextEnabled(count: number, current: number, delta: 1 | -1, enabled: (i: number) => boolean = () => true, wrap = true): number {
  if (count <= 0) return -1;
  const start = current < 0 || current >= count ? (delta > 0 ? -1 : count) : current;
  for (let step = 1; step <= count; step++) {
    let i = start + delta * step;
    if (wrap) i = ((i % count) + count) % count;
    else if (i < 0 || i >= count) return -1;
    if (enabled(i)) return i;
  }
  return -1;
}

export function firstEnabled(count: number, enabled?: (i: number) => boolean): number {
  return nextEnabled(count, -1, 1, enabled, false);
}

export function lastEnabled(count: number, enabled?: (i: number) => boolean): number {
  return nextEnabled(count, count, -1, enabled, false);
}

/** Harfle atlama için etiket: Türkçe küçük harf, baştaki boşluk ve imler ("(", "…") atılır. */
export function foldLabel(s: string): string {
  return s.toLocaleLowerCase("tr-TR").replace(/^[^\p{L}\p{N}]+/u, "");
}

/**
 * Harfle atlama: query ile başlayan etkin öğe. Tek harfte arama current'tan sonra başlar (aynı
 * harfle başlayan öğeler arasında sırayla dolaşılır, "kk" de böyle sayılır); birden çok harfte
 * current'ın kendisinden başlar ki yazmaya devam edince seçim yerinde kalsın. Bulunamazsa -1.
 */
export function matchByPrefix(labels: string[], current: number, query: string, enabled: (i: number) => boolean = () => true): number {
  const q = query.toLocaleLowerCase("tr-TR");
  if (!q.trim() || !labels.length) return -1;
  const chars = Array.from(q);
  const needle = chars.every((c) => c === chars[0]) ? chars[0] : q;
  const n = labels.length;
  const from = needle.length === 1 ? current + 1 : Math.max(current, 0);
  for (let k = 0; k < n; k++) {
    const i = (((from + k) % n) + n) % n;
    if (enabled(i) && foldLabel(labels[i]).startsWith(needle)) return i;
  }
  return -1;
}

/** Hızlı art arda basılan harfleri biriktirir; timeout ms ara verilince sıfırlanır. */
export function createTypeahead(timeout = 600): (ch: string, now?: number) => string {
  let buf = "";
  let last = -Infinity;
  return (ch, now = Date.now()) => {
    if (now - last > timeout) buf = "";
    last = now;
    buf += ch;
    return buf;
  };
}

export interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export type Direction = "up" | "down" | "left" | "right";

const overlapsV = (a: Box, b: Box) => a.top < b.bottom - 1 && a.bottom > b.top + 1;

/**
 * Liste ve ızgara karışık menülerde (renk paleti, galeri) ok tuşu hedefi.
 * Yukarı/aşağı: bir sonraki satırda sol kenarı en yakın öğe; uçta öbür uca döner (tek sütunlu
 * listede düz sıra). Sol/sağ: aynı satırdaki bitişik öğe; yoksa -1 (alt menü açma/kapama için).
 */
export function spatialNext(boxes: Box[], current: number, dir: Direction, enabled: (i: number) => boolean = () => true): number {
  const cur = boxes[current];
  if (!cur) return -1;
  const idx = boxes.map((_, i) => i).filter((i) => i !== current && enabled(i));
  if (dir === "left" || dir === "right") {
    const row = idx.filter((i) => overlapsV(boxes[i], cur));
    const side = dir === "right" ? row.filter((i) => boxes[i].left >= cur.right - 1) : row.filter((i) => boxes[i].right <= cur.left + 1);
    if (!side.length) return -1;
    return side.reduce((a, b) => (dir === "right" ? (boxes[b].left < boxes[a].left ? b : a) : boxes[b].right > boxes[a].right ? b : a));
  }
  const down = dir === "down";
  const below = (b: Box) => b.top >= cur.bottom - 2 && b.top > cur.top + 1;
  const above = (b: Box) => b.bottom <= cur.top + 2 && b.bottom < cur.bottom - 1;
  let cands = idx.filter((i) => (down ? below(boxes[i]) : above(boxes[i])));
  // Uçta dön: aşağı giderken en üst satıra, yukarı giderken en alt satıra.
  if (!cands.length) cands = idx.filter((i) => (down ? above(boxes[i]) : below(boxes[i])));
  if (!cands.length) return -1;
  // Hedef satır: aşağıda en üstteki, yukarıda en alttaki aday satır.
  const edge = cands.reduce((a, b) => (down ? (boxes[b].top < boxes[a].top ? b : a) : boxes[b].bottom > boxes[a].bottom ? b : a));
  const row = cands.filter((i) => overlapsV(boxes[i], boxes[edge]));
  return row.reduce((a, b) => (Math.abs(boxes[b].left - cur.left) < Math.abs(boxes[a].left - cur.left) ? b : a));
}

/** Liste içinde öğeyi görünür tutmak için gereken kaydırma (scrollTop); değişmesi gerekmiyorsa aynı değer. */
export function scrollToReveal(scrollTop: number, viewHeight: number, itemTop: number, itemHeight: number, pad = 2): number {
  if (itemTop - pad < scrollTop) return Math.max(0, itemTop - pad);
  if (itemTop + itemHeight + pad > scrollTop + viewHeight) return itemTop + itemHeight + pad - viewHeight;
  return scrollTop;
}
