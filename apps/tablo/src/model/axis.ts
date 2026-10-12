/**
 * Seyrek satır yüksekliği / sütun genişliği deposu. Varsayılandan farklı
 * boyutlar ve gizli öğeler sıralı dizide tutulur; konum hesapları ikili
 * aramayla O(log n) yapılır (1.048.576 satır için de hızlı).
 */
export class AxisSizes {
  private sizes = new Map<number, number>();
  private hiddenSet = new Set<number>();
  /** Kullanıcının elle boyutlandırdığı öğeler (satırda customHeight); diğerleri içeriğe göre otomatik. */
  private manualSet = new Set<number>();
  private idx: number[] | null = null;
  private pre: number[] = [];

  constructor(
    public defaultSize: number,
    public readonly count: number,
  ) {}

  private invalidate(): void {
    this.idx = null;
  }

  private build(): void {
    const keys = new Set<number>([...this.sizes.keys(), ...this.hiddenSet]);
    const idx = [...keys].filter((k) => k >= 0 && k < this.count).sort((a, b) => a - b);
    const pre = new Array(idx.length + 1);
    pre[0] = 0;
    for (let k = 0; k < idx.length; k++) pre[k + 1] = pre[k] + (this.size(idx[k]) - this.defaultSize);
    this.idx = idx;
    this.pre = pre;
  }

  /** Görünen boyut (gizliyse 0). */
  size(i: number): number {
    if (this.hiddenSet.has(i)) return 0;
    return this.sizes.get(i) ?? this.defaultSize;
  }

  /** Gizlilikten bağımsız boyut. */
  rawSize(i: number): number {
    return this.sizes.get(i) ?? this.defaultSize;
  }

  isCustom(i: number): boolean {
    return this.sizes.has(i);
  }

  /** Boyut atar; manual=false otomatik boyuttur (içerik değişince yeniden hesaplanabilir). */
  set(i: number, px: number, manual = true): void {
    const v = Math.max(0, Math.round(px * 100) / 100);
    if (v === this.defaultSize && !manual) this.sizes.delete(i);
    else if (v === this.defaultSize) this.sizes.delete(i);
    else this.sizes.set(i, v);
    if (manual) this.manualSet.add(i);
    else this.manualSet.delete(i);
    this.invalidate();
  }

  reset(i: number): void {
    this.sizes.delete(i);
    this.manualSet.delete(i);
    this.invalidate();
  }

  /** Elle boyutlandırılmış mı (otomatik sığdırma bunu değiştirmez). */
  isManual(i: number): boolean {
    return this.manualSet.has(i);
  }

  setDefault(px: number): void {
    this.defaultSize = px;
    this.invalidate();
  }

  isHidden(i: number): boolean {
    return this.hiddenSet.has(i);
  }

  setHidden(i: number, hidden: boolean): void {
    if (hidden) this.hiddenSet.add(i);
    else this.hiddenSet.delete(i);
    this.invalidate();
  }

  /** Toplu gizle/göster (tek geçersizleştirme). */
  setHiddenMany(items: Iterable<number>, hidden: boolean): void {
    if (hidden) for (const i of items) this.hiddenSet.add(i);
    else for (const i of items) this.hiddenSet.delete(i);
    this.invalidate();
  }

  hiddenCount(): number {
    return this.hiddenSet.size;
  }

  hiddenItems(): number[] {
    return [...this.hiddenSet].sort((a, b) => a - b);
  }

  manualItems(): number[] {
    return [...this.manualSet];
  }

  customItems(): [number, number][] {
    return [...this.sizes.entries()].sort((a, b) => a[0] - b[0]);
  }

  /** i. öğenin başlangıç konumu. */
  offset(i: number): number {
    if (!this.idx) this.build();
    const idx = this.idx!;
    let lo = 0;
    let hi = idx.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (idx[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.defaultSize + this.pre[lo];
  }

  /** Konumu içeren öğe (gizli öğeler atlanır). */
  indexAt(px: number): number {
    if (px <= 0) return this.firstVisible(0);
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    // Sıfır genişlikli (gizli) öğeleri atla
    while (lo < this.count - 1 && this.size(lo) === 0) lo++;
    return lo;
  }

  firstVisible(from: number): number {
    let i = from;
    while (i < this.count - 1 && this.size(i) === 0) i++;
    return i;
  }

  total(): number {
    return this.offset(this.count);
  }

  /** Ekleme (count > 0) / silme (count < 0) sonrası dizinleri kaydırır. */
  shift(at: number, count: number): void {
    const move = <T>(m: Map<number, T>): Map<number, T> => {
      const out = new Map<number, T>();
      for (const [k, v] of m) {
        if (count > 0) out.set(k >= at ? k + count : k, v);
        else if (k < at) out.set(k, v);
        else if (k >= at - count) out.set(k + count, v);
      }
      return out;
    };
    this.sizes = move(this.sizes);
    const hs = new Map<number, true>();
    for (const k of this.hiddenSet) hs.set(k, true);
    this.hiddenSet = new Set(move(hs).keys());
    const ms = new Map<number, true>();
    for (const k of this.manualSet) ms.set(k, true);
    this.manualSet = new Set(move(ms).keys());
    this.invalidate();
  }

  clone(): AxisSizes {
    const a = new AxisSizes(this.defaultSize, this.count);
    a.sizes = new Map(this.sizes);
    a.hiddenSet = new Set(this.hiddenSet);
    a.manualSet = new Set(this.manualSet);
    return a;
  }

  assign(other: AxisSizes): void {
    this.defaultSize = other.defaultSize;
    this.sizes = new Map(other.sizes);
    this.hiddenSet = new Set(other.hiddenSet);
    this.manualSet = new Set(other.manualSet);
    this.invalidate();
  }
}

/** Excel sütun genişliği (karakter) ↔ piksel (Calibri 11, en geniş rakam 7 px). */
export function widthCharsToPx(w: number): number {
  return Math.round(w * 7 + 5);
}

export function widthPxToChars(px: number): number {
  return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
}

export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100;
}
