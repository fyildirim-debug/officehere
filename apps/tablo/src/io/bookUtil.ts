/** Biçim okuyucularının ortak yardımcıları. */
import { MAX_COLS, MAX_ROWS } from "@officehere/formula";
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

/** Dosyadaki sayfa adını geçerli ve benzersiz bir ada çevirir (31 karakter, yasak karakterler "_"). */
export function safeSheetName(wb: WorkbookModel, raw: string): string {
  let base = (raw || "Sayfa").replace(/[\\/?*[\]:]/g, "_").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Sayfa";
  if (!wb.sheetByName(base)) return base;
  for (let n = 2; n < 10000; n++) {
    const suffix = ` (${n})`;
    const cand = base.slice(0, 31 - suffix.length) + suffix;
    if (!wb.sheetByName(cand)) return cand;
  }
  base = wb.nextSheetName();
  return base;
}

/** Bir hücreye yazılacak en çok hücre sayısı (bozuk/kötü niyetli dosyalara karşı). */
export const CELL_BUDGET = 4_000_000;

/**
 * Dev birleştirme eşiği (hücre). Bundan büyükler okumada kullanılan alana kırpılır (bozuk dosyada
 * A1:XFD1048576 olabilir) ve gizli hücrelerinin değeri korunur; yazmada alt hücrelerine ana hücrenin
 * stili verilmez. Birleştirmelerin tamamı exceljs'i atlayarak doğrudan okunur ve yazılır.
 */
export const BIG_MERGE_CELLS = 100_000;

const area = (m: Rect) => (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1);

/** En küçük yığın (r2'ye göre): süpürmede süresi dolan birleştirmeler. */
function heapPush<T extends Rect>(h: T[], x: T): void {
  let i = h.push(x) - 1;
  while (i > 0) {
    const p = (i - 1) >> 1;
    if (h[p].r2 <= x.r2) break;
    h[i] = h[p];
    i = p;
  }
  h[i] = x;
}

function heapPop<T extends Rect>(h: T[]): T {
  const top = h[0];
  const last = h.pop()!;
  if (h.length) {
    let i = 0;
    for (;;) {
      let k = 2 * i + 1;
      if (k >= h.length) break;
      if (k + 1 < h.length && h[k + 1].r2 < h[k].r2) k++;
      if (h[k].r2 >= last.r2) break;
      h[i] = h[k];
      i = k;
    }
    h[i] = last;
  }
  return top;
}

/**
 * Satır süpürmesi: o anki satırı kesen (birbiriyle çakışmayan) birleştirmeler sütuna göre sıralı
 * tutulur; ekleme, çıkarma ve sorgu ikili aramayla. Satırlar azalmayan sırada ilerletilmelidir.
 */
class ActiveMerges<T extends Rect = Rect> {
  /** c1'e göre sıralı; aynı satırı kestikleri ve çakışmadıkları için c2'ye göre de sıralı. */
  private list: T[] = [];
  private heap: T[] = [];

  /** r satırına geçer: r2 < r olanlar çıkar. */
  advance(r: number): void {
    while (this.heap.length && this.heap[0].r2 < r) this.remove(heapPop(this.heap));
  }

  /** Etkin listeden çıkarır (yığındaki kaydı süresi dolunca kendiliğinden düşer). */
  remove(x: T): void {
    const i = this.lower(x.c1);
    const j = this.list[i] === x ? i : this.list.indexOf(x);
    if (j >= 0) this.list.splice(j, 1);
  }

  /** c2 >= c olan ilk birleştirmenin sırası. */
  private lower(c: number): number {
    let lo = 0;
    let hi = this.list.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.list[mid].c2 < c) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  /** [c1, c2] sütunlarıyla kesişen ilk etkin birleştirme. */
  hit(c1: number, c2: number): T | undefined {
    const x = this.list[this.lower(c1)];
    return x && x.c1 <= c2 ? x : undefined;
  }

  /** [c1, c2] sütunlarıyla kesişen tüm etkin birleştirmeler (sütun sırasıyla). */
  hits(c1: number, c2: number): T[] {
    const out: T[] = [];
    for (let i = this.lower(c1); i < this.list.length && this.list[i].c1 <= c2; i++) out.push(this.list[i]);
    return out;
  }

  /** Etkin birleştirmelerle çakışmayan birleştirmeyi ekler. */
  add(m: T): void {
    this.list.splice(this.lower(m.c1), 0, m);
    heapPush(this.heap, m);
  }
}

/** Sayfa sınırları içinde, en az iki hücrelik, tamsayı aralık. */
function validMerge(m: Rect): boolean {
  return Number.isInteger(m.r1) && Number.isInteger(m.c1) && Number.isInteger(m.r2) && Number.isInteger(m.c2) && m.r1 >= 0 && m.c1 >= 0 && m.r1 <= m.r2 && m.c1 <= m.c2 && m.r2 < MAX_ROWS && m.c2 < MAX_COLS && (m.r1 !== m.r2 || m.c1 !== m.c2);
}

/**
 * Geçerli ve birbiriyle çakışmayan birleştirmeler (özgün sıra korunur). Tek hücrelik, ters ya da
 * sayfa dışındaki aralıklar `invalid` sayılıp atılır. Çakışmada sol üstü önde olan kalır (`overlapping`);
 * ancak dev birleştirme (`big` hücreden büyük; bozuk dosyada A1:XFD1048576 olabilir) küçüğe yenilir
 * (`bigDropped`). Süpürme O(n log n) (exceljs'in denetimi her birleştirmede öncekilerin hepsini gezer).
 */
export function disjointMerges(list: readonly Rect[], big = BIG_MERGE_CELLS): { merges: Rect[]; invalid: number; overlapping: number; bigDropped: number } {
  type Item = Rect & { i: number; big: boolean };
  const items: Item[] = [];
  let invalid = 0;
  list.forEach((m, i) => {
    if (validMerge(m)) items.push({ r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2, i, big: area(m) > big });
    else invalid++;
  });
  items.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1 || a.i - b.i);
  const keep = new Uint8Array(list.length);
  const active = new ActiveMerges<Item>();
  let overlapping = 0;
  let bigDropped = 0;
  for (const m of items) {
    active.advance(m.r1);
    if (active.hit(m.c1, m.c2)) {
      const hits = active.hits(m.c1, m.c2);
      if (m.big || hits.some((h) => !h.big)) {
        if (m.big) bigDropped++;
        else overlapping++;
        continue;
      }
      // Küçük birleştirme, önceden kabul edilmiş dev birleştirmelerin yerini alır
      for (const h of hits) {
        active.remove(h);
        keep[h.i] = 0;
        bigDropped++;
      }
    }
    active.add(m);
    keep[m.i] = 1;
  }
  const merges = invalid || overlapping || bigDropped ? list.filter((_, i) => keep[i]) : [...list];
  return { merges, invalid, overlapping, bigDropped };
}

/**
 * Çakışmayan birleştirmelerde hücre sorgusu: hücreyi kapsayan birleştirme. Satır sırasıyla
 * (okuyucunun dolaştığı gibi) sorulunca hücre başına O(log n); geriye dönülürse baştan kurulur.
 */
export class MergeLookup {
  private sorted: Rect[];
  private next = 0;
  private row = -1;
  private active = new ActiveMerges();

  constructor(merges: readonly Rect[]) {
    this.sorted = [...merges].sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  }

  at(r: number, c: number): Rect | undefined {
    if (r < this.row) {
      this.next = 0;
      this.active = new ActiveMerges();
    }
    if (r !== this.row) {
      this.row = r;
      this.active.advance(r);
      while (this.next < this.sorted.length && this.sorted[this.next].r1 <= r) {
        const m = this.sorted[this.next++];
        if (m.r2 >= r) this.active.add(m);
      }
    }
    return this.active.hit(c, c);
  }
}

/**
 * Büyük birleştirmeyi denetler: makul boyuttaki (en çok `limit` hücre, sayfa kenarına
 * dayanmayan) birleştirme olduğu gibi kalır; diğerleri sayfanın kullanılan alanına
 * kırpılır. Tek hücreye iner ya da yine çok büyükse null.
 */
export function clampMerge(r: Rect, used: { rows: number; cols: number }, limit = 1_000_000): Rect | null {
  const size = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1);
  if (size(r) <= limit && r.r2 < 1048575 && r.c2 < 16383) return r;
  const out = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, used.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, used.cols - 1)) };
  if (out.r1 === out.r2 && out.c1 === out.c2) return null;
  return size(out) > limit ? null : out;
}

/** Sayfanın kullanılan alanı (değerli ya da biçimli hücreler). */
export function usedExtent(wb: WorkbookModel, s: SheetModel): { rows: number; cols: number } {
  let rows = 1;
  let cols = 1;
  wb.engine.forEachCell(s.id, (r, c) => {
    if (r + 1 > rows) rows = r + 1;
    if (c + 1 > cols) cols = c + 1;
  });
  for (const k of s.styles.keys()) {
    const r = Math.floor(k / 16384);
    const c = k % 16384;
    if (r + 1 > rows) rows = r + 1;
    if (c + 1 > cols) cols = c + 1;
  }
  return { rows, cols };
}

/** Sayfadaki dev birleştirmeleri kullanılan alana kırpar (clampMerge); değişiklik varsa uyarı metni. */
export function fitMerges(wb: WorkbookModel, s: SheetModel): string | null {
  let used: { rows: number; cols: number } | undefined;
  let note: string | null = null;
  s.merges = s.merges.flatMap((m) => {
    if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= BIG_MERGE_CELLS && m.r2 < 1048575 && m.c2 < 16383) return [m];
    used ??= usedExtent(wb, s);
    const c = clampMerge(m, used);
    if (c === m) return [m];
    note = c ? "çok büyük birleştirilmiş alanlar (kullanılan alana kırpıldı)" : "çok büyük birleştirilmiş alanlar";
    return c ? [c] : [];
  });
  return note;
}

/** Okuma sırasında kullanıcıya anlaşılır hata. */
export class FormatError extends Error {
  /** Teknik ayrıntı (özgün kitaplık iletisi): "Ayrıntılar" bölümünde ve günlükte gösterilir. */
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "FormatError";
    if (detail) this.detail = detail;
  }
}
