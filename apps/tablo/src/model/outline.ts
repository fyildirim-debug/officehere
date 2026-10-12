/**
 * Anahat (Veri > Anahat): satır/sütun gruplandırma, grubu çözme, ayrıntıyı
 * gizle/göster, düzey düğmeleri (1-2-3) ve Alt Toplam. Düzeyler eksen
 * deposunda (AxisSizes) tutulur; özet satırı/sütunu grubun altında/sağındadır.
 */
import { colToName } from "@officehere/formula";
import type { AxisSizes } from "./axis";
import type { Rect } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

export type OutlineAxis = "row" | "col";

export interface OutlineGroup {
  level: number;
  start: number;
  end: number;
  /** Özet satırı/sütunu (+/- düğmesinin yeri). */
  summary: number;
  collapsed: boolean;
}

function axisOf(sheet: SheetModel, axis: OutlineAxis): AxisSizes {
  return axis === "row" ? sheet.rows : sheet.cols;
}

/** Eksendeki tüm gruplar (her düzey için ardışık düzey ≥ L öğe dizileri). */
export function outlineGroups(a: AxisSizes): OutlineGroup[] {
  const items = a.levelItems();
  const max = a.maxLevel();
  const out: OutlineGroup[] = [];
  for (let L = 1; L <= max; L++) {
    let start = -1;
    let prev = -2;
    const flush = () => {
      if (start < 0) return;
      const summary = prev + 1;
      const collapsed = a.isCollapsed(summary) || isRunHidden(a, start, prev);
      out.push({ level: L, start, end: prev, summary, collapsed });
    };
    for (const [i, lv] of items) {
      if (lv < L) continue;
      if (i !== prev + 1) {
        flush();
        start = i;
      }
      prev = i;
    }
    flush();
  }
  return out;
}

function isRunHidden(a: AxisSizes, start: number, end: number): boolean {
  for (let i = start; i <= end; i++) if (!a.isHidden(i)) return false;
  return true;
}

/** Gruplandır: öğelerin düzeyini bir artırır (en çok 7). */
export function groupItems(a: AxisSizes, start: number, end: number): void {
  for (let i = start; i <= end; i++) a.setLevel(i, Math.min(7, a.level(i) + 1));
}

/** Grubu Çöz: düzeyi bir azaltır; gizlenmiş öğeler gösterilir. */
export function ungroupItems(a: AxisSizes, start: number, end: number): boolean {
  let any = false;
  const show: number[] = [];
  for (let i = start; i <= end; i++) {
    const lv = a.level(i);
    if (!lv) continue;
    any = true;
    a.setLevel(i, lv - 1);
    if (a.isHidden(i)) show.push(i);
  }
  if (show.length) a.setHiddenMany(show, false);
  a.setCollapsed(end + 1, false);
  return any;
}

/** Grubu daraltır (Ayrıntıyı Gizle) ya da genişletir (Ayrıntıyı Göster). */
export function setGroupCollapsed(a: AxisSizes, g: OutlineGroup, collapse: boolean): void {
  if (collapse) {
    const hide: number[] = [];
    for (let i = g.start; i <= g.end; i++) hide.push(i);
    a.setHiddenMany(hide, true);
    a.setCollapsed(g.summary, true);
    return;
  }
  a.setCollapsed(g.summary, false);
  const show: number[] = [];
  for (let i = g.start; i <= g.end; i++) show.push(i);
  a.setHiddenMany(show, false);
  // İç içe daraltılmış gruplar daraltılmış kalır
  for (const n of outlineGroups(a)) {
    if (n.level <= g.level || n.start < g.start || n.end > g.end || !a.isCollapsed(n.summary)) continue;
    const hide: number[] = [];
    for (let i = n.start; i <= n.end; i++) hide.push(i);
    a.setHiddenMany(hide, true);
  }
}

/** Düzey düğmesi: N düzeyine kadar göster (daha derin düzeyler gizlenir). */
export function showLevel(a: AxisSizes, level: number): void {
  const show: number[] = [];
  const hide: number[] = [];
  for (const [i, lv] of a.levelItems()) (lv >= level ? hide : show).push(i);
  a.setHiddenMany(show, false);
  a.setHiddenMany(hide, true);
  for (const g of outlineGroups(a)) a.setCollapsed(g.summary, g.level >= level);
}

/** Etkin hücreyi içeren en iç grup (ya da özet satırı bu hücre olan grup). */
export function groupAt(a: AxisSizes, i: number): OutlineGroup | null {
  const gs = outlineGroups(a).filter((g) => (i >= g.start && i <= g.end) || g.summary === i);
  if (!gs.length) return null;
  return gs.sort((x, y) => y.level - x.level)[0];
}

/* ------------------------------ Alt Toplam ------------------------------ */

export const SUBTOTAL_FUNCS: { id: number; label: string; name: string }[] = [
  { id: 9, label: "Topla", name: "Toplam" },
  { id: 3, label: "Say", name: "Say" },
  { id: 1, label: "Ortalama", name: "Ortalama" },
  { id: 4, label: "Maks", name: "Maks" },
  { id: 5, label: "Min", name: "Min" },
  { id: 6, label: "Çarpım", name: "Çarpım" },
  { id: 2, label: "Sayıları Say", name: "Say" },
  { id: 7, label: "StdSapma", name: "StdSapma" },
  { id: 10, label: "Var", name: "Var" },
];

export interface SubtotalSpec {
  /** Bölge (başlık satırı dahil). */
  range: Rect;
  /** "Her değişiklikte" sütunu (sayfa sütunu). */
  groupCol: number;
  fn: number;
  /** Alt toplam eklenecek sayfa sütunları. */
  cols: number[];
}

/**
 * Alt toplam satırlarını ekler (alttan yukarı, satır ekleyerek) ve anahat
 * düzeylerini kurar: ayrıntı 3, alt toplam 2, genel toplam 1 (düğmeler 1-2-3).
 * Satır ekleme işlemini çağıran sağlar (geri alma için).
 */
export function applySubtotals(wb: WorkbookModel, sheet: SheetModel, spec: SubtotalSpec, insertRows: (at: number, count: number) => void): Rect {
  const { range: r, groupCol, fn, cols } = spec;
  const fname = SUBTOTAL_FUNCS.find((f) => f.id === fn)?.name ?? "Toplam";
  const key = (row: number) => {
    const v = wb.getValue(sheet, row, groupCol);
    return v === null ? "" : typeof v === "string" ? v.toLocaleUpperCase("tr-TR") : String(v);
  };
  // Grupları bul (veri satırları: başlığın altı)
  const groups: { start: number; end: number; label: string }[] = [];
  let start = r.r1 + 1;
  for (let row = r.r1 + 1; row <= r.r2; row++) {
    if (row === r.r2 || key(row + 1) !== key(row)) {
      const v = wb.getValue(sheet, row, groupCol);
      groups.push({ start, end: row, label: v === null ? "" : String(v) });
      start = row + 1;
    }
  }
  // Alttan yukarı: her grubun altına bir satır ekle
  for (let gi = groups.length - 1; gi >= 0; gi--) insertRows(groups[gi].end + 1, 1);
  // Kaydırılmış grup konumları
  let offset = 0;
  const placed = groups.map((g) => {
    const p = { start: g.start + offset, end: g.end + offset, label: g.label };
    offset++;
    return p;
  });
  const formulaFor = (c: number, a: number, b: number) => `=SUBTOTAL(${fn},${colToName(c)}${a + 1}:${colToName(c)}${b + 1})`;
  for (const g of placed) {
    const row = g.end + 1;
    wb.setInputRaw(sheet, row, groupCol, `${g.label} ${fname}`);
    for (const c of cols) wb.setInputRaw(sheet, row, c, { formula: formulaFor(c, g.start, g.end) });
    sheet.rows.setLevel(row, 1);
    for (let i = g.start; i <= g.end; i++) sheet.rows.setLevel(i, 2);
  }
  // Genel toplam
  const last = placed.length ? placed[placed.length - 1].end + 1 : r.r2;
  const grand = last + 1;
  insertRows(grand, 1);
  wb.setInputRaw(sheet, grand, groupCol, `Genel ${fname}`);
  const first = r.r1 + 1;
  for (const c of cols) wb.setInputRaw(sheet, grand, c, { formula: formulaFor(c, first, last) });
  // Grafik sırası: alt toplam satırları kalın
  return { r1: r.r1, c1: r.c1, r2: grand, c2: r.c2 };
}

/** Tümünü Kaldır: SUBTOTAL formüllü satırları siler ve anahattı temizler. Silinecek satırları döndürür. */
export function subtotalRows(wb: WorkbookModel, sheet: SheetModel, r: Rect): number[] {
  const rows: number[] = [];
  for (let row = r.r1 + 1; row <= r.r2; row++) {
    for (let c = r.c1; c <= r.c2; c++) {
      const f = wb.getFormula(sheet, row, c);
      if (f && /^=\s*SUBTOTAL\(/i.test(f)) {
        rows.push(row);
        break;
      }
    }
  }
  return rows;
}

export { axisOf };
