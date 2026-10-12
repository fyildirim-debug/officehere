/**
 * Sayfalara bölünmüş yazdırma: her sayfa ayrı tablo (önce aşağı, sonra
 * yana), sayfa aralığı, kopya sayısı ve yazdırılacak kapsam (etkin sayfa,
 * tüm kitap, seçim). Önizleme de aynı sayfa dikdörtgenlerini kullanır.
 */
import { PAGE_SIZES } from "@officehere/core";
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { computePageBreaks, fitScale, pageInner, preparePrint, printRange, rangeToHtml } from "./print";

export type PrintScope = "active" | "workbook" | "selection";

export interface PrintOptions {
  scope?: PrintScope;
  /** Seçim yazdırılacaksa aralık. */
  selection?: Rect;
  /** Sayfa aralığı (1 tabanlı, dahil). */
  from?: number;
  to?: number;
  copies?: number;
}

export interface PrintPageRef {
  sheet: SheetModel;
  rect: Rect;
  /** Sayfanın ölçeği (sığdırma). */
  scale: number;
}

/** Sayfanın yazdırma sayfaları (önce aşağı, sonra yana). */
export function sheetPages(wb: WorkbookModel, sheet: SheetModel, area?: Rect | null): PrintPageRef[] {
  const r = area ?? printRange(wb, sheet);
  if (!r) return [];
  const scale = fitScale(sheet, r);
  const { rows, cols } = computePageBreaks(wb, sheet, r);
  const out: PrintPageRef[] = [];
  const rowSpans = rows.map((end, i) => [i ? rows[i - 1] : r.r1, end - 1] as const).filter(([a, b]) => b >= a);
  const colSpans = cols.map((end, i) => [i ? cols[i - 1] : r.c1, end - 1] as const).filter(([a, b]) => b >= a);
  // Sayfa sırası: önce aşağı sonra yana (varsayılan) ya da önce yana sonra aşağı
  if (sheet.page.pageOrder === "overThenDown") {
    for (const [r1, r2] of rowSpans) for (const [c1, c2] of colSpans) out.push({ sheet, rect: { r1, c1, r2, c2 }, scale });
  } else for (const [c1, c2] of colSpans) for (const [r1, r2] of rowSpans) out.push({ sheet, rect: { r1, c1, r2, c2 }, scale });
  return out;
}

/** Seçilen kapsamdaki tüm sayfalar. */
export function printPages(wb: WorkbookModel, active: SheetModel, opts: PrintOptions = {}): PrintPageRef[] {
  const scope = opts.scope ?? "active";
  if (scope === "workbook") return wb.sheets.filter((s) => !s.hidden).flatMap((s) => sheetPages(wb, s));
  if (scope === "selection" && opts.selection) return sheetPages(wb, active, opts.selection);
  return sheetPages(wb, active);
}

/** Tek sayfanın HTML'i (yazdırma ve önizleme). */
export function pageHtml(wb: WorkbookModel, p: PrintPageRef): string {
  const page = p.sheet.page;
  const { rowList, colList } = titledLists(p.sheet, p.rect);
  const html = rangeToHtml(wb, p.sheet, p.rect, { gridlines: page.printGridlines && !page.draft, headings: page.printHeadings, objects: !page.draft && !rowList && !colList, rowList, colList });
  return page.blackAndWhite ? `<div style="filter:grayscale(1)">${html}</div>` : html;
}

/** Yazdırma başlıkları: sayfa başlık satırlarını/sütunlarını içermiyorsa önüne eklenir. */
export function titledLists(sheet: SheetModel, r: Rect): { rowList?: number[]; colList?: number[] } {
  const tr = sheet.page.titleRows;
  const tc = sheet.page.titleCols;
  const span = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  const out: { rowList?: number[]; colList?: number[] } = {};
  if (tr && r.r1 > tr.r2) out.rowList = [...span(tr.r1, tr.r2), ...span(r.r1, r.r2)];
  if (tc && r.c1 > tc.c2) out.colList = [...span(tc.c1, tc.c2), ...span(r.c1, r.c2)];
  return out;
}

/** Yazdırılacak tüm sayfaların HTML'i (sayfa aralığı ve kopyalar uygulanmış). */
export function buildPrintHtml(wb: WorkbookModel, active: SheetModel, opts: PrintOptions = {}): string {
  const pages = printPages(wb, active, opts);
  const from = Math.max(1, opts.from ?? 1);
  const to = Math.min(pages.length, opts.to ?? pages.length);
  const chosen = pages.slice(from - 1, to);
  const one = chosen.map((p) => `<div class="tb-print-page" style="zoom:${p.scale}">${pageHtml(wb, p)}</div>`).join("");
  const copies = Math.max(1, Math.min(99, Math.floor(opts.copies ?? 1)));
  return Array.from({ length: copies }, () => one).join("");
}

/** Önizleme kağıdının piksel boyutu ve iç alanı. */
export function paperBox(sheet: SheetModel): { w: number; h: number; inner: { w: number; h: number }; margin: { top: number; left: number } } {
  const p = PAGE_SIZES[sheet.page.paper] ?? PAGE_SIZES.A4;
  const land = sheet.page.orientation === "landscape";
  return { w: land ? p.height : p.width, h: land ? p.width : p.height, inner: pageInner(sheet), margin: { top: sheet.page.margins.top * 96, left: sheet.page.margins.left * 96 } };
}

/** Seçilen ayarlarla tarayıcının yazdırma iletişim kutusunu açar. */
export function printWithOptions(wb: WorkbookModel, active: SheetModel, opts: PrintOptions): void {
  preparePrint(wb, active, buildPrintHtml(wb, active, opts));
  window.print();
}
