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
  let c1 = r.c1;
  for (const cEnd of cols) {
    let r1 = r.r1;
    for (const rEnd of rows) {
      if (rEnd - 1 >= r1 && cEnd - 1 >= c1) out.push({ sheet, rect: { r1, c1, r2: rEnd - 1, c2: cEnd - 1 }, scale });
      r1 = rEnd;
    }
    c1 = cEnd;
  }
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
  return rangeToHtml(wb, p.sheet, p.rect, { gridlines: page.printGridlines, headings: page.printHeadings, objects: true });
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
