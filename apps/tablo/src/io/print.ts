/** Yazdırma: kullanılan alan (ya da yazdırma alanı) HTML tablosu olarak basılır; sayfa sonları hesabı. */
import { PAGE_SIZES } from "@officehere/core";
import { colToName, formatValue } from "@officehere/formula";
import { styleToCss } from "../model/clipboard";
import { computeView } from "../model/display";
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { CondFormatter } from "../model/condfmt";

export function printRange(wb: WorkbookModel, sheet: SheetModel): Rect | null {
  if (sheet.page.printArea) return sheet.page.printArea;
  const u = wb.usedRange(sheet);
  if (!u) return null;
  return { r1: 0, c1: 0, r2: u.r2, c2: u.c2 };
}

function pageInner(sheet: SheetModel): { w: number; h: number } {
  const p = PAGE_SIZES[sheet.page.paper] ?? PAGE_SIZES.A4;
  const land = sheet.page.orientation === "landscape";
  const m = sheet.page.margins;
  const w = (land ? p.height : p.width) - (m.left + m.right) * 96;
  const h = (land ? p.width : p.height) - (m.top + m.bottom) * 96;
  return { w, h };
}

/** Sayfa sonu satır/sütun dizinleri (Sayfa Sonu Önizleme). */
export function computePageBreaks(wb: WorkbookModel, sheet: SheetModel): { rows: number[]; cols: number[] } {
  const r = printRange(wb, sheet);
  if (!r) return { rows: [], cols: [] };
  const { w, h } = pageInner(sheet);
  const rows: number[] = [];
  const cols: number[] = [];
  let acc = 0;
  for (let c = r.c1; c <= r.c2; c++) {
    const cw = sheet.cols.size(c);
    if (acc + cw > w && acc > 0) {
      cols.push(c);
      acc = 0;
    }
    acc += cw;
  }
  cols.push(r.c2 + 1);
  acc = 0;
  for (let row = r.r1; row <= r.r2; row++) {
    const rh = sheet.rows.size(row);
    if (acc + rh > h && acc > 0) {
      rows.push(row);
      acc = 0;
    }
    acc += rh;
  }
  rows.push(r.r2 + 1);
  return { rows, cols };
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Aralığı HTML tablosuna çevirir (yazdırma, PDF ve HTML dışa aktarma). */
export function rangeToHtml(wb: WorkbookModel, sheet: SheetModel, r: Rect, opts: { gridlines?: boolean; headings?: boolean } = {}): string {
  const cf = new CondFormatter(wb);
  const covered = new Set<string>();
  const cols: string[] = [];
  if (opts.headings) cols.push(`<col style="width:32px">`);
  for (let c = r.c1; c <= r.c2; c++) cols.push(`<col style="width:${sheet.cols.size(c)}px">`);
  const rows: string[] = [];
  if (opts.headings) {
    const hs = ["<th></th>"];
    for (let c = r.c1; c <= r.c2; c++) if (sheet.cols.size(c) > 0) hs.push(`<th>${colToName(c)}</th>`);
    rows.push(`<tr class="hdr">${hs.join("")}</tr>`);
  }
  for (let row = r.r1; row <= r.r2; row++) {
    const rh = sheet.rows.size(row);
    if (rh === 0) continue;
    const tds: string[] = [];
    if (opts.headings) tds.push(`<th>${row + 1}</th>`);
    for (let c = r.c1; c <= r.c2; c++) {
      if (sheet.cols.size(c) === 0) continue;
      if (covered.has(`${row},${c}`)) continue;
      const m = sheet.mergeAt(row, c);
      let span = "";
      if (m && m.r1 === row && m.c1 === c) {
        for (let a = m.r1; a <= m.r2; a++) for (let b = m.c1; b <= m.c2; b++) if (a !== row || b !== c) covered.add(`${a},${b}`);
        span = `${m.r2 > m.r1 ? ` rowspan="${m.r2 - m.r1 + 1}"` : ""}${m.c2 > m.c1 ? ` colspan="${m.c2 - m.c1 + 1}"` : ""}`;
      }
      const v = computeView(wb, sheet, row, c, cf);
      const css = styleToCss({ ...v.style, fill: v.fill, color: v.color, hAlign: v.style.hAlign ?? (v.align === "left" ? undefined : v.align) }, v.isNumber);
      tds.push(`<td${span} style="${css}">${esc(v.text).replace(/\n/g, "<br>")}</td>`);
    }
    rows.push(`<tr style="height:${rh}px">${tds.join("")}</tr>`);
  }
  return `<table class="tb-print-table${opts.gridlines ? " has-grid" : ""}"><colgroup>${cols.join("")}</colgroup>${rows.join("")}</table>`;
}

/** Tarayıcının yazdırma iletişim kutusunu açar (PDF olarak kaydet dahil). */
export function printSheet(wb: WorkbookModel, sheet: SheetModel): void {
  const r = printRange(wb, sheet);
  let host = document.getElementById("tb-print");
  if (!host) {
    host = document.createElement("div");
    host.id = "tb-print";
    document.body.appendChild(host);
  }
  const page = sheet.page;
  const header = page.header ? `<div class="tb-print-hf">${esc(page.header)}</div>` : "";
  const footer = page.footer ? `<div class="tb-print-hf">${esc(page.footer)}</div>` : "";
  host.innerHTML = header + (r ? rangeToHtml(wb, sheet, r, { gridlines: page.printGridlines, headings: page.printHeadings }) : "<p></p>") + footer;
  let st = document.getElementById("tb-print-page") as HTMLStyleElement | null;
  if (!st) {
    st = document.createElement("style");
    st.id = "tb-print-page";
    document.head.appendChild(st);
  }
  const m = page.margins;
  const size = page.paper === "Letter" ? "letter" : page.paper === "Legal" ? "legal" : page.paper;
  st.textContent = `@page { size: ${size} ${page.orientation}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }`;
  window.print();
}

export function formatForPrint(v: number, code?: string): string {
  return formatValue(v, code).text;
}
