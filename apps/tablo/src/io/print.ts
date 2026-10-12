/** Yazdırma: kullanılan alan (ya da yazdırma alanı) HTML tablosu olarak basılır; sayfa sonları hesabı. */
import { getNative, PAGE_SIZES } from "@officehere/core";
import { colToName, formatValue } from "@officehere/formula";
import { styleToCss } from "../model/clipboard";
import { computeView } from "../model/display";
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { CondFormatter } from "../model/condfmt";
import { objectsExtent, objectsHtml } from "./printObjects";
import { sparkSvgForCell } from "../sparkline/geometry";

export function printRange(wb: WorkbookModel, sheet: SheetModel): Rect | null {
  if (sheet.page.printArea) return sheet.page.printArea;
  const u = wb.usedRange(sheet);
  const o = objectsExtent(sheet);
  if (!u && !o) return null;
  return { r1: 0, c1: 0, r2: Math.max(u?.r2 ?? 0, o?.r2 ?? 0), c2: Math.max(u?.c2 ?? 0, o?.c2 ?? 0) };
}

export function pageInner(sheet: SheetModel): { w: number; h: number } {
  const p = PAGE_SIZES[sheet.page.paper] ?? PAGE_SIZES.A4;
  const land = sheet.page.orientation === "landscape";
  const m = sheet.page.margins;
  const w = (land ? p.height : p.width) - (m.left + m.right) * 96;
  const h = (land ? p.width : p.height) - (m.top + m.bottom) * 96;
  return { w, h };
}

/** Geçerli ölçek kipi (eski fitToWidth = tüm sütunlar). */
export function fitModeOf(sheet: SheetModel): "sheet" | "cols" | "rows" | undefined {
  return sheet.page.fitMode ?? (sheet.page.fitToWidth ? "cols" : undefined);
}

/** Sayfa sonu satır/sütun dizinleri (Sayfa Sonu Önizleme). */
export interface PageBreakInfo {
  /** Her sayfanın bittiği yerden sonraki satır/sütun (son öğe aralığın sonu + 1). */
  rows: number[];
  cols: number[];
  /** Bunlardan el ile eklenenler. */
  manualRows: number[];
  manualCols: number[];
  /** Yazdırılan aralık (boş sayfada null). */
  range: Rect | null;
}

/** Bir eksendeki sayfa sonları: el ile sonlar önce, sığmayan yerlerde otomatik sonlar. */
function axisBreaks(start: number, end: number, size: (i: number) => number, limit: number, manual: Set<number>): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = start; i <= end; i++) {
    const sz = size(i);
    if ((manual.has(i) && i > start) || (acc + sz > limit + 0.5 && acc > 0)) {
      out.push(i);
      acc = 0;
    }
    acc += sz;
  }
  out.push(end + 1);
  return out;
}

export function computePageBreaks(wb: WorkbookModel, sheet: SheetModel, range?: Rect | null): PageBreakInfo {
  const r = range ?? printRange(wb, sheet);
  if (!r) return { rows: [], cols: [], manualRows: [], manualCols: [], range: null };
  const inner = pageInner(sheet);
  const scale = fitScale(sheet, r);
  const mode = fitModeOf(sheet);
  // Tek sayfaya sığdırılan eksende el ile sonlar yok sayılır (Excel gibi)
  const mRows = new Set(mode === "sheet" || mode === "rows" ? [] : (sheet.page.rowBreaks ?? []));
  const mCols = new Set(mode === "sheet" || mode === "cols" ? [] : (sheet.page.colBreaks ?? []));
  // Yazdırma başlıkları her sayfada yinelendiği için sayfanın kullanılabilir alanından düşülür
  const tr = sheet.page.titleRows;
  const tc = sheet.page.titleCols;
  const titleH = tr ? sheet.rows.offset(tr.r2 + 1) - sheet.rows.offset(tr.r1) : 0;
  const titleW = tc ? sheet.cols.offset(tc.c2 + 1) - sheet.cols.offset(tc.c1) : 0;
  const cols = axisBreaks(r.c1, r.c2, (c) => sheet.cols.size(c), Math.max(40, inner.w / scale - titleW), mCols);
  const rows = axisBreaks(r.r1, r.r2, (row) => sheet.rows.size(row), Math.max(20, inner.h / scale - 0.5 - titleH), mRows);
  return { rows, cols, manualRows: rows.filter((x) => mRows.has(x)), manualCols: cols.filter((x) => mCols.has(x)), range: r };
}

/** Sığdırma ölçeği (1 = ölçeksiz): sayfayı / tüm sütunları / tüm satırları tek sayfaya. */
export function fitScale(sheet: SheetModel, r: Rect): number {
  const mode = fitModeOf(sheet);
  if (!mode) return 1;
  let totalW = sheet.page.printHeadings ? 32 : 0;
  for (let c = r.c1; c <= r.c2; c++) totalW += sheet.cols.size(c);
  let totalH = sheet.page.printHeadings ? 20 : 0;
  if (mode !== "cols") totalH += sheet.rows.offset(r.r2 + 1) - sheet.rows.offset(r.r1);
  const inner = pageInner(sheet);
  const sw = totalW > 0 ? inner.w / totalW : 1;
  const sh = totalH > 0 ? inner.h / totalH : 1;
  if (mode === "cols") return Math.min(1, sw);
  if (mode === "rows") return Math.min(1, sh);
  return Math.min(1, sw, sh);
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Aralığı HTML tablosuna çevirir (yazdırma, PDF ve HTML dışa aktarma). */
export function rangeToHtml(wb: WorkbookModel, sheet: SheetModel, r: Rect, opts: { gridlines?: boolean; headings?: boolean; cellAttrs?: boolean; objects?: boolean; rowList?: number[]; colList?: number[] } = {}): string {
  const cf = new CondFormatter(wb);
  const covered = new Set<string>();
  const cols: string[] = [];
  // Yazdırma başlıkları için satır/sütun listesi verilebilir (aralığın önüne başlık satırları)
  const rowList = opts.rowList ?? Array.from({ length: r.r2 - r.r1 + 1 }, (_, i) => r.r1 + i);
  const colList = opts.colList ?? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => r.c1 + i);
  if (opts.headings) cols.push(`<col style="width:32px">`);
  // Tablo genişliği açıkça verilir: yoksa table-layout:fixed uygulanmaz ve sütunlar içeriğe göre kayar
  let tableW = opts.headings ? 32 : 0;
  for (const c of colList) {
    tableW += sheet.cols.size(c);
    cols.push(`<col style="width:${sheet.cols.size(c)}px">`);
  }
  const rows: string[] = [];
  if (opts.headings) {
    const hs = ["<th></th>"];
    for (const c of colList) if (sheet.cols.size(c) > 0) hs.push(`<th>${colToName(c)}</th>`);
    rows.push(`<tr class="hdr" style="height:20px">${hs.join("")}</tr>`);
  }
  for (const row of rowList) {
    const rh = sheet.rows.size(row);
    if (rh === 0) continue;
    const tds: string[] = [];
    if (opts.headings) tds.push(`<th>${row + 1}</th>`);
    for (const c of colList) {
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
      const attrs = opts.cellAttrs ? ` data-r="${row}" data-c="${c}"` : "";
      tds.push(`<td${span}${attrs} style="${css}${opts.objects && sheet.sparklines.length ? sparkCss(wb, sheet, row, c, m) : ""}">${esc(v.text).replace(/\n/g, "<br>")}</td>`);
    }
    rows.push(`<tr style="height:${rh}px">${tds.join("")}</tr>`);
  }
  const table = `<table class="tb-print-table${opts.gridlines ? " has-grid" : ""}" style="width:${tableW}px"><colgroup>${cols.join("")}</colgroup>${rows.join("")}</table>`;
  if (!opts.objects || !(sheet.charts.length || sheet.images.length || sheet.shapes.length)) return table;
  return `<div class="tb-print-sheet">${table}${objectsHtml(wb, sheet, r, !!opts.headings)}</div>`;
}

/** Mini grafikli hücrenin arka planı (SVG). */
function sparkCss(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, m: Rect | null | undefined): string {
  const span = m && m.r1 === row && m.c1 === col ? m : { r1: row, c1: col, r2: row, c2: col };
  let w = 0;
  let h = 0;
  for (let c = span.c1; c <= span.c2; c++) w += sheet.cols.size(c);
  for (let r = span.r1; r <= span.r2; r++) h += sheet.rows.size(r);
  const svg = sparkSvgForCell(wb, sheet, row, col, w, h);
  return svg ? `;background-image:url('data:image/svg+xml;utf8,${encodeURIComponent(svg)}');background-size:100% 100%;background-repeat:no-repeat` : "";
}

/** Yazdırma kökünü (#tb-print) ve @page kuralını hazırlar; temizleme işlevini döndürür. */
export function preparePrint(wb: WorkbookModel, sheet: SheetModel, html?: string): () => void {
  const r = printRange(wb, sheet);
  let host = document.getElementById("tb-print");
  if (!host) {
    host = document.createElement("div");
    host.id = "tb-print";
    document.body.appendChild(host);
  }
  const page = sheet.page;
  if (html !== undefined) {
    // Sayfalara bölünmüş içerik (her sayfa kendi ölçeğini taşır)
    host.innerHTML = html || "<p></p>";
    host.style.zoom = "";
  } else {
    host.innerHTML = r ? rangeToHtml(wb, sheet, r, { gridlines: page.printGridlines, headings: page.printHeadings, objects: true }) : "<p></p>";
    host.style.zoom = r ? String(fitScale(sheet, r)) : "";
  }
  let st = document.getElementById("tb-print-page") as HTMLStyleElement | null;
  if (!st) {
    st = document.createElement("style");
    st.id = "tb-print-page";
    document.head.appendChild(st);
  }
  const m = page.margins;
  const size = page.paper === "Letter" ? "letter" : page.paper === "Legal" ? "legal" : page.paper;
  const box = (where: string, text?: string) => (text ? `@${where} { content: ${hfCss(text, sheet.name, wb.props.title ?? "")}; font: 9pt sans-serif; color: #333; }` : "");
  st.textContent = `@page { size: ${size} ${page.orientation}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; ${box("top-center", page.header)} ${box("bottom-center", page.footer)} }`;
  const host2 = host;
  return () => {
    host2.innerHTML = "";
  };
}

/** Tarayıcının yazdırma iletişim kutusunu açar (PDF olarak kaydet dahil). */
export function printSheet(wb: WorkbookModel, sheet: SheetModel): void {
  preparePrint(wb, sheet);
  window.print();
}

/**
 * PDF olarak dışarı aktarır: masaüstü kabuğu exportPdf sunuyorsa doğrudan
 * .pdf kaydedilir; yoksa yazdırma iletişim kutusu ("PDF olarak kaydet") açılır.
 */
export async function exportPdf(wb: WorkbookModel, sheet: SheetModel, suggestedName: string): Promise<void> {
  const native = getNative();
  if (!native?.exportPdf) {
    printSheet(wb, sheet);
    return;
  }
  const cleanup = preparePrint(wb, sheet);
  try {
    await native.exportPdf(suggestedName, { landscape: sheet.page.orientation === "landscape", pageSize: sheet.page.paper });
  } finally {
    cleanup();
  }
}

/** Üst/alt bilgi kodlarını (&[Sayfa], &[Sayfalar], &[Sekme], &[Dosya], &[Tarih]) metne çevirir. */
export function expandHeaderFooter(text: string, ctx: { page: number; pages: number; sheet: string; file: string }): string {
  const today = new Date().toLocaleDateString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" });
  return text
    .replace(/&\[Sayfalar\]/g, String(ctx.pages))
    .replace(/&\[Sayfa\]/g, String(ctx.page))
    .replace(/&\[Sekme\]/g, ctx.sheet)
    .replace(/&\[Dosya\]/g, ctx.file)
    .replace(/&\[Tarih\]/g, today);
}

/** Üst/alt bilgiyi CSS @page kenar kutusu içeriğine çevirir (sayfa sayaçlarıyla). */
export function hfCss(text: string, sheet: string, file: string): string {
  const q = (s: string) => `"${s.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
  const marked = text.replace(/&\[Sayfalar\]/g, "\u0000\u0001N\u0000").replace(/&\[Sayfa\]/g, "\u0000\u0001P\u0000");
  const parts: string[] = [];
  for (const seg of expandHeaderFooter(marked, { page: 0, pages: 0, sheet, file }).split("\u0000")) {
    if (seg === "\u0001P") parts.push("counter(page)");
    else if (seg === "\u0001N") parts.push("counter(pages)");
    else if (seg) parts.push(q(seg));
  }
  return parts.join(" ") || '""';
}

export function formatForPrint(v: number, code?: string): string {
  return formatValue(v, code).text;
}
