/**
 * HTML tabloları: web sayfasından içe aktarma (her <table> bir sayfa) ve
 * etkin sayfa / tüm sayfalar için biçimli HTML dışa aktarma.
 */
import { parseHtml } from "@officehere/core";
import { MAX_COLS, MAX_ROWS, parseCellInput, cellKey, type Scalar } from "@officehere/formula";
import type { CellStyle } from "../model/styles";
import { WorkbookModel, type SheetModel } from "../model/workbook";
import { FormatError, safeSheetName } from "./bookUtil";
import { decodeText } from "./csv";
import { printRange, rangeToHtml } from "./print";

function cssColor(c: string): string | undefined {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c);
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return undefined;
    return "#" + [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  if (/^#[0-9a-f]{6}$/i.test(c)) return c.toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(c)) return ("#" + c.slice(1).replace(/./g, (x) => x + x)).toUpperCase();
  return undefined;
}

function cellStyleOf(el: HTMLElement): CellStyle {
  const st: CellStyle = {};
  const css = el.style;
  const tag = el.tagName.toLowerCase();
  if (tag === "th" || css.fontWeight === "bold" || Number(css.fontWeight) >= 600 || el.querySelector("b,strong")) st.bold = true;
  if (css.fontStyle === "italic" || el.querySelector("i,em")) st.italic = true;
  if (/underline/.test(css.textDecoration) || el.querySelector("u")) st.underline = "single";
  const color = css.color ? cssColor(css.color) : undefined;
  if (color && color !== "#000000") st.color = color;
  const bg = cssColor(css.backgroundColor || css.background || el.getAttribute("bgcolor") || "");
  if (bg && bg !== "#FFFFFF") st.fill = bg;
  const al = (css.textAlign || el.getAttribute("align") || "").toLowerCase();
  if (al === "center" || al === "right" || al === "left" || al === "justify") st.hAlign = al;
  const fs = /^([\d.]+)(px|pt)$/.exec(css.fontSize ?? "");
  if (fs) st.size = Math.round((fs[2] === "px" ? (Number(fs[1]) * 72) / 96 : Number(fs[1])) * 2) / 2;
  return st;
}

function cellText(el: HTMLElement): string {
  const clone = el.cloneNode(true) as HTMLElement;
  for (const br of Array.from(clone.querySelectorAll("br"))) br.replaceWith("\n");
  return (clone.textContent ?? "").replace(/ /g, " ").replace(/[ \t\r]*\n[ \t\r]*/g, "\n").replace(/[ \t\r]+/g, " ").trim();
}

/** HTML belgesindeki tabloları çalışma kitabına aktarır. */
export function readHtml(data: ArrayBuffer | Uint8Array, baseName = "Sayfa1"): WorkbookModel {
  if (typeof DOMParser === "undefined") throw new FormatError("HTML okuyucusu bu ortamda kullanılamıyor.");
  const doc = parseHtml(decodeText(data));
  const tables = Array.from(doc.querySelectorAll("table")).filter((t) => !t.parentElement?.closest("table")).slice(0, 200);
  if (!tables.length) throw new FormatError("Web sayfasında tablo bulunamadı.");
  const wb = new WorkbookModel(true);
  tables.forEach((table, ti) => {
    const caption = table.caption?.textContent?.trim();
    const sheet = wb.addSheet(safeSheetName(wb, caption || (tables.length === 1 ? baseName : `Tablo${ti + 1}`)));
    readTable(wb, sheet, table);
  });
  wb.engine.recalc();
  return wb;
}

function readTable(wb: WorkbookModel, sheet: SheetModel, table: HTMLTableElement): void {
  const occupied = new Set<number>();
  // Sütun genişlikleri (<col style="width:..px">)
  Array.from(table.querySelectorAll("col")).forEach((col, i) => {
    const w = /^([\d.]+)px$/.exec((col as HTMLElement).style.width ?? "");
    if (w && i < MAX_COLS) sheet.cols.set(i, Math.round(Number(w[1])));
  });
  Array.from(table.rows).forEach((tr, r) => {
    if (r >= MAX_ROWS) return;
    let c = 0;
    for (const td of Array.from(tr.cells)) {
      while (occupied.has(cellKey(r, c))) c++;
      if (c >= MAX_COLS) break;
      const el = td as HTMLTableCellElement;
      const text = cellText(el);
      const st = cellStyleOf(el);
      if (text) {
        const p = parseCellInput(text, {});
        const v: Scalar | { formula: string } = p.kind === "formula" ? text : p.kind === "empty" ? null : (p.value as Scalar);
        wb.setInputRaw(sheet, r, c, v);
        if (p.numFmt) st.numFmt = p.numFmt;
      }
      const link = el.querySelector("a[href]")?.getAttribute("href");
      if (link && /^(https?:|mailto:)/i.test(link)) sheet.links.set(cellKey(r, c), link);
      if (Object.keys(st).length) sheet.styles.set(cellKey(r, c), wb.styles.intern(st));
      const rs = Math.max(1, Math.min(el.rowSpan || 1, 1000));
      const cs = Math.max(1, Math.min(el.colSpan || 1, 1000));
      if (rs > 1 || cs > 1) {
        sheet.merges.push({ r1: r, c1: c, r2: Math.min(MAX_ROWS - 1, r + rs - 1), c2: Math.min(MAX_COLS - 1, c + cs - 1) });
        for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) if (a || b) occupied.add(cellKey(r + a, c + b));
      }
      c += cs;
    }
  });
}

const HTML_CSS =
  "body{font-family:Calibri,Carlito,sans-serif;margin:16px}h2{font-size:15px;margin:20px 0 8px}" +
  "table{border-collapse:collapse;table-layout:fixed;font-family:Calibri,Carlito,sans-serif}td{padding:0 4px;border:1px solid #d4d4d4;overflow:hidden;white-space:nowrap}" +
  ".tb-print-sheet{position:relative;display:inline-block}.tb-print-obj{position:absolute;overflow:hidden}.tb-chart-tip{display:none}";

function escHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"));
}

/** Biçimli HTML: verilen sayfaların kullanılan aralıkları (birden çoksa başlıklarla). */
export function writeHtml(wb: WorkbookModel, sheets: SheetModel[], title: string): string {
  const parts = sheets.map((s) => {
    const r = printRange(wb, s);
    const body = r ? rangeToHtml(wb, s, r, { gridlines: true, objects: true }) : "<p></p>";
    return sheets.length > 1 ? `<section><h2>${escHtml(s.name)}</h2>${body}</section>` : body;
  });
  return `<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="generator" content="OfficeHere Tablo"><title>${escHtml(title)}</title><style>${HTML_CSS}</style></head><body>${parts.join("")}</body></html>`;
}
