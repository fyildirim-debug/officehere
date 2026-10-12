/**
 * Pano: iç pano verisi (formüller, stiller, birleştirmeler) ve sistem panosu
 * için TSV/HTML üretimi; dışarıdan TSV/HTML ayrıştırma.
 */
import { fontStack } from "@officehere/core";
import { cellKey, keyCol, keyRow, parseCellInput, shiftFormula, type Scalar } from "@officehere/formula";
import { displayText } from "./display";
import type { Ops } from "./ops";
import { rect, rectCols, rectRows } from "./rect";
import { DEFAULT_FONT, DEFAULT_SIZE, type BorderSide, type CellStyle } from "./styles";
import type { Rect } from "./types";
import type { CellInput, SheetModel, WorkbookModel } from "./workbook";
import { copySparkline, moveSparklines } from "../sparkline/sparkOps";

export interface ClipCell {
  input: CellInput;
  value: Scalar;
  style?: CellStyle;
  text: string;
}

export interface ClipPayload {
  id: string;
  sheetId: number;
  sheetName: string;
  src: Rect;
  rows: number;
  cols: number;
  cells: (ClipCell | null)[][];
  colWidths: number[];
  rowHeights: number[];
  merges: Rect[];
}

let lastPayload: ClipPayload | null = null;

export function getInternalClipboard(): ClipPayload | null {
  return lastPayload;
}

export function buildClipboard(wb: WorkbookModel, sheet: SheetModel, r: Rect, remember = true): { payload: ClipPayload; tsv: string; html: string } {
  const ext = wb.dataExtent(sheet);
  // Tüm sütun/satır seçimleri kullanılan alana kırpılır
  const src: Rect = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, ext.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, ext.cols - 1)) };
  const rows = rectRows(src);
  const cols = rectCols(src);
  const cells: (ClipCell | null)[][] = [];
  for (let i = 0; i < rows; i++) {
    const row: (ClipCell | null)[] = [];
    for (let j = 0; j < cols; j++) {
      const rr = src.r1 + i;
      const cc = src.c1 + j;
      const input = wb.getInput(sheet, rr, cc);
      const styleId = wb.styleIdAt(sheet, rr, cc);
      if (input === null && styleId === 0) {
        row.push(null);
        continue;
      }
      row.push({ input, value: wb.getValue(sheet, rr, cc), style: styleId ? wb.styles.get(styleId) : undefined, text: displayText(wb, sheet, rr, cc) });
    }
    cells.push(row);
  }
  const payload: ClipPayload = {
    id: `oh${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    sheetId: sheet.id,
    sheetName: sheet.name,
    src,
    rows,
    cols,
    cells,
    colWidths: Array.from({ length: cols }, (_, j) => sheet.cols.rawSize(src.c1 + j)),
    rowHeights: Array.from({ length: rows }, (_, i) => sheet.rows.rawSize(src.r1 + i)),
    merges: sheet.mergesIn(src).map((m) => ({ r1: m.r1 - src.r1, c1: m.c1 - src.c1, r2: m.r2 - src.r1, c2: m.c2 - src.c1 })),
  };
  if (remember) lastPayload = payload;
  return { payload, tsv: toTsv(cells), html: remember ? toHtml(payload) : "" };
}

function tsvField(s: string): string {
  return /[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function toTsv(cells: (ClipCell | null)[][]): string {
  return cells.map((r) => r.map((c) => tsvField(c?.text ?? "")).join("\t")).join("\r\n") + "\r\n";
}

function borderCss(b?: BorderSide): string | undefined {
  if (!b) return undefined;
  const w = b.style === "medium" ? "2px" : b.style === "thick" ? "3px" : "1px";
  const kind = b.style === "double" ? "double" : b.style.includes("ash") ? "dashed" : b.style.includes("ot") ? "dotted" : "solid";
  return `${b.style === "double" ? "3px" : w} ${kind} ${b.color ?? "#000"}`;
}

export function styleToCss(s: CellStyle | undefined, numeric = false): string {
  const parts: string[] = [];
  const st = s ?? {};
  parts.push(`font-family:${fontStack(st.font ?? DEFAULT_FONT).replace(/"/g, "'")}`);
  parts.push(`font-size:${st.size ?? DEFAULT_SIZE}pt`);
  if (st.bold) parts.push("font-weight:700");
  if (st.italic) parts.push("font-style:italic");
  const deco = [st.underline ? "underline" : "", st.strike ? "line-through" : ""].filter(Boolean).join(" ");
  if (deco) parts.push(`text-decoration:${deco}`);
  if (st.color) parts.push(`color:${st.color}`);
  if (st.fill) parts.push(`background:${st.fill}`);
  const align = st.hAlign && st.hAlign !== "general" ? (st.hAlign === "centerContinuous" ? "center" : st.hAlign) : numeric ? "right" : undefined;
  if (align) parts.push(`text-align:${align}`);
  if (st.vAlign) parts.push(`vertical-align:${st.vAlign === "middle" ? "middle" : st.vAlign}`);
  if (st.wrap) parts.push("white-space:pre-wrap");
  else parts.push("white-space:nowrap");
  const bt = borderCss(st.bTop);
  const bb = borderCss(st.bBottom);
  const bl = borderCss(st.bLeft);
  const br = borderCss(st.bRight);
  if (bt) parts.push(`border-top:${bt}`);
  if (bb) parts.push(`border-bottom:${bb}`);
  if (bl) parts.push(`border-left:${bl}`);
  if (br) parts.push(`border-right:${br}`);
  return parts.join(";");
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function toHtml(p: ClipPayload): string {
  const covered = new Set<string>();
  const rows = p.cells.map((r, i) => {
    const tds: string[] = [];
    r.forEach((c, j) => {
      if (covered.has(`${i},${j}`)) return;
      const m = p.merges.find((x) => x.r1 === i && x.c1 === j);
      let span = "";
      if (m) {
        for (let a = m.r1; a <= m.r2; a++) for (let b = m.c1; b <= m.c2; b++) if (a !== i || b !== j) covered.add(`${a},${b}`);
        span = `${m.r2 > m.r1 ? ` rowspan="${m.r2 - m.r1 + 1}"` : ""}${m.c2 > m.c1 ? ` colspan="${m.c2 - m.c1 + 1}"` : ""}`;
      }
      tds.push(`<td${span} style="${styleToCss(c?.style, typeof c?.value === "number")}">${esc(c?.text ?? "").replace(/\n/g, "<br>")}</td>`);
    });
    return `<tr style="height:${p.rowHeights[i]}px">${tds.join("")}</tr>`;
  });
  const cols = p.colWidths.map((w) => `<col style="width:${w}px">`).join("");
  return `<html><head><meta charset="utf-8"><meta name="generator" content="OfficeHere Tablo"><meta name="officehere-clip" content="${p.id}"></head><body><table style="border-collapse:collapse">${cols}${rows.join("")}</table></body></html>`;
}

/** TSV metnini hücre metinlerine ayırır (tırnaklı alanlar desteklenir). */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let i = 0;
  let quoted = false;
  const t = text.replace(/\r\n?/g, "\n");
  while (i < t.length) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === "") {
      quoted = true;
      i++;
      continue;
    }
    if (ch === "\t") {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
    i++;
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** HTML tablosunu hücre metni + temel stillere ayrıştırır. */
export function parseHtmlTable(html: string): { text: string; style: CellStyle }[][] | null {
  if (typeof DOMParser === "undefined") return null;
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;
  const out: { text: string; style: CellStyle }[][] = [];
  const occupied = new Set<string>();
  Array.from(table.rows).forEach((tr, i) => {
    out[i] ??= [];
    let j = 0;
    Array.from(tr.cells).forEach((td) => {
      while (occupied.has(`${i},${j}`)) j++;
      const st: CellStyle = {};
      const el = td as HTMLElement;
      const css = el.style;
      if (css.fontWeight === "bold" || Number(css.fontWeight) >= 600 || el.querySelector("b,strong")) st.bold = true;
      if (css.fontStyle === "italic" || el.querySelector("i,em")) st.italic = true;
      if (css.color) st.color = cssColor(css.color);
      const bg = css.backgroundColor || css.background;
      if (bg) st.fill = cssColor(bg);
      if (css.textAlign === "center" || css.textAlign === "right" || css.textAlign === "left") st.hAlign = css.textAlign;
      const text = (el.innerText ?? el.textContent ?? "").replace(/ /g, " ").trim();
      out[i][j] = { text, style: st };
      const rs = (td as HTMLTableCellElement).rowSpan || 1;
      const cs = (td as HTMLTableCellElement).colSpan || 1;
      for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) if (a || b) occupied.add(`${i + a},${j + b}`);
      j += cs;
    });
  });
  return out;
}

function cssColor(c: string): string | undefined {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c);
  if (m) {
    if (m[4] !== undefined && Number(m[4]) === 0) return undefined;
    return "#" + [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  if (/^#[0-9a-f]{3,6}$/i.test(c)) return c.toUpperCase();
  return undefined;
}

export interface PasteOptions {
  what: "all" | "values" | "formulas" | "formats" | "widths" | "allExceptBorders" | "valuesAndFormats" | "formulasAndFormats";
  op?: "none" | "add" | "sub" | "mul" | "div";
  skipBlanks?: boolean;
  transpose?: boolean;
}

/** İç pano verisini hedefe yapıştırır. Kesme işlemi motorun moveRange'i ile yapılır. */
export function pastePayload(ops: Ops, sheet: SheetModel, target: Rect, p: ClipPayload, opts: PasteOptions = { what: "all" }, cut = false): Rect {
  const wb = ops.wb;
  const srcSheet = wb.sheetById(p.sheetId);
  if (cut && srcSheet && opts.what === "all" && !opts.transpose) {
    const dr = target.r1 - p.src.r1;
    const dc = target.c1 - p.src.c1;
    const dst = rect(target.r1, target.c1, target.r1 + p.rows - 1, target.c1 + p.cols - 1);
    ops.snapshotOp("Kes ve Yapıştır", () => {
      wb.engine.moveRange(srcSheet.id, p.src.r1, p.src.c1, p.src.r2, p.src.c2, sheet.id, dr, dc);
      // Stiller, notlar, birleştirmeler
      const styles: [number, number | undefined][] = [];
      for (let i = 0; i < p.rows; i++)
        for (let j = 0; j < p.cols; j++) {
          const k = cellKey(p.src.r1 + i, p.src.c1 + j);
          styles.push([cellKey(dst.r1 + i, dst.c1 + j), srcSheet.styles.get(k)]);
          srcSheet.styles.delete(k);
        }
      for (const [k, id] of styles) {
        if (id === undefined) sheet.styles.delete(k);
        else sheet.styles.set(k, id);
      }
      if (srcSheet === sheet) moveSparklines(sheet, p.src, dr, dc, wb);
      srcSheet.merges = srcSheet.merges.filter((m) => !(m.r1 >= p.src.r1 && m.r2 <= p.src.r2 && m.c1 >= p.src.c1 && m.c2 <= p.src.c2));
      for (const m of p.merges) sheet.merges.push({ r1: m.r1 + dst.r1, c1: m.c1 + dst.c1, r2: m.r2 + dst.r1, c2: m.c2 + dst.c1 });
    });
    return dst;
  }
  const srcRows = opts.transpose ? p.cols : p.rows;
  const srcCols = opts.transpose ? p.rows : p.cols;
  // Hedef kaynak boyutunun katıysa döşe
  const tr = rectRows(target);
  const tc = rectCols(target);
  const repR = tr > srcRows && tr % srcRows === 0 ? tr / srcRows : 1;
  const repC = tc > srcCols && tc % srcCols === 0 ? tc / srcCols : 1;
  const dst = rect(target.r1, target.c1, target.r1 + srcRows * repR - 1, target.c1 + srcCols * repC - 1);
  const label = opts.what === "all" ? "Yapıştır" : "Özel Yapıştır";
  const apply = () => {
    for (let ri = 0; ri < repR; ri++)
      for (let ci = 0; ci < repC; ci++)
        for (let i = 0; i < srcRows; i++)
          for (let j = 0; j < srcCols; j++) {
            const si = opts.transpose ? j : i;
            const sj = opts.transpose ? i : j;
            const cell = p.cells[si][sj];
            const row = dst.r1 + ri * srcRows + i;
            const col = dst.c1 + ci * srcCols + j;
            if (opts.skipBlanks && (!cell || cell.input === null)) continue;
            const w = opts.what;
            if (w !== "formats" && w !== "widths") {
              let input: CellInput = cell?.input ?? null;
              const isF = input !== null && typeof input === "object" && "formula" in input;
              if (isF && (w === "values" || w === "valuesAndFormats")) input = cell!.value;
              else if (isF) {
                const dr = row - (p.src.r1 + si);
                const dc = col - (p.src.c1 + sj);
                input = { formula: shiftFormula((input as { formula: string }).formula, dr, dc) };
              }
              if (opts.op && opts.op !== "none") {
                const cur = wb.getValue(sheet, row, col);
                const a = typeof cur === "number" ? cur : 0;
                const b = typeof cell?.value === "number" ? cell.value : 0;
                if (typeof cur === "number" || typeof cell?.value === "number") {
                  input = opts.op === "add" ? a + b : opts.op === "sub" ? a - b : opts.op === "mul" ? a * b : b === 0 ? a : a / b;
                }
              }
              wb.setInputRaw(sheet, row, col, input);
            }
            if (w === "all" || w === "formats" || w === "valuesAndFormats" || w === "formulasAndFormats" || w === "allExceptBorders") {
              let style = cell?.style;
              if (w === "allExceptBorders" && style) style = { ...style, bTop: undefined, bBottom: undefined, bLeft: undefined, bRight: undefined };
              wb.setStyleRaw(sheet, row, col, style ? wb.styles.intern(style) : undefined);
            }
            if (opts.what === "all" && !opts.transpose && srcSheet === sheet && sheet.sparklines.length) copySparkline(sheet, { row: p.src.r1 + si, col: p.src.c1 + sj }, { row, col }, wb);
          }
    if (opts.what === "widths" || opts.what === "all") {
      if (opts.what === "widths") p.colWidths.forEach((w, j) => sheet.cols.set(dst.c1 + j, w));
    }
    if (opts.what === "all" && !opts.transpose) {
      sheet.merges = sheet.merges.filter((m) => !(m.r1 <= dst.r2 && dst.r1 <= m.r2 && m.c1 <= dst.c2 && dst.c1 <= m.c2));
      for (const m of p.merges) sheet.merges.push({ r1: m.r1 + dst.r1, c1: m.c1 + dst.c1, r2: m.r2 + dst.r1, c2: m.c2 + dst.c1 });
    }
  };
  if (opts.what === "widths" || (opts.what === "all" && (p.merges.length || sheet.mergesIn(dst).length))) ops.sheetOp(sheet, label, apply, [dst]);
  else ops.cellsOp(sheet, [dst], label, apply);
  return dst;
}

/** Dış metni (TSV/HTML) hedefe yapıştırır. */
export function pasteText(ops: Ops, sheet: SheetModel, target: Rect, text: string, html?: string): Rect | null {
  const wb = ops.wb;
  const table = html ? parseHtmlTable(html) : null;
  const grid = table ? table.map((r) => r.map((c) => c?.text ?? "")) : parseTsv(text);
  if (!grid.length) return null;
  const rows = grid.length;
  const cols = Math.max(...grid.map((r) => r.length));
  const dst = rect(target.r1, target.c1, target.r1 + rows - 1, target.c1 + cols - 1);
  ops.cellsOp(sheet, [dst], "Yapıştır", () => {
    grid.forEach((r, i) =>
      r.forEach((t, j) => {
        const row = dst.r1 + i;
        const col = dst.c1 + j;
        const p = parseCellInput(t, { numFmt: wb.styleAt(sheet, row, col).numFmt });
        wb.setInputRaw(sheet, row, col, p.kind === "formula" ? { formula: p.formula! } : p.kind === "empty" ? null : (p.value as Scalar));
        const st = table?.[i]?.[j]?.style;
        if (st && Object.keys(st).length) wb.setStyleRaw(sheet, row, col, wb.styles.patch(wb.styleIdAt(sheet, row, col), st));
        if (p.numFmt) wb.setStyleRaw(sheet, row, col, wb.styles.patch(wb.styleIdAt(sheet, row, col), { numFmt: p.numFmt }));
      }),
    );
  });
  return dst;
}

/** HTML'de bizim pano kimliğimiz var mı? */
export function clipIdFromHtml(html: string): string | null {
  const m = /name="officehere-clip" content="([^"]+)"/.exec(html);
  return m ? m[1] : null;
}

export { keyRow, keyCol };
