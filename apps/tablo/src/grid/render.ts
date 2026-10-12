/**
 * Canvas ızgara çizici: yalnız görünen hücreleri çizer. Sıra: zemin,
 * kılavuz çizgileri, dolgular, taşan metin ve birleştirme zeminleri,
 * kenarlıklar, metin, işaretler; ardından seçim katmanları ve başlıklar.
 */
import { fontStack } from "@officehere/core";
import { colToName, formatGeneral, keyRow, keyCol } from "@officehere/formula";
import type { CondFormatter } from "../model/condfmt";
import { computeView, type CellView } from "../model/display";
import { contains, intersects, MAX_COLS, MAX_ROWS } from "../model/rect";
import { DEFAULT_FONT, DEFAULT_SIZE, fontCss, type BorderSide, type CellStyle } from "../model/styles";
import type { Rect, Selection } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import type { GridLayout, Quadrant } from "./layout";

export interface RenderState {
  wb: WorkbookModel;
  sheet: SheetModel;
  layout: GridLayout;
  cf: CondFormatter;
  dpr: number;
  selection: Selection;
  showSelection: boolean;
  refs: { rect: Rect; color: string }[];
  clip: { rect: Rect; cut: boolean; offset: number } | null;
  fillPreview: Rect | null;
  movePreview: Rect | null;
  editing: { row: number; col: number } | null;
  filterButtons: { row: number; col: number; active: boolean }[];
  spill: Rect | null;
  dark: boolean;
  printArea?: Rect | null;
  pageBreaks?: { rows: number[]; cols: number[] } | null;
}

const COLORS = {
  grid: "#D4D4D4",
  headerBg: "#F5F5F5",
  headerText: "#3B3B3B",
  headerLine: "#D0D0D0",
  headerSelBg: "#D3E6DA",
  headerSelText: "#0F6B39",
  headerFullBg: "#A8D2B8",
  accent: "#107C41",
  selFill: "rgba(20, 110, 60, 0.12)",
};

let measureCtx: CanvasRenderingContext2D | null = null;
export function measureText(text: string, font: string): number {
  if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
  if (!measureCtx) return text.length * 7;
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}

/** Hücre metni için kullanılabilir yazı tipi tanımı. */
export function cellFont(s: CellStyle, z: number): string {
  return fontCss(s, z, fontStack);
}

function snap(v: number, dpr: number): number {
  return Math.round(v * dpr) / dpr;
}

export function renderGrid(ctx: CanvasRenderingContext2D, st: RenderState): void {
  const { layout: L, dpr } = st;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, L.width, L.height);
  ctx.textBaseline = "alphabetic";
  const quads = L.quadrants();
  for (const q of quads) drawQuadrant(ctx, st, q);
  for (const q of quads) drawOverlays(ctx, st, q);
  drawHeaders(ctx, st, quads);
  // Dondurulmuş bölme çizgileri
  ctx.fillStyle = "#9E9E9E";
  if (L.fr > 0) ctx.fillRect(0, snap(L.cellsY, dpr) - 1 / dpr, L.width, 1 / dpr);
  if (L.fc > 0) ctx.fillRect(snap(L.cellsX, dpr) - 1 / dpr, 0, 1 / dpr, L.height);
}

interface RowViews {
  views: (CellView | null)[];
  c0: number;
}

function drawQuadrant(ctx: CanvasRenderingContext2D, st: RenderState, q: Quadrant): void {
  const { layout: L, sheet, wb, dpr } = st;
  if (q.w <= 0 || q.h <= 0) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.clip();
  const z = L.z;
  const px = 1 / dpr;
  const EXTRA = q.frozenCols ? 0 : 24;
  const c0 = Math.max(0, q.c1 - EXTRA);
  const cEnd = Math.min(MAX_COLS - 1, q.c2 + EXTRA);

  // Satır görünümleri
  const rows: RowViews[] = [];
  for (let r = q.r1; r <= q.r2; r++) {
    const views: (CellView | null)[] = [];
    if (L.rowH(r) > 0) {
      for (let c = c0; c <= cEnd; c++) {
        const hasValue = !wb.isEmpty(sheet, r, c);
        const styled = wb.styleIdAt(sheet, r, c) !== 0 || sheet.tables.length > 0 || sheet.condFormats.length > 0;
        views.push(hasValue || styled ? computeView(wb, sheet, r, c, st.cf) : null);
      }
    }
    rows.push({ views, c0 });
  }

  // Kılavuz çizgileri
  if (sheet.showGridlines) {
    ctx.fillStyle = COLORS.grid;
    for (let c = q.c1; c <= q.c2 + 1; c++) {
      const x = snap(L.colX(c), dpr) - px;
      if (c <= q.c2 && L.colW(c) === 0) continue;
      ctx.fillRect(x, q.y, px, q.h);
    }
    for (let r = q.r1; r <= q.r2 + 1; r++) {
      if (r <= q.r2 && L.rowH(r) === 0) continue;
      const y = snap(L.rowY(r), dpr) - px;
      ctx.fillRect(q.x, y, q.w, px);
    }
  }

  const merges = sheet.merges.filter((m) => m.r2 >= q.r1 && m.r1 <= q.r2 && m.c2 >= q.c1 && m.c1 <= q.c2);
  const inMerge = (r: number, c: number) => merges.some((m) => contains(m, r, c));

  // Dolgular ve veri çubukları
  for (let i = 0; i < rows.length; i++) {
    const r = q.r1 + i;
    const rv = rows[i];
    if (!rv.views.length) continue;
    const y = snap(L.rowY(r), dpr);
    const h = snap(L.rowY(r) + L.rowH(r), dpr) - y;
    for (let c = q.c1; c <= q.c2; c++) {
      const v = rv.views[c - c0];
      if (!v) continue;
      if (inMerge(r, c)) continue;
      const x = snap(L.colX(c), dpr);
      const w = snap(L.colX(c) + L.colW(c), dpr) - x;
      if (w <= 0) continue;
      if (v.fill) {
        ctx.fillStyle = v.fill;
        ctx.fillRect(x - px, y - px, w + px, h + px);
      }
      if (v.style.pattern) drawPattern(ctx, v.style.pattern, x, y, w, h);
      if (v.bar) {
        ctx.fillStyle = v.bar.color;
        const bx = x + 2 + (w - 4) * v.bar.from;
        const bw = Math.max(1, (w - 4) * (v.bar.to - v.bar.from));
        const grad = ctx.createLinearGradient(bx, 0, bx + bw, 0);
        grad.addColorStop(0, v.bar.color);
        grad.addColorStop(1, "#FFFFFF");
        ctx.fillStyle = grad;
        ctx.fillRect(bx, y + 2, bw, h - 4);
        ctx.strokeStyle = v.bar.color;
        ctx.lineWidth = px;
        ctx.strokeRect(bx + px / 2, y + 2 + px / 2, bw - px, h - 4 - px);
      }
    }
  }

  // Birleştirilmiş hücre zeminleri
  for (const m of merges) {
    const v = computeView(wb, sheet, m.r1, m.c1, st.cf);
    const rc = L.cellRect(m.r1, m.c1, m.r2, m.c2);
    ctx.fillStyle = v.fill ?? "#FFFFFF";
    ctx.fillRect(snap(rc.x, dpr), snap(rc.y, dpr), snap(rc.w, dpr) - px, snap(rc.h, dpr) - px);
  }

  // Metin taşma alanları
  type Span = { r: number; c: number; x: number; w: number; view: CellView };
  const spans: Span[] = [];
  for (let i = 0; i < rows.length; i++) {
    const r = q.r1 + i;
    const rv = rows[i];
    if (!rv.views.length) continue;
    for (let c = c0; c <= cEnd; c++) {
      const v = rv.views[c - c0];
      if (!v || !v.text || inMerge(r, c)) continue;
      const vis = c >= q.c1 && c <= q.c2;
      const s = v.style;
      const overflowable = typeof v.value === "string" || (st.sheet.showFormulas && v.text);
      const cellW = L.colW(c);
      if (cellW === 0) continue;
      let x = L.colX(c);
      let w = cellW;
      if (overflowable && !s.wrap && !s.shrink && s.hAlign !== "fill" && s.hAlign !== "justify") {
        const tw = measureText(v.text.split("\n")[0], cellFont(s, z)) + 6 * z + (s.indent ?? 0) * 9 * z;
        if (tw > cellW) {
          const isEmpty = (cc: number) => cc >= 0 && cc < MAX_COLS && wb.isEmpty(sheet, r, cc) && !inMerge(r, cc) && !sheet.mergeAt(r, cc);
          if (v.align === "left" || v.align === "center") {
            let need = v.align === "center" ? (tw - cellW) / 2 : tw - cellW;
            let cc = c + 1;
            while (need > 0 && cc <= cEnd + 40 && isEmpty(cc)) {
              need -= L.colW(cc);
              w += L.colW(cc);
              cc++;
            }
          }
          if (v.align === "right" || v.align === "center") {
            let need = v.align === "center" ? (tw - cellW) / 2 : tw - cellW;
            let cc = c - 1;
            while (need > 0 && cc >= 0 && isEmpty(cc) && (q.frozenCols || cc >= L.fc)) {
              need -= L.colW(cc);
              x -= L.colW(cc);
              w += L.colW(cc);
              cc--;
            }
          }
        }
      }
      if (!vis && x + w < L.colX(q.c1)) continue;
      if (!vis && x > L.colX(q.c2) + L.colW(q.c2)) continue;
      spans.push({ r, c, x, w, view: v });
      // Taşan bölgedeki kılavuz çizgilerini ört
      if (w > cellW) {
        const y = snap(L.rowY(r), dpr);
        const h = snap(L.rowY(r) + L.rowH(r), dpr) - y;
        const cx = snap(L.colX(c), dpr);
        const cw = snap(L.colX(c) + cellW, dpr);
        const sx = snap(x, dpr);
        const sx2 = snap(x + w, dpr);
        for (let cc = c0; cc <= cEnd; cc++) {
          if (cc === c) continue;
          const ox = snap(L.colX(cc), dpr);
          const ox2 = snap(L.colX(cc) + L.colW(cc), dpr);
          if (ox2 <= sx || ox >= sx2) continue;
          const nv = rv.views[cc - c0];
          ctx.fillStyle = nv?.fill ?? "#FFFFFF";
          const left = ox <= cx ? ox : ox - px;
          const right = ox2 >= cw ? ox2 - px : ox2;
          ctx.fillRect(left, y, right - left, h - px);
        }
      }
    }
  }

  // Kenarlıklar
  for (let i = 0; i < rows.length; i++) {
    const r = q.r1 + i;
    const rv = rows[i];
    if (!rv.views.length) continue;
    for (let c = Math.max(0, q.c1 - 1); c <= Math.min(MAX_COLS - 1, q.c2 + 1); c++) {
      const v = rv.views[c - c0];
      if (!v) continue;
      const s = v.style;
      if (!s.bTop && !s.bBottom && !s.bLeft && !s.bRight) continue;
      const m = sheet.mergeAt(r, c);
      const x = L.colX(c);
      const y = L.rowY(r);
      const w = L.colW(c);
      const h = L.rowH(r);
      if (w === 0) continue;
      if (s.bTop && (!m || r === m.r1)) drawBorder(ctx, s.bTop, x, y, x + w, y, dpr, "h");
      if (s.bBottom && (!m || r === m.r2)) drawBorder(ctx, s.bBottom, x, y + h, x + w, y + h, dpr, "h");
      if (s.bLeft && (!m || c === m.c1)) drawBorder(ctx, s.bLeft, x, y, x, y + h, dpr, "v");
      if (s.bRight && (!m || c === m.c2)) drawBorder(ctx, s.bRight, x + w, y, x + w, y + h, dpr, "v");
    }
  }
  for (const m of merges) {
    const s = wb.styleAt(sheet, m.r1, m.c1);
    const rc = L.cellRect(m.r1, m.c1, m.r2, m.c2);
    if (s.bTop) drawBorder(ctx, s.bTop, rc.x, rc.y, rc.x + rc.w, rc.y, dpr, "h");
    if (s.bLeft) drawBorder(ctx, s.bLeft, rc.x, rc.y, rc.x, rc.y + rc.h, dpr, "v");
    const sb = wb.styleAt(sheet, m.r2, m.c1).bBottom ?? s.bBottom;
    const sr = wb.styleAt(sheet, m.r1, m.c2).bRight ?? s.bRight;
    if (sb) drawBorder(ctx, sb, rc.x, rc.y + rc.h, rc.x + rc.w, rc.y + rc.h, dpr, "h");
    if (sr) drawBorder(ctx, sr, rc.x + rc.w, rc.y, rc.x + rc.w, rc.y + rc.h, dpr, "v");
  }

  // Metinler
  for (const sp of spans) {
    if (st.editing && st.editing.row === sp.r && st.editing.col === sp.c) continue;
    const y = L.rowY(sp.r);
    const h = L.rowH(sp.r);
    drawCellText(ctx, sp.view, sp.x, y, sp.w, h, L.colX(sp.c), L.colW(sp.c), z, dpr);
  }
  for (const m of merges) {
    if (st.editing && st.editing.row === m.r1 && st.editing.col === m.c1) continue;
    const v = computeView(wb, sheet, m.r1, m.c1, st.cf);
    if (!v.text) continue;
    const rc = L.cellRect(m.r1, m.c1, m.r2, m.c2);
    drawCellText(ctx, v, rc.x, rc.y, rc.w, rc.h, rc.x, rc.w, z, dpr);
  }

  // Not işaretleri (kırmızı üçgen)
  if (sheet.notes.size) {
    ctx.fillStyle = "#E81123";
    for (const k of sheet.notes.keys()) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (r < q.r1 || r > q.r2 || c < q.c1 || c > q.c2) continue;
      const m = sheet.mergeAt(r, c);
      const rc = m ? L.cellRect(m.r1, m.c1, m.r2, m.c2) : L.cellRect(r, c);
      const s = Math.max(4, 6 * z);
      ctx.beginPath();
      ctx.moveTo(rc.x + rc.w - s - px, rc.y);
      ctx.lineTo(rc.x + rc.w - px, rc.y);
      ctx.lineTo(rc.x + rc.w - px, rc.y + s);
      ctx.closePath();
      ctx.fill();
    }
  }

  // Filtre düğmeleri
  for (const fb of st.filterButtons) {
    if (fb.row < q.r1 || fb.row > q.r2 || fb.col < q.c1 || fb.col > q.c2) continue;
    const rc = L.cellRect(fb.row, fb.col);
    drawFilterButton(ctx, rc.x + rc.w - Math.min(rc.h, 18 * z) - 1, rc.y + 1, Math.min(rc.h - 2, 17 * z), fb.active, dpr);
  }
  ctx.restore();
}

function drawPattern(ctx: CanvasRenderingContext2D, p: { type: string; color: string }, x: number, y: number, w: number, h: number): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.strokeStyle = p.color;
  ctx.lineWidth = 1;
  const step = p.type.includes("light") ? 6 : 4;
  ctx.beginPath();
  if (p.type.includes("Horizontal") || p.type.includes("Grid") || p.type.includes("Gray") || p.type === "gray125") for (let yy = y; yy < y + h; yy += step) {
    ctx.moveTo(x, yy + 0.5);
    ctx.lineTo(x + w, yy + 0.5);
  }
  if (p.type.includes("Vertical") || p.type.includes("Grid")) for (let xx = x; xx < x + w; xx += step) {
    ctx.moveTo(xx + 0.5, y);
    ctx.lineTo(xx + 0.5, y + h);
  }
  if (p.type.includes("Down") || p.type.includes("Trellis")) for (let k = -h; k < w; k += step) {
    ctx.moveTo(x + k, y);
    ctx.lineTo(x + k + h, y + h);
  }
  if (p.type.includes("Up") || p.type.includes("Trellis")) for (let k = 0; k < w + h; k += step) {
    ctx.moveTo(x + k, y);
    ctx.lineTo(x + k - h, y + h);
  }
  ctx.stroke();
  ctx.restore();
}

function borderWidth(s: BorderSide["style"]): number {
  switch (s) {
    case "medium":
    case "mediumDashed":
    case "mediumDashDot":
    case "slantDashDot":
      return 2;
    case "thick":
      return 3;
    case "double":
      return 3;
    default:
      return 1;
  }
}

function drawBorder(ctx: CanvasRenderingContext2D, b: BorderSide, x1: number, y1: number, x2: number, y2: number, dpr: number, dir: "h" | "v"): void {
  const color = b.color ?? "#000000";
  const w = borderWidth(b.style) / dpr;
  const px = 1 / dpr;
  ctx.fillStyle = color;
  const sx1 = snap(x1, dpr);
  const sy1 = snap(y1, dpr);
  const sx2 = snap(x2, dpr);
  const sy2 = snap(y2, dpr);
  if (b.style === "double") {
    if (dir === "h") {
      ctx.fillRect(sx1 - px, sy1 - 2 * px, sx2 - sx1 + px, px);
      ctx.fillRect(sx1 - px, sy1, sx2 - sx1 + px, px);
    } else {
      ctx.fillRect(sx1 - 2 * px, sy1 - px, px, sy2 - sy1 + px);
      ctx.fillRect(sx1, sy1 - px, px, sy2 - sy1 + px);
    }
    return;
  }
  const dashed = b.style.toLowerCase().includes("dash") || b.style === "dotted" || b.style === "hair";
  if (!dashed) {
    const off = Math.floor((borderWidth(b.style) - 1) / 2) / dpr + px;
    if (dir === "h") ctx.fillRect(sx1 - px, sy1 - off, sx2 - sx1 + px, w);
    else ctx.fillRect(sx1 - off, sy1 - px, w, sy2 - sy1 + px);
    return;
  }
  const pattern = b.style === "dotted" || b.style === "hair" ? [1, 1] : b.style.includes("DotDot") ? [3, 1, 1, 1, 1, 1] : b.style.includes("Dot") ? [3, 1, 1, 1] : [3, 1];
  const len = dir === "h" ? sx2 - sx1 : sy2 - sy1;
  let pos = 0;
  let k = 0;
  while (pos < len) {
    const seg = pattern[k % pattern.length] * (borderWidth(b.style) > 1 ? 2 : 1.5);
    if (k % 2 === 0) {
      const l = Math.min(seg, len - pos);
      if (dir === "h") ctx.fillRect(sx1 + pos, sy1 - px, l, w);
      else ctx.fillRect(sx1 - px, sy1 + pos, w, l);
    }
    pos += seg;
    k++;
  }
}

/** Kaydırmalı metin satırları. */
export function wrapLines(text: string, font: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    const words = para.split(/(\s+)/);
    let line = "";
    for (const w of words) {
      const test = line + w;
      if (line && measureText(test.trimEnd(), font) > maxW) {
        out.push(line.trimEnd());
        line = w.trimStart();
        // Tek sözcük sığmıyorsa harf harf böl
        while (measureText(line, font) > maxW && line.length > 1) {
          let k = line.length - 1;
          while (k > 1 && measureText(line.slice(0, k), font) > maxW) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = test;
    }
    out.push(line.trimEnd());
  }
  return out;
}

function drawCellText(ctx: CanvasRenderingContext2D, v: CellView, x: number, y: number, w: number, h: number, cellX: number, cellW: number, z: number, dpr: number): void {
  const s = v.style;
  let font = cellFont(s, z);
  const size = (s.size ?? DEFAULT_SIZE) * (96 / 72) * z;
  const pad = 3 * z;
  const indent = (s.indent ?? 0) * 9 * z;
  let text = v.text;
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.font = font;
  ctx.fillStyle = v.color;
  const avail = cellW - pad * 2 - indent;
  // Sayılar sığmazsa ### (Genel biçimde önce basamak azaltılır)
  if (v.isNumber && !s.wrap) {
    let tw = ctx.measureText(text).width;
    if (tw > avail && s.shrink) {
      const k = Math.max(0.3, avail / tw);
      font = fontCss({ ...s, size: (s.size ?? DEFAULT_SIZE) * k }, z, fontStack);
      ctx.font = font;
      tw = ctx.measureText(text).width;
    }
    if (tw > avail) {
      if ((!s.numFmt || s.numFmt === "General") && typeof v.value === "number") {
        let done = false;
        for (let p = 10; p >= 1 && !done; p--) {
          const t = shortGeneral(v.value, p);
          if (ctx.measureText(t).width <= avail) {
            text = t;
            done = true;
          }
        }
        if (!done) text = "#".repeat(Math.max(1, Math.floor(avail / Math.max(1, ctx.measureText("#").width))));
      } else text = "#".repeat(Math.max(1, Math.floor(avail / Math.max(1, ctx.measureText("#").width))));
    }
  }
  if (v.fillChar && !s.wrap) {
    const fw = ctx.measureText(v.fillChar.char).width || 3;
    const tw = ctx.measureText(text).width;
    const n = Math.max(0, Math.floor((avail - tw) / fw));
    text = text.slice(0, v.fillChar.index) + v.fillChar.char.repeat(n) + text.slice(v.fillChar.index);
  }
  const lineH = size * 1.2;
  const lines = s.wrap ? wrapLines(text, font, Math.max(4, cellW - pad * 2 - indent)) : text.split("\n").slice(0, 1).concat(text.includes("\n") && !s.wrap ? [] : []);
  const totalH = lines.length * lineH;
  let ty: number;
  const va = s.vAlign ?? "bottom";
  if (va === "top") ty = y + 2 * z + size * 0.95;
  else if (va === "middle" || va === "justify" || va === "distributed") ty = y + (h - totalH) / 2 + size * 0.95;
  else ty = y + h - totalH - 2 * z + size * 0.95 + (lineH - size) * 0.5;
  const rot = s.rotation ?? 0;
  if (rot && !s.wrap) {
    const angle = rot === 255 ? 0 : rot > 90 ? (rot - 90) * (Math.PI / 180) : -rot * (Math.PI / 180);
    ctx.translate(cellX + cellW / 2, y + h / 2);
    if (rot === 255) {
      // Dikey yığılmış metin
      ctx.textAlign = "center";
      const chars = text.split("");
      chars.forEach((ch, i) => ctx.fillText(ch, 0, -((chars.length - 1) * lineH) / 2 + i * lineH + size / 3));
    } else {
      ctx.rotate(angle);
      ctx.textAlign = "center";
      ctx.fillText(text, 0, size / 3);
    }
    ctx.restore();
    return;
  }
  for (const line of lines) {
    const tw = ctx.measureText(line).width;
    let tx: number;
    if (v.align === "left") tx = cellX + pad + indent;
    else if (v.align === "right") tx = cellX + cellW - pad - tw - indent;
    else tx = cellX + (cellW - tw) / 2;
    if (s.hAlign === "fill" && line) {
      const reps = Math.max(1, Math.floor((cellW - pad * 2) / Math.max(1, tw)));
      ctx.fillText(line.repeat(reps), cellX + pad, ty);
    } else ctx.fillText(line, tx, ty);
    if (s.underline || s.strike) {
      const lw = Math.max(1 / dpr, size / 14);
      if (s.underline) {
        ctx.fillRect(tx, ty + size * 0.12, tw, lw);
        if (s.underline === "double") ctx.fillRect(tx, ty + size * 0.12 + lw * 2, tw, lw);
      }
      if (s.strike) ctx.fillRect(tx, ty - size * 0.3, tw, lw);
    }
    ty += lineH;
  }
  ctx.restore();
}

function shortGeneral(v: number, p: number): string {
  const abs = Math.abs(v);
  if (abs >= 1e11 || (abs < 1e-4 && abs > 0)) {
    const s = v.toExponential(Math.max(0, p - 1));
    const m = /^(-?)(\d)(?:\.(\d+))?e([+-])(\d+)$/.exec(s);
    if (!m) return s;
    const frac = (m[3] ?? "").replace(/0+$/, "");
    return `${m[1]}${m[2]}${frac ? "," + frac : ""}E${m[4]}${m[5].padStart(2, "0")}`;
  }
  const r = Number(v.toPrecision(p));
  return formatGeneral(r).replace(/^(-?\d+),(\d*?)0+$/, "$1,$2").replace(/,$/, "");
}

function drawFilterButton(ctx: CanvasRenderingContext2D, x: number, y: number, s: number, active: boolean, dpr: number): void {
  const px = 1 / dpr;
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(x, y, s, s);
  ctx.strokeStyle = "#ABABAB";
  ctx.lineWidth = px;
  ctx.strokeRect(snap(x, dpr) + px / 2, snap(y, dpr) + px / 2, s - px, s - px);
  ctx.fillStyle = "#444444";
  const cx = x + s / 2;
  const cy = y + s / 2;
  const a = s * 0.18;
  if (active) {
    // Huni simgesi
    ctx.beginPath();
    ctx.moveTo(cx - a * 1.6, cy - a * 1.2);
    ctx.lineTo(cx + a * 1.6, cy - a * 1.2);
    ctx.lineTo(cx + a * 0.3, cy + a * 0.2);
    ctx.lineTo(cx + a * 0.3, cy + a * 1.5);
    ctx.lineTo(cx - a * 0.3, cy + a * 1.1);
    ctx.lineTo(cx - a * 0.3, cy + a * 0.2);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.moveTo(cx - a * 1.3, cy - a * 0.5);
    ctx.lineTo(cx + a * 1.3, cy - a * 0.5);
    ctx.lineTo(cx, cy + a * 0.9);
    ctx.closePath();
    ctx.fill();
  }
}

/* ------------------------------------------------------------------ */
/* Seçim ve vurgu katmanları                                            */
/* ------------------------------------------------------------------ */

function clipQuad(ctx: CanvasRenderingContext2D, q: Quadrant): void {
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.clip();
}

function visibleRect(st: RenderState, q: Quadrant, r: Rect): { x: number; y: number; w: number; h: number } | null {
  const L = st.layout;
  const vr: Rect = { r1: q.r1, c1: q.c1, r2: q.r2, c2: q.c2 };
  if (!intersects(vr, r)) return null;
  const r1 = Math.max(r.r1, q.r1 - 1 < 0 ? 0 : r.r1 < q.r1 ? q.r1 - 1 : r.r1);
  const c1 = Math.max(r.c1, r.c1 < q.c1 ? q.c1 - 1 : r.c1);
  const r2 = Math.min(r.r2, q.r2 + 1, MAX_ROWS - 1);
  const c2 = Math.min(r.c2, q.c2 + 1, MAX_COLS - 1);
  const x = r.c1 < q.c1 ? q.x - 10 : L.colX(c1);
  const y = r.r1 < q.r1 ? q.y - 10 : L.rowY(r1);
  const x2 = r.c2 > q.c2 ? q.x + q.w + 10 : L.colX(c2) + L.colW(c2);
  const y2 = r.r2 > q.r2 ? q.y + q.h + 10 : L.rowY(r2) + L.rowH(r2);
  return { x, y, w: x2 - x, h: y2 - y };
}

function drawOverlays(ctx: CanvasRenderingContext2D, st: RenderState, q: Quadrant): void {
  const { layout: L, dpr } = st;
  const px = 1 / dpr;
  ctx.save();
  clipQuad(ctx, q);
  // Taşan dizi çerçevesi
  if (st.spill) {
    const vr = visibleRect(st, q, st.spill);
    if (vr) {
      ctx.strokeStyle = "#4A8FE7";
      ctx.lineWidth = px;
      ctx.setLineDash([]);
      ctx.strokeRect(snap(vr.x, dpr) - px / 2, snap(vr.y, dpr) - px / 2, snap(vr.w, dpr), snap(vr.h, dpr));
    }
  }
  if (st.showSelection) {
    const sel = st.selection;
    const multi = sel.ranges.length > 1;
    for (let i = 0; i < sel.ranges.length; i++) {
      const r = sel.ranges[i];
      const vr = visibleRect(st, q, r);
      if (!vr) continue;
      const single = r.r1 === r.r2 && r.c1 === r.c2;
      const merge = st.sheet.mergeAt(r.r1, r.c1);
      const isMerge = merge && merge.r1 === r.r1 && merge.c1 === r.c1 && merge.r2 === r.r2 && merge.c2 === r.c2;
      if (!single && !isMerge) {
        ctx.fillStyle = COLORS.selFill;
        ctx.fillRect(vr.x, vr.y, vr.w, vr.h);
      }
    }
    // Etkin hücre beyaz kalır
    const a = sel.active;
    const am = st.sheet.mergeAt(a.row, a.col);
    const ar = am ?? { r1: a.row, c1: a.col, r2: a.row, c2: a.col };
    const avr = visibleRect(st, q, ar);
    if (avr && (sel.ranges.length > 1 || !(sel.ranges[0].r1 === ar.r1 && sel.ranges[0].r2 === ar.r2 && sel.ranges[0].c1 === ar.c1 && sel.ranges[0].c2 === ar.c2))) {
      // Seçim dolgusunu etkin hücreden kaldırmak için yeniden çizmek yerine ince çerçeve
      ctx.strokeStyle = COLORS.accent;
      ctx.lineWidth = px;
      ctx.strokeRect(snap(avr.x, dpr) + px / 2, snap(avr.y, dpr) + px / 2, snap(avr.w, dpr) - 2 * px, snap(avr.h, dpr) - 2 * px);
    }
    // Son aralığın kalın çerçevesi ve doldurma tutamacı
    const last = sel.ranges[sel.ranges.length - 1];
    const vr = visibleRect(st, q, last);
    if (vr) {
      ctx.strokeStyle = COLORS.accent;
      ctx.lineWidth = multi ? px * 1 : 2;
      const x = snap(vr.x, dpr);
      const y = snap(vr.y, dpr);
      const w = snap(vr.w, dpr);
      const h = snap(vr.h, dpr);
      if (multi) ctx.strokeRect(x - px / 2, y - px / 2, w, h);
      else ctx.strokeRect(x - px, y - px, w, h);
      if (!multi && last.r2 >= q.r1 && last.r2 <= q.r2 && last.c2 >= q.c1 && last.c2 <= q.c2) {
        const hx = x + w - px;
        const hy = y + h - px;
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(hx - 4, hy - 4, 8, 8);
        ctx.fillStyle = COLORS.accent;
        ctx.fillRect(hx - 3, hy - 3, 6, 6);
      }
    }
  }
  // Formül başvuruları
  for (const ref of st.refs) {
    const vr = visibleRect(st, q, ref.rect);
    if (!vr) continue;
    ctx.fillStyle = hexAlpha(ref.color, 0.1);
    ctx.fillRect(vr.x, vr.y, vr.w, vr.h);
    ctx.strokeStyle = ref.color;
    ctx.lineWidth = 1.5;
    ctx.strokeRect(snap(vr.x, dpr), snap(vr.y, dpr), snap(vr.w, dpr) - px, snap(vr.h, dpr) - px);
    ctx.fillStyle = ref.color;
    const s = 4;
    for (const [cx, cy] of [
      [vr.x, vr.y],
      [vr.x + vr.w, vr.y],
      [vr.x, vr.y + vr.h],
      [vr.x + vr.w, vr.y + vr.h],
    ]) ctx.fillRect(cx - s / 2, cy - s / 2, s, s);
  }
  // Kopyalama kayan çerçevesi
  if (st.clip) {
    const vr = visibleRect(st, q, st.clip.rect);
    if (vr) {
      ctx.lineWidth = 2;
      ctx.strokeStyle = COLORS.accent;
      ctx.setLineDash([4, 3]);
      ctx.lineDashOffset = -st.clip.offset;
      ctx.strokeRect(snap(vr.x, dpr), snap(vr.y, dpr), snap(vr.w, dpr) - px, snap(vr.h, dpr) - px);
      ctx.setLineDash([]);
    }
  }
  for (const pr of [st.fillPreview, st.movePreview]) {
    if (!pr) continue;
    const vr = visibleRect(st, q, pr);
    if (!vr) continue;
    ctx.strokeStyle = "#5F5F5F";
    ctx.lineWidth = 1.5;
    ctx.setLineDash([3, 2]);
    ctx.strokeRect(snap(vr.x, dpr), snap(vr.y, dpr), snap(vr.w, dpr) - px, snap(vr.h, dpr) - px);
    ctx.setLineDash([]);
  }
  // Sayfa sonları (Sayfa Sonu Önizleme)
  if (st.pageBreaks) {
    ctx.strokeStyle = "#1F5FBF";
    ctx.lineWidth = 2;
    ctx.setLineDash([6, 4]);
    for (const r of st.pageBreaks.rows) {
      if (r < q.r1 || r > q.r2 + 1) continue;
      const y = L.rowY(r);
      ctx.beginPath();
      ctx.moveTo(q.x, y);
      ctx.lineTo(q.x + q.w, y);
      ctx.stroke();
    }
    for (const c of st.pageBreaks.cols) {
      if (c < q.c1 || c > q.c2 + 1) continue;
      const x = L.colX(c);
      ctx.beginPath();
      ctx.moveTo(x, q.y);
      ctx.lineTo(x, q.y + q.h);
      ctx.stroke();
    }
    ctx.setLineDash([]);
  }
  ctx.restore();
}

function hexAlpha(hex: string, a: number): string {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
}

/* ------------------------------------------------------------------ */
/* Başlıklar                                                            */
/* ------------------------------------------------------------------ */

function drawHeaders(ctx: CanvasRenderingContext2D, st: RenderState, quads: Quadrant[]): void {
  const { layout: L, dpr, sheet } = st;
  if (!sheet.showHeadings) return;
  const px = 1 / dpr;
  const z = L.z;
  const fontSize = Math.max(8, 11 * Math.max(0.65, Math.min(z, 2.2)));
  const font = `${fontSize}px "Segoe UI", system-ui, sans-serif`;
  const boldFont = `600 ${fontSize}px "Segoe UI", system-ui, sans-serif`;
  const sel = st.selection;
  const colSel = (c: number) => sel.ranges.some((r) => c >= r.c1 && c <= r.c2);
  const rowSel = (r: number) => sel.ranges.some((x) => r >= x.r1 && r <= x.r2);
  const colFull = (c: number) => sel.ranges.some((r) => c >= r.c1 && c <= r.c2 && r.r1 === 0 && r.r2 === MAX_ROWS - 1);
  const rowFull = (r: number) => sel.ranges.some((x) => r >= x.r1 && r <= x.r2 && x.c1 === 0 && x.c2 === MAX_COLS - 1);

  // Sütun başlıkları
  ctx.fillStyle = COLORS.headerBg;
  ctx.fillRect(0, 0, L.width, L.headerH);
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  const colQuads = quads.filter((q) => !q.frozenRows);
  for (const q of colQuads) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(q.x, 0, q.w, L.headerH);
    ctx.clip();
    for (let c = q.c1; c <= q.c2; c++) {
      const w = L.colW(c);
      if (w === 0) continue;
      const x = L.colX(c);
      const s = colSel(c);
      if (s) {
        ctx.fillStyle = colFull(c) ? COLORS.headerFullBg : COLORS.headerSelBg;
        ctx.fillRect(x, 0, w, L.headerH);
        ctx.fillStyle = COLORS.accent;
        ctx.fillRect(x, L.headerH - 2, w, 2);
      }
      ctx.fillStyle = COLORS.headerLine;
      ctx.fillRect(snap(x + w, dpr) - px, 0, px, L.headerH);
      ctx.font = s ? boldFont : font;
      ctx.fillStyle = s ? COLORS.headerSelText : COLORS.headerText;
      if (w > fontSize * 0.8) ctx.fillText(colToName(c), x + w / 2, L.headerH / 2 + 1);
    }
    ctx.restore();
  }
  ctx.fillStyle = COLORS.headerLine;
  ctx.fillRect(0, snap(L.headerH, dpr) - px, L.width, px);

  // Satır başlıkları
  ctx.fillStyle = COLORS.headerBg;
  ctx.fillRect(0, L.headerH, L.headerW, L.height - L.headerH);
  const rowQuads = quads.filter((q) => !q.frozenCols);
  for (const q of rowQuads) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, q.y, L.headerW, q.h);
    ctx.clip();
    for (let r = q.r1; r <= q.r2; r++) {
      const h = L.rowH(r);
      if (h === 0) continue;
      const y = L.rowY(r);
      const s = rowSel(r);
      if (s) {
        ctx.fillStyle = rowFull(r) ? COLORS.headerFullBg : COLORS.headerSelBg;
        ctx.fillRect(0, y, L.headerW, h);
        ctx.fillStyle = COLORS.accent;
        ctx.fillRect(L.headerW - 2, y, 2, h);
      }
      ctx.fillStyle = COLORS.headerLine;
      ctx.fillRect(0, snap(y + h, dpr) - px, L.headerW, px);
      ctx.font = s ? boldFont : font;
      ctx.fillStyle = s ? COLORS.headerSelText : COLORS.headerText;
      if (h > fontSize * 0.7) ctx.fillText(String(r + 1), L.headerW / 2, y + h / 2 + 1);
    }
    ctx.restore();
  }
  ctx.fillStyle = COLORS.headerLine;
  ctx.fillRect(snap(L.headerW, dpr) - px, 0, px, L.height);
  // Köşe (Tümünü Seç)
  ctx.fillStyle = COLORS.headerBg;
  ctx.fillRect(0, 0, L.headerW - px, L.headerH - px);
  ctx.fillStyle = "#B4B4B4";
  ctx.beginPath();
  const t = Math.min(L.headerW, L.headerH) * 0.55;
  ctx.moveTo(L.headerW - 3, L.headerH - 3 - t);
  ctx.lineTo(L.headerW - 3, L.headerH - 3);
  ctx.lineTo(L.headerW - 3 - t, L.headerH - 3);
  ctx.closePath();
  ctx.fill();
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
  void DEFAULT_FONT;
}
