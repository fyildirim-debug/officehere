/**
 * Canvas ızgara çizici: yalnız görünen hücreleri çizer. Sıra: zemin,
 * kılavuz çizgileri, dolgular, taşan metin ve birleştirme zeminleri,
 * kenarlıklar, metin, işaretler; ardından seçim katmanları ve başlıklar.
 */
import { colToName, keyRow, keyCol } from "@officehere/formula";
import type { CondFormatter } from "../model/condfmt";
import { computeView, type CellView } from "../model/display";
import { contains, intersects, MAX_COLS, MAX_ROWS } from "../model/rect";
import type { Rect, Selection } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import type { GridLayout, Quadrant } from "./layout";
import { cellFont, drawBorder, drawCellText, drawFilterButton, drawPattern, measureText, snap } from "./renderCell";
import { drawArrows, drawErrorTriangles } from "./renderAudit";
import { drawIcon } from "./icons";
import { drawSparkGroupOutline, drawSparklines } from "./renderSpark";
import type { TraceArrow } from "../actions/audit";

export { cellFont, measureText, wrapLines } from "./renderCell";

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
  /** Formül Denetimi okları. */
  arrows?: TraceArrow[];
  /** Arka planda hata denetimi (yeşil üçgenler). */
  errorCheck?: boolean;
}

const LIGHT = {
  grid: "#D4D4D4",
  headerBg: "#F5F5F5",
  headerText: "#3B3B3B",
  headerLine: "#D0D0D0",
  headerSelBg: "#D3E6DA",
  headerSelText: "#0F6B39",
  headerFullBg: "#A8D2B8",
  accent: "#107C41",
  selFill: "rgba(20, 110, 60, 0.12)",
  corner: "#B4B4B4",
  freeze: "#9E9E9E",
};

/** Koyu arayüz temalarında başlıklar koyu, hücre alanı beyaz kalır (Excel gibi). */
const DARK: typeof LIGHT = {
  ...LIGHT,
  headerBg: "#2B2B2B",
  headerText: "#D2D2D2",
  headerLine: "#464646",
  headerSelBg: "#38503F",
  headerSelText: "#8FD8A8",
  headerFullBg: "#2F6B47",
  corner: "#6E6E6E",
  freeze: "#7A7A7A",
};

let COLORS = LIGHT;

export function renderGrid(ctx: CanvasRenderingContext2D, st: RenderState): void {
  const { layout: L, dpr } = st;
  COLORS = st.dark ? DARK : LIGHT;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, L.width, L.height);
  ctx.textBaseline = "alphabetic";
  const quads = L.quadrants();
  for (const q of quads) {
    drawQuadrant(ctx, st, q);
    if (st.sheet.sparklines.length) drawSparklines(ctx, st.wb, st.sheet, L, q);
    if (st.errorCheck && !st.sheet.showFormulas) drawErrorTriangles(ctx, st.wb, st.sheet, L, q);
  }
  for (const q of quads) drawOverlays(ctx, st, q);
  if (st.sheet.sparklines.length && st.showSelection) drawSparkGroupOutline(ctx, st.sheet, L, st.selection.active);
  if (st.arrows?.length) drawArrows(ctx, st.arrows, st.sheet, L);
  drawHeaders(ctx, st, quads);
  // Dondurulmuş bölme çizgileri
  ctx.fillStyle = COLORS.freeze;
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

  // Metinler (simge kümesi varsa simge solda, metin kalan alanda)
  for (const sp of spans) {
    if (st.editing && st.editing.row === sp.r && st.editing.col === sp.c) continue;
    const y = L.rowY(sp.r);
    const h = L.rowH(sp.r);
    const iw = sp.view.icon ? iconBox(h, z) + 3 * z : 0;
    drawCellText(ctx, sp.view, sp.x + (sp.x === L.colX(sp.c) ? iw : 0), y, sp.w - (sp.x === L.colX(sp.c) ? iw : 0), h, L.colX(sp.c) + iw, L.colW(sp.c) - iw, z, dpr);
  }
  // Simge kümesi simgeleri
  for (let i = 0; i < rows.length; i++) {
    const r = q.r1 + i;
    const rv = rows[i];
    if (!rv.views.length) continue;
    for (let c = q.c1; c <= q.c2; c++) {
      const v = rv.views[c - c0];
      if (!v?.icon || L.colW(c) === 0) continue;
      const y = L.rowY(r);
      const h = L.rowH(r);
      const s = Math.min(iconBox(h, z), L.colW(c) - 2 * z);
      if (s < 4) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(L.colX(c), y, L.colW(c), h);
      ctx.clip();
      drawIcon(ctx, v.icon.set, v.icon.index, L.colX(c) + 2 * z, y + (h - s) / 2, s);
      ctx.restore();
    }
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

/** Simge kutusu boyutu (satır yüksekliğine göre). */
function iconBox(h: number, z: number): number {
  return Math.max(6, Math.min(h - 4 * z, 15 * z));
}

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
  ctx.fillStyle = COLORS.corner;
  ctx.beginPath();
  const t = Math.min(L.headerW, L.headerH) * 0.55;
  ctx.moveTo(L.headerW - 3, L.headerH - 3 - t);
  ctx.lineTo(L.headerW - 3, L.headerH - 3);
  ctx.lineTo(L.headerW - 3 - t, L.headerH - 3);
  ctx.closePath();
  ctx.fill();
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
}
