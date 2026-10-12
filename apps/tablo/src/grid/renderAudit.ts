/** Formül denetimi çizimleri: izleme okları ve arka plan hata üçgenleri. */
import type { TraceArrow } from "../actions/audit";
import { cellError } from "../actions/audit";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import type { GridLayout, Quadrant } from "./layout";

const BLUE = "#2F6FD8";
const RED = "#E81123";
const GREEN = "#2AA048";

/** Görünen hücrelerde hata göstergesi (sol üst köşede yeşil üçgen). */
export function drawErrorTriangles(ctx: CanvasRenderingContext2D, wb: WorkbookModel, sheet: SheetModel, L: GridLayout, q: Quadrant): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.clip();
  ctx.fillStyle = GREEN;
  const s = Math.max(4, 6 * L.z);
  const ext = wb.dataExtent(sheet);
  const r2 = Math.min(q.r2, ext.rows - 1);
  const c2 = Math.min(q.c2, ext.cols - 1);
  for (let r = q.r1; r <= r2; r++) {
    if (L.rowH(r) === 0) continue;
    for (let c = q.c1; c <= c2; c++) {
      if (!wb.engine.hasInput(sheet.id, r, c)) continue;
      if (!cellError(wb, sheet, r, c)) continue;
      const rc = L.cellRect(r, c);
      if (rc.w === 0) continue;
      ctx.beginPath();
      ctx.moveTo(rc.x + 1, rc.y + 1);
      ctx.lineTo(rc.x + 1 + s, rc.y + 1);
      ctx.lineTo(rc.x + 1, rc.y + 1 + s);
      ctx.closePath();
      ctx.fill();
    }
  }
  ctx.restore();
}

function arrowHead(ctx: CanvasRenderingContext2D, x1: number, y1: number, x2: number, y2: number, size: number): void {
  const a = Math.atan2(y2 - y1, x2 - x1);
  ctx.beginPath();
  ctx.moveTo(x2, y2);
  ctx.lineTo(x2 - size * Math.cos(a - 0.4), y2 - size * Math.sin(a - 0.4));
  ctx.lineTo(x2 - size * Math.cos(a + 0.4), y2 - size * Math.sin(a + 0.4));
  ctx.closePath();
  ctx.fill();
}

/** Başka sayfa simgesi (dış başvurular). */
function sheetIcon(ctx: CanvasRenderingContext2D, x: number, y: number): void {
  ctx.fillStyle = "#FFFFFF";
  ctx.strokeStyle = "#333333";
  ctx.lineWidth = 1;
  ctx.fillRect(x - 9, y - 7, 18, 14);
  ctx.strokeRect(x - 9 + 0.5, y - 7 + 0.5, 17, 13);
  ctx.beginPath();
  for (let i = 1; i < 3; i++) {
    ctx.moveTo(x - 9, y - 7 + i * 4.5);
    ctx.lineTo(x + 9, y - 7 + i * 4.5);
  }
  ctx.moveTo(x - 3, y - 7);
  ctx.lineTo(x - 3, y + 7);
  ctx.moveTo(x + 3, y - 7);
  ctx.lineTo(x + 3, y + 7);
  ctx.stroke();
}

/** İzleme oklarını hücre alanı üzerine çizer. */
export function drawArrows(ctx: CanvasRenderingContext2D, arrows: TraceArrow[], sheet: SheetModel, L: GridLayout): void {
  const mine = arrows.filter((a) => a.sheetId === sheet.id);
  if (!mine.length) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(L.headerW, L.headerH, L.width - L.headerW, L.height - L.headerH);
  ctx.clip();
  const center = (row: number, col: number) => {
    const rc = L.cellRect(row, col);
    return { x: rc.x + rc.w / 2, y: rc.y + rc.h / 2, rc };
  };
  for (const a of mine) {
    const color = a.error ? RED : BLUE;
    ctx.lineWidth = 1.6;
    if (a.from && a.to) {
      const multi = a.from.r1 !== a.from.r2 || a.from.c1 !== a.from.c2;
      const s = center(a.from.r1, a.from.c1);
      if (multi) {
        // Aralık kutusu
        const rc = L.cellRect(a.from.r1, a.from.c1, a.from.r2, a.from.c2);
        ctx.strokeStyle = color;
        ctx.setLineDash([]);
        ctx.strokeRect(rc.x + 1, rc.y + 1, rc.w - 2, rc.h - 2);
      }
      const t = center(a.to.row, a.to.col);
      ctx.strokeStyle = color;
      ctx.fillStyle = color;
      ctx.setLineDash([]);
      ctx.beginPath();
      ctx.moveTo(s.x, s.y);
      ctx.lineTo(t.x, t.y);
      ctx.stroke();
      ctx.beginPath();
      ctx.arc(s.x, s.y, 3, 0, Math.PI * 2);
      ctx.fill();
      arrowHead(ctx, s.x, s.y, t.x, t.y, 9);
    } else if (a.to) {
      // Başka sayfadaki etkileyen: simge → hücre
      const t = center(a.to.row, a.to.col);
      const ix = t.x - 46;
      const iy = t.y - 28;
      ctx.strokeStyle = "#333333";
      ctx.fillStyle = "#333333";
      ctx.setLineDash([4, 3]);
      ctx.beginPath();
      ctx.moveTo(ix, iy);
      ctx.lineTo(t.x, t.y);
      ctx.stroke();
      ctx.setLineDash([]);
      arrowHead(ctx, ix, iy, t.x, t.y, 9);
      sheetIcon(ctx, ix, iy);
    } else if (a.from) {
      // Başka sayfadaki etkilenen: hücre → simge
      const s = center(a.from.r1, a.from.c1);
      const ix = s.x + 46;
      const iy = s.y - 28;
      ctx.strokeStyle = "#333333";
      ctx.fillStyle = "#333333";
      ctx.setLineDash([4, 3]);
      ctx.beginPath();
      ctx.moveTo(s.x, s.y);
      ctx.lineTo(ix, iy);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.beginPath();
      ctx.arc(s.x, s.y, 3, 0, Math.PI * 2);
      ctx.fill();
      arrowHead(ctx, s.x, s.y, ix, iy, 9);
      sheetIcon(ctx, ix, iy);
    }
  }
  ctx.restore();
}
