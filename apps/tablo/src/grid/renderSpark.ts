/** Mini grafiklerin ızgaraya çizimi ve etkin grubun mavi çerçevesi. */
import type { CellPos } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { drawPrims, groupBounds, sparkAt, sparkPrims, sparkValues } from "../sparkline/geometry";
import type { GridLayout, Quadrant } from "./layout";

export function drawSparklines(ctx: CanvasRenderingContext2D, wb: WorkbookModel, sheet: SheetModel, L: GridLayout, q: Quadrant): void {
  if (q.w <= 0 || q.h <= 0) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.clip();
  for (const g of sheet.sparklines) {
    let bounds: { min: number; max: number } | null | undefined;
    for (const it of g.items) {
      if (it.row < q.r1 || it.row > q.r2 || it.col < q.c1 || it.col > q.c2) continue;
      if (L.rowH(it.row) === 0 || L.colW(it.col) === 0) continue;
      const m = sheet.mergeAt(it.row, it.col);
      const rc = m && m.r1 === it.row && m.c1 === it.col ? L.cellRect(m.r1, m.c1, m.r2, m.c2) : L.cellRect(it.row, it.col);
      if (bounds === undefined) bounds = groupBounds(wb, sheet, g);
      const prims = sparkPrims(g, sparkValues(wb, sheet, g, it), { x: rc.x, y: rc.y, w: rc.w - 1, h: rc.h - 1 }, L.z, bounds);
      if (!prims.length) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(rc.x, rc.y, rc.w - 1, rc.h - 1);
      ctx.clip();
      drawPrims(ctx, prims);
      ctx.restore();
    }
  }
  ctx.restore();
}

/** Etkin hücre bir mini grafikse grubunun hücreleri mavi çerçeveyle gösterilir. */
export function drawSparkGroupOutline(ctx: CanvasRenderingContext2D, sheet: SheetModel, L: GridLayout, active: CellPos): void {
  const s = sparkAt(sheet, active.row, active.col);
  if (!s) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(L.headerW, L.headerH, L.width - L.headerW, L.height - L.headerH);
  ctx.clip();
  ctx.strokeStyle = "#3C7FD9";
  ctx.lineWidth = 1.5;
  const rows = s.group.items.map((i) => i.row);
  const cols = s.group.items.map((i) => i.col);
  const r1 = Math.min(...rows);
  const r2 = Math.max(...rows);
  const c1 = Math.min(...cols);
  const c2 = Math.max(...cols);
  const contiguous = (r2 - r1 + 1) * (c2 - c1 + 1) === s.group.items.length;
  const boxes = contiguous ? [L.cellRect(r1, c1, r2, c2)] : s.group.items.map((i) => L.cellRect(i.row, i.col));
  for (const b of boxes) if (b.w > 0 && b.h > 0) ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 2, b.h - 2);
  ctx.restore();
}
