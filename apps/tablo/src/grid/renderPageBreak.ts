/**
 * Sayfa Sonu Önizleme çizimi: yazdırılmayan alan gri, her sayfanın ortasında
 * soluk "Sayfa N" yazısı, yazdırma alanı kenarı ve el ile sayfa sonları kalın
 * düz mavi, otomatik sayfa sonları kesikli mavi. Sürüklenen sayfa sonu
 * önizlemesi de burada çizilir.
 */
import type { PageBreakInfo } from "../io/print";
import type { GridLayout, Quadrant } from "./layout";

export interface PageBreakDragPreview {
  axis: "row" | "col";
  /** Çizginin gideceği satır/sütun (o satırın/sütunun üst/sol kenarı). */
  at: number;
}

const BLUE = "#2F5BD3";

/** Satırın ya da sütunun ekrandaki kenarı; çeyrek dışındaysa çeyreğin biraz dışına sıkıştırılır. */
function edge(L: GridLayout, q: Quadrant, axis: "row" | "col", i: number): number {
  if (axis === "row") {
    if (i < q.r1) return q.y - 10;
    if (i > q.r2 + 1) return q.y + q.h + 10;
    return L.rowY(i);
  }
  if (i < q.c1) return q.x - 10;
  if (i > q.c2 + 1) return q.x + q.w + 10;
  return L.colX(i);
}

export function drawPageBreakPreview(ctx: CanvasRenderingContext2D, L: GridLayout, q: Quadrant, pb: PageBreakInfo, drag: PageBreakDragPreview | null | undefined, dark: boolean): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.clip();
  const gray = dark ? "rgba(70, 70, 70, 0.55)" : "rgba(150, 150, 150, 0.42)";
  const r = pb.range;
  if (!r) {
    ctx.fillStyle = gray;
    ctx.fillRect(q.x, q.y, q.w, q.h);
    ctx.restore();
    return;
  }
  const x1 = edge(L, q, "col", r.c1);
  const x2 = edge(L, q, "col", r.c2 + 1);
  const y1 = edge(L, q, "row", r.r1);
  const y2 = edge(L, q, "row", r.r2 + 1);
  // Yazdırılmayan alan
  ctx.fillStyle = gray;
  ctx.beginPath();
  ctx.rect(q.x, q.y, q.w, q.h);
  ctx.rect(x1, y1, Math.max(0, x2 - x1), Math.max(0, y2 - y1));
  ctx.fill("evenodd");
  // "Sayfa N" yazıları (önce aşağı, sonra yana numaralanır)
  const rowStarts = [r.r1, ...pb.rows.slice(0, -1)];
  const colStarts = [r.c1, ...pb.cols.slice(0, -1)];
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = dark ? "rgba(200, 200, 200, 0.28)" : "rgba(120, 120, 120, 0.32)";
  colStarts.forEach((c0, ci) => {
    const c1 = pb.cols[ci];
    if (c1 <= q.c1 || c0 > q.c2 + 1) return;
    rowStarts.forEach((r0, ri) => {
      const r1 = pb.rows[ri];
      if (r1 <= q.r1 || r0 > q.r2 + 1) return;
      const px = L.colX(c0);
      const py = L.rowY(r0);
      const w = L.colX(c1) - px;
      const h = L.rowY(r1) - py;
      if (w < 30 || h < 20) return;
      const size = Math.max(12, Math.min(72 * L.z, w / 5, h / 3));
      ctx.font = `600 ${size}px "Segoe UI", system-ui, sans-serif`;
      ctx.fillText(`Sayfa ${ci * rowStarts.length + ri + 1}`, px + w / 2, py + h / 2);
    });
  });
  // Sayfa sonları
  const line = (axis: "row" | "col", i: number, manual: boolean) => {
    ctx.lineWidth = manual ? 2.5 : 2;
    ctx.setLineDash(manual ? [] : [7, 4]);
    ctx.beginPath();
    if (axis === "row") {
      const y = L.rowY(i);
      ctx.moveTo(Math.max(q.x, x1), y);
      ctx.lineTo(Math.min(q.x + q.w, x2), y);
    } else {
      const x = L.colX(i);
      ctx.moveTo(x, Math.max(q.y, y1));
      ctx.lineTo(x, Math.min(q.y + q.h, y2));
    }
    ctx.stroke();
  };
  ctx.strokeStyle = BLUE;
  const mr = new Set(pb.manualRows);
  const mc = new Set(pb.manualCols);
  for (const row of pb.rows.slice(0, -1)) if (row >= q.r1 && row <= q.r2 + 1) line("row", row, mr.has(row));
  for (const col of pb.cols.slice(0, -1)) if (col >= q.c1 && col <= q.c2 + 1) line("col", col, mc.has(col));
  // Yazdırma alanı kenarı
  ctx.setLineDash([]);
  ctx.lineWidth = 3;
  ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
  // Sürüklenen sayfa sonu
  if (drag) {
    ctx.strokeStyle = dark ? "#9DB7FF" : "#1A3FA8";
    ctx.setLineDash([2, 2]);
    ctx.lineWidth = 2;
    ctx.beginPath();
    if (drag.axis === "row") {
      const y = L.rowY(drag.at);
      ctx.moveTo(q.x, y);
      ctx.lineTo(q.x + q.w, y);
    } else {
      const x = L.colX(drag.at);
      ctx.moveTo(x, q.y);
      ctx.lineTo(x, q.y + q.h);
    }
    ctx.stroke();
  }
  ctx.restore();
}
