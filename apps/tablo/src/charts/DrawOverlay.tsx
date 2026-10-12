/** Çizim kipi (Ekle > Metin Kutusu): ızgara üzerinde sürükleyerek metin kutusu çizilir; tıklama varsayılan boyutta ekler. */
import { useRef, useState } from "react";
import type { Controller } from "../controller";
import type { GridLayout } from "../grid/layout";
import { anchorAt } from "./FloatingLayer";
import { insertShapeAt } from "../actions/objects";

/** Izgara (canvas) koordinatı → sayfa koordinatı (piksel, yakınlaştırmasız). */
function toSheet(L: GridLayout, x: number, y: number): { x: number; y: number } {
  const sx = x < L.cellsX && L.fc > 0 ? (x - L.headerW) / L.z : L.sheet.cols.offset(L.fc) + (Math.max(x, L.cellsX) - L.cellsX + L.scrollX) / L.z;
  const sy = y < L.cellsY && L.fr > 0 ? (y - L.headerH) / L.z : L.sheet.rows.offset(L.fr) + (Math.max(y, L.cellsY) - L.cellsY + L.scrollY) / L.z;
  return { x: Math.max(0, sx), y: Math.max(0, sy) };
}

export function DrawOverlay({ ctl, layout }: { ctl: Controller; layout: GridLayout | null }) {
  const [box, setBox] = useState<{ x1: number; y1: number; x2: number; y2: number } | null>(null);
  const origin = useRef<{ x: number; y: number } | null>(null);
  if (!ctl.drawShape || !layout) return null;
  const local = (e: React.PointerEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };
  return (
    <div
      className="tb-draw-overlay"
      style={{ left: 0, top: 0, width: layout.width, height: layout.height }}
      onPointerDown={(e) => {
        if (e.button !== 0) return;
        e.stopPropagation();
        e.preventDefault();
        const p = local(e);
        origin.current = p;
        setBox({ x1: p.x, y1: p.y, x2: p.x, y2: p.y });
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
      }}
      onPointerMove={(e) => {
        if (!origin.current) return;
        const p = local(e);
        setBox({ x1: origin.current.x, y1: origin.current.y, x2: p.x, y2: p.y });
      }}
      onPointerUp={(e) => {
        const o = origin.current;
        origin.current = null;
        setBox(null);
        if (!o) return;
        const p = local(e);
        const a = toSheet(layout, Math.min(o.x, p.x), Math.min(o.y, p.y));
        const b = toSheet(layout, Math.max(o.x, p.x), Math.max(o.y, p.y));
        const dragged = Math.abs(p.x - o.x) > 4 || Math.abs(p.y - o.y) > 4;
        const kind = ctl.drawShape!;
        ctl.drawShape = null;
        ctl.status = "";
        insertShapeAt(ctl, kind, anchorAt(ctl.sheet, a.x, a.y), dragged ? Math.max(30, b.x - a.x) : 140, dragged ? Math.max(20, b.y - a.y) : 60);
      }}
    >
      {box && <div className="tb-draw-box" style={{ left: Math.min(box.x1, box.x2), top: Math.min(box.y1, box.y2), width: Math.abs(box.x2 - box.x1), height: Math.abs(box.y2 - box.y1) }} />}
    </div>
  );
}
