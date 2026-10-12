/** Noktaları Düzenle kipinde köşe tutamaçları (sürüklenerek taşınır). */
import { useRef } from "react";
import { moveVertex, normalizeBounds, vertices, type VertexRef } from "../model/editPoints";
import type { SpShape } from "../model/types";
import type { Actions } from "../state/actions";

export function PointEditor({ shape, zoom, actions }: { shape: SpShape; zoom: number; actions: Actions }) {
  const drag = useRef<{ ref: VertexRef; x: number; y: number; id: number } | null>(null);
  const z = zoom;
  const rad = (shape.rot * Math.PI) / 180;
  /** Ekran farkını şeklin kendi eksenine çevirir (döndürme ve çevirme hesaba katılır). */
  const local = (dx: number, dy: number): [number, number] => {
    const c = Math.cos(-rad);
    const s = Math.sin(-rad);
    let lx = (dx * c - dy * s) / z;
    let ly = (dx * s + dy * c) / z;
    if (shape.flipH) lx = -lx;
    if (shape.flipV) ly = -ly;
    return [lx, ly];
  };
  const update = (fn: (s: SpShape) => SpShape, merge: string) => actions.updateShapes((s) => (s.kind === "sp" && s.id === shape.id ? fn(s) : s), [shape.id], merge);
  const vs = vertices(shape);
  const flip = `${shape.flipH ? " scaleX(-1)" : ""}${shape.flipV ? " scaleY(-1)" : ""}`;
  return (
    <div
      className="point-editor"
      style={{ left: shape.x * z, top: shape.y * z, width: shape.w * z, height: shape.h * z, transform: `rotate(${shape.rot}deg)${flip}` }}
      aria-label="Noktaları düzenle"
    >
      <svg className="point-editor-path" width={shape.w * z} height={shape.h * z} overflow="visible" aria-hidden="true">
        {vs.map((v, i) => {
          const next = vs[i + 1];
          return next && next.path === v.path ? <line key={i} x1={v.x * z} y1={v.y * z} x2={next.x * z} y2={next.y * z} /> : null;
        })}
      </svg>
      {vs.map((v) => (
        <span
          key={`${v.path}:${v.cmd}`}
          className="point-handle"
          data-point={`${v.path}:${v.cmd}`}
          style={{ left: v.x * z - 4, top: v.y * z - 4 }}
          onPointerDown={(e) => {
            if (e.button !== 0) return;
            e.stopPropagation();
            e.preventDefault();
            (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
            drag.current = { ref: { path: v.path, cmd: v.cmd }, x: e.clientX, y: e.clientY, id: Date.now() };
          }}
          onPointerMove={(e) => {
            const d = drag.current;
            if (!d) return;
            const [lx, ly] = local(e.clientX - d.x, e.clientY - d.y);
            if (!lx && !ly) return;
            d.x = e.clientX;
            d.y = e.clientY;
            update((s) => moveVertex(s, d.ref, lx, ly), `points:${d.id}`);
          }}
          onPointerUp={() => {
            const d = drag.current;
            drag.current = null;
            if (d) update(normalizeBounds, `points:${d.id}`);
          }}
        />
      ))}
    </div>
  );
}
