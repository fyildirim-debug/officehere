/** Noktaları Düzenle kipinde köşe tutamaçları (sürüklenerek taşınır). */
import { useRef } from "react";
import { Menu, MenuItem, MenuSeparator, useContextMenu } from "@officehere/ui";
import { addVertexAfter, deleteVertex, isPathClosed, moveVertex, normalizeBounds, setSegmentKind, setVertexKind, togglePathClosed, vertices, type VertexRef } from "../model/editPoints";
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
  const ctx = useContextMenu();
  const op = (fn: (s: SpShape) => SpShape) => update((s) => normalizeBounds(fn(s)), `pointop:${Date.now()}`);
  /** Nokta ya da kesim üzerinde sağ tık menüsü. */
  const menu = (e: React.MouseEvent, ref: VertexRef, onVertex: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    const closed = isPathClosed(shape, ref.path);
    ctx.open(
      e,
      <Menu>
        <MenuItem label="Nokta Ekle" onClick={() => op((s) => addVertexAfter(s, ref))} />
        {onVertex && <MenuItem label="Noktayı Sil" onClick={() => op((s) => deleteVertex(s, ref))} />}
        <MenuItem label={closed ? "Yolu Aç" : "Yolu Kapat"} onClick={() => op((s) => togglePathClosed(s, ref.path))} />
        <MenuSeparator />
        {onVertex ? (
          <>
            <MenuItem label="Düzgün Nokta" onClick={() => op((s) => setVertexKind(s, ref, "smooth"))} />
            <MenuItem label="Düz Nokta" onClick={() => op((s) => setVertexKind(s, ref, "straight"))} />
            <MenuItem label="Köşe Noktası" onClick={() => op((s) => setVertexKind(s, ref, "corner"))} />
          </>
        ) : (
          <>
            <MenuItem label="Düz Kesim" onClick={() => op((s) => setSegmentKind(s, { path: ref.path, cmd: ref.cmd + 1 }, "straight"))} />
            <MenuItem label="Eğri Kesim" onClick={() => op((s) => setSegmentKind(s, { path: ref.path, cmd: ref.cmd + 1 }, "curved"))} />
          </>
        )}
        <MenuSeparator />
        <MenuItem label="Noktaları Düzenlemekten Çık" onClick={() => actions.store.set({ pointEdit: null })} />
      </Menu>,
    );
  };
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
          if (!next || next.path !== v.path) return null;
          return (
            <g key={i} data-seg={`${v.path}:${v.cmd}`} onContextMenu={(e) => menu(e, v, false)} onPointerDown={(e) => e.button === 2 && e.stopPropagation()}>
              <line className="seg-hit" x1={v.x * z} y1={v.y * z} x2={next.x * z} y2={next.y * z} />
              <line x1={v.x * z} y1={v.y * z} x2={next.x * z} y2={next.y * z} />
            </g>
          );
        })}
      </svg>
      {vs.map((v) => (
        <span
          key={`${v.path}:${v.cmd}`}
          className="point-handle"
          data-point={`${v.path}:${v.cmd}`}
          style={{ left: v.x * z - 4, top: v.y * z - 4 }}
          onContextMenu={(e) => menu(e, v, true)}
          onPointerDown={(e) => {
            if (e.button !== 0) {
              e.stopPropagation();
              return;
            }
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
      {/* Menü portalda çizilir ama React olayları ağaçta kabarcıklanır: slayt alanına ulaşmasın. */}
      <div onPointerDown={(e) => e.stopPropagation()} onPointerUp={(e) => e.stopPropagation()} onDoubleClick={(e) => e.stopPropagation()} onContextMenu={(e) => e.stopPropagation()}>
        {ctx.element}
      </div>
    </div>
  );
}
