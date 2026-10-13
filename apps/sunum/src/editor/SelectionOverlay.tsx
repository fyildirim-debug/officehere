/** Seçim çerçeveleri, boyut/döndürme/ayar tutamaçları ve kılavuz çizgileri (ekran koordinatlarında). */
import type { GeomOverride } from "../render/SlideView";
import { isLineGeom, presetGeometry } from "../render/geometry";
import type { Shape } from "../model/types";
import { ADJUST_HANDLES, effectiveAdj } from "./adjust";
import { cursorFor, HANDLES, lineEnds, rotateVec } from "./interact";
import type { GuideLine } from "./snap";

export interface OverlayProps {
  shapes: Shape[];
  overrides?: Map<string, GeomOverride>;
  zoom: number;
  editingId?: string | null;
  /** Seçili alt şeklin üst grubu (kesik çizgili çerçeve). */
  parentGroup?: Shape | null;
  guides: GuideLine[];
  lasso: { x: number; y: number; w: number; h: number } | null;
  preview: { x: number; y: number; w: number; h: number; line?: boolean; flipH?: boolean; flipV?: boolean; geom?: string } | null;
  animBadges?: { shape: Shape; labels: string[] }[];
  cellSel?: { shape: Shape; rects: { x: number; y: number; w: number; h: number }[] } | null;
  rotInfo?: { x: number; y: number; text: string } | null;
  cropGhost?: { src: string; x: number; y: number; w: number; h: number } | null;
  comments?: { id: string; x: number; y: number; author: string; text: string }[];
  /** Kırpma modundaki resim (kırpma tutamaçlarıyla çizilir). */
  cropId?: string | null;
}

export function SelectionOverlay({ shapes, overrides, zoom: z, editingId, parentGroup, guides, lasso, preview, animBadges, cellSel, rotInfo, cropGhost, comments, cropId }: OverlayProps) {
  return (
    <div className="sel-layer">
      {cropGhost && <img className="crop-ghost" draggable={false} src={cropGhost.src} alt="" style={{ left: cropGhost.x * z, top: cropGhost.y * z, width: cropGhost.w * z, height: cropGhost.h * z }} />}
      {comments?.map((c, i) => (
        <div key={c.id} className="comment-marker" data-comment={c.id} title={`${c.author}: ${c.text}`} style={{ left: c.x * z, top: c.y * z }}>
          {i + 1}
        </div>
      ))}
      {animBadges?.map(({ shape, labels }) => {
        const g = overrides?.get(shape.id) ?? shape;
        return (
          <div key={`a-${shape.id}`} className="anim-badges" style={{ left: g.x * z - 4, top: g.y * z }}>
            {labels.map((l, i) => (
              <span key={i} className="anim-badge">
                {l}
              </span>
            ))}
          </div>
        );
      })}
      {parentGroup && (
        <div
          className="sel-frame is-group"
          style={{ left: parentGroup.x * z, top: parentGroup.y * z, width: parentGroup.w * z, height: parentGroup.h * z, transform: parentGroup.rot ? `rotate(${parentGroup.rot}deg)` : undefined }}
        />
      )}
      {shapes.map((s) => {
        const g = { ...s, ...(overrides?.get(s.id) ?? {}) };
        const line = s.kind === "sp" && isLineGeom(s.geom);
        if (line) {
          const [a, b] = lineEnds(g);
          return (
            <div key={s.id}>
              <div className="sel-line" />
              <div className="sel-handle is-end" data-handle="start" data-shape={s.id} style={{ left: a[0] * z, top: a[1] * z, cursor: "crosshair" }} />
              <div className="sel-handle is-end" data-handle="end" data-shape={s.id} style={{ left: b[0] * z, top: b[1] * z, cursor: "crosshair" }} />
            </div>
          );
        }
        const editing = editingId === s.id;
        const crop = cropId === s.id;
        const adjDefs = s.kind === "sp" && !crop ? ADJUST_HANDLES[s.geom] : undefined;
        const adj = s.kind === "sp" ? effectiveAdj(s.geom, s.adj) : {};
        const W = g.w * z;
        const H = g.h * z;
        return (
          <div
            key={s.id}
            className={`sel-frame${editing ? " is-editing" : ""}${crop ? " is-crop" : ""}`}
            style={{ left: g.x * z, top: g.y * z, width: W, height: H, transform: g.rot ? `rotate(${g.rot}deg)` : undefined }}
          >
            <div className="sel-move is-t" data-handle="move" data-shape={s.id} />
            <div className="sel-move is-b" data-handle="move" data-shape={s.id} />
            <div className="sel-move is-l" data-handle="move" data-shape={s.id} />
            <div className="sel-move is-r" data-handle="move" data-shape={s.id} />
            {s.kind !== "table" || !editing
              ? HANDLES.map((h) => (
                  <div
                    key={h.id}
                    className={`sel-handle${h.hx && h.hy ? " is-corner" : " is-edge"}`}
                    data-handle={h.id}
                    data-shape={s.id}
                    style={{ left: ((h.hx + 1) / 2) * W, top: ((h.hy + 1) / 2) * H, cursor: cursorFor(h.hx, h.hy, g.rot) }}
                  />
                ))
              : null}
            {s.kind !== "table" && !crop && (
              <div className="sel-rot" data-handle="rot" data-shape={s.id} style={{ left: W / 2 }}>
                <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true">
                  <path d="M12.5 5.5A5 5 0 1 0 13 9" fill="none" stroke="currentColor" strokeWidth="1.4" />
                  <path d="M13.5 2.5v3.5H10" fill="none" stroke="currentColor" strokeWidth="1.4" />
                </svg>
              </div>
            )}
            {adjDefs?.map((d, i) => {
              let [ax, ay] = d.pos(g.w, g.h, adj);
              if (g.flipH) ax = g.w - ax;
              if (g.flipV) ay = g.h - ay;
              return <div key={`adj${i}`} className="sel-adj" data-handle={`adj${i}`} data-shape={s.id} style={{ left: ax * z, top: ay * z }} />;
            })}
          </div>
        );
      })}
      {cellSel?.rects.map((r, i) => (
        <div key={`c${i}`} className="cell-sel" style={{ left: r.x * z, top: r.y * z, width: r.w * z, height: r.h * z }} />
      ))}
      {guides.map((g, i) =>
        g.gap ? (
          // Eşit aralık işareti: iki nesne arasındaki boşluk, uçları çentikli.
          <div
            key={`g${i}`}
            className={`smart-gap ${g.dir === "h" ? "is-h" : "is-v"}`}
            style={g.dir === "h" ? { top: g.pos * z, left: g.from * z, width: (g.to - g.from) * z } : { left: g.pos * z, top: g.from * z, height: (g.to - g.from) * z }}
          />
        ) : g.dir === "v" ? (
          <div key={`g${i}`} className="smart-guide is-v" style={{ left: g.pos * z, top: g.from * z, height: (g.to - g.from) * z }} />
        ) : (
          <div key={`g${i}`} className="smart-guide is-h" style={{ top: g.pos * z, left: g.from * z, width: (g.to - g.from) * z }} />
        ),
      )}
      {lasso && <div className="lasso" style={{ left: lasso.x * z, top: lasso.y * z, width: lasso.w * z, height: lasso.h * z }} />}
      {preview &&
        (preview.line ? (
          <svg className="insert-preview-line" style={{ left: preview.x * z, top: preview.y * z }} width={Math.max(1, preview.w * z)} height={Math.max(1, preview.h * z)} overflow="visible">
            <line
              x1={preview.flipH ? preview.w * z : 0}
              y1={preview.flipV ? preview.h * z : 0}
              x2={preview.flipH ? 0 : preview.w * z}
              y2={preview.flipV ? 0 : preview.h * z}
            />
          </svg>
        ) : preview.geom ? (
          <ShapePreview geom={preview.geom} x={preview.x * z} y={preview.y * z} w={preview.w * z} h={preview.h * z} />
        ) : (
          <div className="insert-preview" style={{ left: preview.x * z, top: preview.y * z, width: preview.w * z, height: preview.h * z }} />
        ))}
      {rotInfo && (
        <div className="rot-info" style={{ left: rotInfo.x, top: rotInfo.y }}>
          {rotInfo.text}
        </div>
      )}
    </div>
  );
}

export { rotateVec };

/** Çizilmekte olan şeklin önizlemesi: gerçek geometri, ince kenar ve yarı saydam dolgu. */
function ShapePreview({ geom, x, y, w, h }: { geom: string; x: number; y: number; w: number; h: number }) {
  if (w < 1 && h < 1) return null;
  const g = presetGeometry(geom, Math.max(1, w), Math.max(1, h));
  return (
    <svg className="insert-preview-shape" style={{ left: x, top: y }} width={Math.max(1, w)} height={Math.max(1, h)} overflow="visible" aria-hidden="true">
      {g.paths.map((p, i) => (
        <path key={i} d={p.d} fillRule={p.evenodd ? "evenodd" : undefined} className={g.open || p.fill === "none" ? "is-open" : undefined} />
      ))}
    </svg>
  );
}
