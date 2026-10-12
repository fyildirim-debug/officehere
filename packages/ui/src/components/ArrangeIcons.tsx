/**
 * Yerleşim çizimleri (projeye özgü, MIT): Metni Kaydır, Konum sayfa önizlemesi, Öne Getir / Arkaya
 * Gönder ve Hizala. Saf SVG; renkler simge rengi (currentColor), vurgu rengi ve sayfa zemininden gelir,
 * koyu temada arrange.css'teki --oh-arrange-* değişkenleriyle açılır. Geometri: arrangeGeometry.ts.
 */
import type { CSSProperties } from "react";
import { alignDrawing, layerDrawing, positionDrawing, wrapDrawing, type AlignKind, type Drawing, type Ink, type LayerKind, type PositionKey, type Prim, type WrapKind } from "./arrangeGeometry";

export type { AlignKind, LayerKind, PositionKey, WrapKind } from "./arrangeGeometry";

const ACCENT = "var(--oh-arrange-accent, var(--oh-accent, #1f5fbf))";
const PAGE = "var(--oh-arrange-page, #ffffff)";

const PAINT: Record<Ink, CSSProperties> = {
  line: { fill: "currentColor", opacity: 0.5 },
  pic: { fill: ACCENT },
  ghost: { fill: ACCENT, opacity: 0.32 },
  mark: { fill: PAGE, opacity: 0.85 },
  page: { fill: PAGE },
  edge: { fill: "currentColor", opacity: 0.5 },
  ink: { fill: "currentColor" },
};

function shape(p: Prim, i: number) {
  const style = PAINT[p.ink];
  if (p.k === "rect") return <rect key={i} x={p.x} y={p.y} width={p.w} height={p.h} style={style} />;
  if (p.k === "frame") {
    const { x, y, w, h, sw } = p;
    const d = `M${x} ${y}h${w}v${h}h${-w}Z M${x + sw} ${y + sw}v${h - 2 * sw}h${w - 2 * sw}v${-(h - 2 * sw)}Z`;
    return <path key={i} d={d} fillRule="evenodd" style={style} />;
  }
  return <path key={i} d={p.d} style={style} />;
}

/** Çizimi SVG olarak verir (ox/oy: kutu içindeki tamsayı kaydırma). */
function DrawingSvg({ d, box, ox = 0, oy = 0 }: { d: Drawing; box?: { w: number; h: number }; ox?: number; oy?: number }) {
  const w = box?.w ?? d.w;
  const h = box?.h ?? d.h;
  const body = d.prims.map(shape);
  return (
    <svg className="oh-arrange-icon" width={w} height={h} viewBox={`0 0 ${w} ${h}`} aria-hidden="true" focusable="false">
      {ox || oy ? <g transform={`translate(${ox} ${oy})`}>{body}</g> : body}
    </svg>
  );
}

/** Metni Kaydır seçeneğinin çizimi (resim ile metin satırlarının akışı). */
export function WrapIcon({ mode, size = 16 }: { mode: WrapKind; size?: number }) {
  return <DrawingSvg d={wrapDrawing(mode, size)} />;
}

/** Katman sırası çizimi: vurgu renkli kart taşınan nesne. */
export function LayerIcon({ kind, size = 16 }: { kind: LayerKind; size?: number }) {
  return <DrawingSvg d={layerDrawing(kind, size)} />;
}

/** Hizala / Dağıt çizimi. */
export function AlignIcon({ kind, size = 16 }: { kind: AlignKind; size?: number }) {
  return <DrawingSvg d={alignDrawing(kind, size)} />;
}

/** Konum galerisi önizlemesi: sayfa, metin satırları ve resmin yeri. */
export function PositionPreview({ pos, width = 40, height = 52 }: { pos: PositionKey; width?: number; height?: number }) {
  return <DrawingSvg d={positionDrawing(pos, width, height)} />;
}

/** Konum düğmesi simgesi: kare kutuda küçük sayfa (resim solda ortada, metin yanından akar). */
export function PositionIcon({ size = 32 }: { size?: number }) {
  const ph = size - 2 * Math.max(1, Math.round(size / 32));
  let pw = Math.round(ph * 0.8);
  // Sayfa kutuda tam piksele ortalansın.
  if ((size - pw) % 2) pw -= 1;
  const d = positionDrawing({ v: "middle", h: "left" }, pw, ph);
  return <DrawingSvg d={d} box={{ w: size, h: size }} ox={(size - pw) / 2} oy={(size - ph) / 2} />;
}
