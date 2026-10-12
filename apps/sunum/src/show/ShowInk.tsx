/**
 * Gösteri sırasında işaretçi araçları: Lazer İşaretçisi, Kalem, Vurgulayıcı,
 * Silgi. Mürekkep slayt koordinatlarında tutulur; gösteri bitince tutulursa
 * slayta mürekkep şekli olarak eklenir.
 */
import { useRef, useState } from "react";
import { inkPathD } from "../model/ink";

export type ShowTool = "pointer" | "laser" | "pen" | "highlighter" | "eraser";

export interface ShowStroke {
  tool: "pen" | "highlighter";
  /** "#RRGGBB" */
  color: string;
  /** pt */
  width: number;
  pts: [number, number][];
}

export const SHOW_INK_COLORS: { name: string; hex: string }[] = [
  { name: "Kırmızı", hex: "#E81123" },
  { name: "Siyah", hex: "#000000" },
  { name: "Beyaz", hex: "#FFFFFF" },
  { name: "Mavi", hex: "#0078D4" },
  { name: "Yeşil", hex: "#107C10" },
  { name: "Sarı", hex: "#FFE600" },
  { name: "Turuncu", hex: "#FF8C00" },
  { name: "Mor", hex: "#8764B8" },
];

export const PEN_WIDTH = 2.25;
export const HIGHLIGHTER_WIDTH = 12;
const PT = 96 / 72;

function distToSeg(px: number, py: number, ax: number, ay: number, bx: number, by: number): number {
  const dx = bx - ax;
  const dy = by - ay;
  const l2 = dx * dx + dy * dy;
  const t = l2 ? Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / l2)) : 0;
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

/** Noktanın üzerinden geçtiği çizgi dizinleri (silgi için). */
export function strokesAt(strokes: ShowStroke[], x: number, y: number, tol = 8): number[] {
  const out: number[] = [];
  strokes.forEach((s, i) => {
    const r = tol + (s.width * PT) / 2;
    for (let k = 0; k < s.pts.length; k++) {
      const a = s.pts[k];
      const b = s.pts[Math.min(k + 1, s.pts.length - 1)];
      if (distToSeg(x, y, a[0], a[1], b[0], b[1]) <= r) {
        out.push(i);
        return;
      }
    }
  });
  return out;
}

export interface InkOverlayProps {
  width: number;
  height: number;
  tool: ShowTool;
  color: string;
  strokes: ShowStroke[];
  visible: boolean;
  onAdd: (s: ShowStroke) => void;
  onErase: (indices: number[]) => void;
}

/** Slaytın üstündeki çizim katmanı (slayt koordinatlarında SVG). */
export function InkOverlay({ width, height, tool, color, strokes, visible, onAdd, onErase }: InkOverlayProps) {
  const ref = useRef<SVGSVGElement>(null);
  const [live, setLive] = useState<[number, number][] | null>(null);
  const [laser, setLaser] = useState<[number, number] | null>(null);
  const drawing = useRef(false);
  const active = tool !== "pointer";

  const toSlide = (e: React.PointerEvent): [number, number] => {
    const r = ref.current!.getBoundingClientRect();
    return [((e.clientX - r.left) / r.width) * width, ((e.clientY - r.top) / r.height) * height];
  };

  const onDown = (e: React.PointerEvent) => {
    if (e.button !== 0 || tool === "pointer" || tool === "laser") return;
    e.stopPropagation();
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
    drawing.current = true;
    const p = toSlide(e);
    if (tool === "eraser") {
      onErase(strokesAt(strokes, p[0], p[1]));
      return;
    }
    setLive([p]);
  };
  const onMove = (e: React.PointerEvent) => {
    const p = toSlide(e);
    if (tool === "laser") setLaser(p);
    if (!drawing.current) return;
    if (tool === "eraser") {
      const hit = strokesAt(strokes, p[0], p[1]);
      if (hit.length) onErase(hit);
      return;
    }
    setLive((l) => (l ? [...l, p] : l));
  };
  const onUp = () => {
    if (!drawing.current) return;
    drawing.current = false;
    if (live && live.length && (tool === "pen" || tool === "highlighter")) {
      onAdd({ tool, color, width: tool === "pen" ? PEN_WIDTH : HIGHLIGHTER_WIDTH, pts: live });
    }
    setLive(null);
  };

  const strokeEl = (s: { tool: string; color: string; width: number; pts: [number, number][] }, key: string | number) => (
    <path
      key={key}
      d={inkPathD(s.pts)}
      fill="none"
      stroke={s.color}
      strokeOpacity={s.tool === "highlighter" ? 0.5 : 1}
      strokeWidth={s.width * PT}
      strokeLinecap={s.tool === "highlighter" ? "square" : "round"}
      strokeLinejoin="round"
    />
  );

  return (
    <svg
      ref={ref}
      className={`show-ink is-${tool}${active ? " is-active" : ""}`}
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      onPointerDown={onDown}
      onPointerMove={onMove}
      onPointerUp={onUp}
      onPointerCancel={onUp}
      onPointerLeave={() => setLaser(null)}
      data-testid="show-ink"
    >
      {visible && strokes.map((s, i) => strokeEl(s, i))}
      {live && strokeEl({ tool, color, width: tool === "pen" ? PEN_WIDTH : HIGHLIGHTER_WIDTH, pts: live }, "live")}
      {tool === "laser" && laser && (
        <g className="show-laser" pointerEvents="none">
          <circle cx={laser[0]} cy={laser[1]} r={14} fill="rgba(255,0,0,0.25)" />
          <circle cx={laser[0]} cy={laser[1]} r={7} fill="#ff1a1a" stroke="#fff" strokeOpacity={0.6} strokeWidth={1.5} />
        </g>
      )}
    </svg>
  );
}
