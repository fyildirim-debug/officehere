/** Animasyonlar sekmesinde hareket yolları: kesik çizgi, yeşil başlangıç ve kırmızı bitiş işareti. */
import { findShape } from "../model/ops";
import type { Slide } from "../model/types";
import { motionPathSvg } from "../show/motionPaths";

export function MotionPathOverlay({ slide, z, W, H, selected }: { slide: Slide; z: number; W: number; H: number; selected: string[] }) {
  const paths = slide.animations.filter((a) => a.cls === "path");
  if (!paths.length) return null;
  return (
    <svg className="motion-paths" width={W * z} height={H * z} viewBox={`0 0 ${W} ${H}`} aria-hidden="true">
      {paths.map((a) => {
        const s = findShape(slide.shapes, a.shapeId);
        if (!s) return null;
        const { d, end } = motionPathSvg(a, s.x + s.w / 2, s.y + s.h / 2, W, H);
        const on = selected.includes(s.id);
        return (
          <g key={a.id} className={on ? "is-selected" : undefined}>
            <path d={d} fill="none" stroke="#7a7a7a" strokeWidth={on ? 2 / z : 1.2 / z} strokeDasharray={`${4 / z} ${3 / z}`} />
            <circle cx={s.x + s.w / 2} cy={s.y + s.h / 2} r={4 / z} fill="#3a9d23" />
            <path d={`M${end[0] - 4 / z} ${end[1] - 4 / z} L${end[0] + 4 / z} ${end[1] + 4 / z} M${end[0] + 4 / z} ${end[1] - 4 / z} L${end[0] - 4 / z} ${end[1] + 4 / z}`} stroke="#c8102e" strokeWidth={2 / z} />
          </g>
        );
      })}
    </svg>
  );
}
