/** Mürekkep (serbest çizim) şekli: Çizim sekmesi ve gösteride tutulan ek açıklamalar. */
import { makeSp, NO_FILL } from "./factory";
import type { PathCmd, SpShape } from "./types";

export interface InkStyle {
  /** "#RRGGBB" */
  color: string;
  /** pt */
  width: number;
  /** 0..1 (vurgulayıcı yarı saydam). */
  alpha?: number;
}

/** Slayt koordinatlarındaki noktalardan yumuşatılmış mürekkep şekli üretir. */
export function inkShape(points: [number, number][], style: InkStyle, name: string): SpShape | null {
  if (!points.length) return null;
  const pad = 0.5;
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const [x, y] of points) {
    x0 = Math.min(x0, x);
    y0 = Math.min(y0, y);
    x1 = Math.max(x1, x);
    y1 = Math.max(y1, y);
  }
  x0 -= pad;
  y0 -= pad;
  const w = Math.max(1, x1 + pad - x0);
  const h = Math.max(1, y1 + pad - y0);
  const loc = points.map(([x, y]) => [x - x0, y - y0] as [number, number]);
  // Noktalar arası orta noktalardan geçen ikinci derece eğrilerle yumuşatılır.
  const cmds: PathCmd[] = [{ c: "M", pts: [loc[0][0], loc[0][1]] }];
  if (loc.length === 1) cmds.push({ c: "L", pts: [loc[0][0] + 0.01, loc[0][1] + 0.01] });
  else if (loc.length === 2) cmds.push({ c: "L", pts: [loc[1][0], loc[1][1]] });
  else {
    for (let k = 1; k < loc.length - 1; k++) {
      const [cx, cy] = loc[k];
      const [nx, ny] = loc[k + 1];
      cmds.push({ c: "Q", pts: [cx, cy, (cx + nx) / 2, (cy + ny) / 2] });
    }
    const last = loc[loc.length - 1];
    cmds.push({ c: "L", pts: [last[0], last[1]] });
  }
  const sh = makeSp("custom", x0, y0, w, h, {
    name,
    fill: NO_FILL,
    line: { fill: { type: "solid", color: { rgb: style.color, alpha: style.alpha } }, width: style.width, dash: "solid", cap: "rnd" },
    paths: [{ w, h, cmds, fill: false }],
  }) as SpShape;
  sh.ink = true;
  return sh;
}

/** SVG yol verisi (gösterideki canlı çizim için). */
export function inkPathD(points: [number, number][]): string {
  if (!points.length) return "";
  const [x0, y0] = points[0];
  if (points.length === 1) return `M${x0} ${y0} L${x0 + 0.01} ${y0 + 0.01}`;
  let d = `M${x0.toFixed(1)} ${y0.toFixed(1)}`;
  if (points.length === 2) return `${d} L${points[1][0].toFixed(1)} ${points[1][1].toFixed(1)}`;
  for (let k = 1; k < points.length - 1; k++) {
    const [cx, cy] = points[k];
    const [nx, ny] = points[k + 1];
    d += ` Q${cx.toFixed(1)} ${cy.toFixed(1)} ${((cx + nx) / 2).toFixed(1)} ${((cy + ny) / 2).toFixed(1)}`;
  }
  const last = points[points.length - 1];
  return `${d} L${last[0].toFixed(1)} ${last[1].toFixed(1)}`;
}
