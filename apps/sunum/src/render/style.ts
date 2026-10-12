/** Dolgu, çizgi ve efektlerin CSS/SVG karşılıkları. */
import { cssColor, hexToRgb, resolveHex } from "../model/color";
import type { Color, DashType, Effects, Fill, Line, Theme } from "../model/types";
import type { FillMode } from "./geometry";

export const PT = 96 / 72;

export function modColor(c: Color, mode: FillMode): Color {
  switch (mode) {
    case "darken":
      return { ...c, shade: (c.shade ?? 1) * 0.6 };
    case "darkenLess":
      return { ...c, shade: (c.shade ?? 1) * 0.8 };
    case "lighten":
      return { ...c, tint: (c.tint ?? 1) * 0.6 };
    case "lightenLess":
      return { ...c, tint: (c.tint ?? 1) * 0.8 };
    default:
      return c;
  }
}

/** CSS arka plan değeri (slayt arka planı, HTML kapsayıcılar). */
export function fillToCss(fill: Fill | undefined, theme: Theme): string {
  if (!fill || fill.type === "none") return "transparent";
  if (fill.type === "solid") return cssColor(fill.color, theme);
  if (fill.type === "gradient") {
    const stops = fill.stops.map((s) => `${cssColor(s.color, theme)} ${(s.pos * 100).toFixed(1)}%`).join(", ");
    if (fill.kind === "radial") return `radial-gradient(circle at center, ${stops})`;
    // DrawingML açısı: 0 = soldan sağa, saat yönünde; CSS: 0 = aşağıdan yukarı.
    return `linear-gradient(${(fill.angle + 90) % 360}deg, ${stops})`;
  }
  return `center / ${fill.mode === "tile" ? "auto" : "100% 100%"} ${fill.mode === "tile" ? "repeat" : "no-repeat"} url("${fill.src}")`;
}

export interface SvgPaint {
  paint: string;
  def?: { kind: "linear" | "radial" | "pattern"; id: string; stops?: { offset: number; color: string }[]; angle?: number; src?: string };
}

export function svgPaint(fill: Fill, theme: Theme, id: string, mode: FillMode = "norm"): SvgPaint {
  if (fill.type === "none" || mode === "none") return { paint: "none" };
  if (fill.type === "solid") return { paint: cssColor(modColor(fill.color, mode), theme) };
  if (fill.type === "gradient") {
    const gid = `${id}-g${mode}`;
    return {
      paint: `url(#${gid})`,
      def: {
        kind: fill.kind,
        id: gid,
        angle: fill.angle,
        stops: fill.stops.map((s) => ({ offset: s.pos, color: cssColor(modColor(s.color, mode), theme) })),
      },
    };
  }
  const pid = `${id}-p`;
  return { paint: `url(#${pid})`, def: { kind: "pattern", id: pid, src: fill.src } };
}

export function dashArray(dash: DashType, widthPx: number): string | undefined {
  const w = Math.max(1, widthPx);
  const m = (arr: number[]) => arr.map((v) => (v * w).toFixed(2)).join(" ");
  switch (dash) {
    case "sysDot":
      return m([1, 1]);
    case "sysDash":
      return m([3, 1]);
    case "dash":
      return m([4, 3]);
    case "dashDot":
      return m([4, 3, 1, 3]);
    case "lgDash":
      return m([8, 3]);
    case "lgDashDot":
      return m([8, 3, 1, 3]);
    case "lgDashDotDot":
      return m([8, 3, 1, 3, 1, 3]);
    default:
      return undefined;
  }
}

export function lineStroke(line: Line, theme: Theme): { stroke: string; width: number; dash?: string } {
  if (line.fill.type === "none") return { stroke: "none", width: 0 };
  const width = Math.max(0.5, line.width * PT);
  return { stroke: cssColor(line.fill.color, theme), width, dash: dashArray(line.dash, width) };
}

/** Efektlerin CSS filter karşılığı (gölge, parlama). */
export function effectsFilter(fx: Effects | undefined, theme: Theme): string | undefined {
  if (!fx) return undefined;
  const parts: string[] = [];
  if (fx.shadow) {
    const r = (fx.shadow.dir * Math.PI) / 180;
    const dx = Math.cos(r) * fx.shadow.dist;
    const dy = Math.sin(r) * fx.shadow.dist;
    parts.push(`drop-shadow(${dx.toFixed(1)}px ${dy.toFixed(1)}px ${(fx.shadow.blur / 2).toFixed(1)}px ${cssColor(fx.shadow.color, theme)})`);
  }
  if (fx.glow) {
    const c = cssColor(fx.glow.color, theme);
    parts.push(`drop-shadow(0 0 ${(fx.glow.radius / 2).toFixed(1)}px ${c})`, `drop-shadow(0 0 ${(fx.glow.radius / 3).toFixed(1)}px ${c})`);
  }
  return parts.length ? parts.join(" ") : undefined;
}

export function reflectionCss(fx: Effects | undefined): string | undefined {
  if (!fx?.reflection) return undefined;
  return "below 2px linear-gradient(transparent 55%, rgba(255,255,255,0.35))";
}

/** Yazı rengi koyu zemin üzerinde mi (önizleme kontrastı için). */
export function isDarkFill(fill: Fill | undefined, theme: Theme): boolean {
  if (!fill || fill.type === "none") return false;
  const c = fill.type === "solid" ? fill.color : fill.type === "gradient" ? fill.stops[0]?.color : undefined;
  if (!c) return false;
  const [r, g, b] = hexToRgb(resolveHex(c, theme));
  return 0.299 * r + 0.587 * g + 0.114 * b < 128;
}
