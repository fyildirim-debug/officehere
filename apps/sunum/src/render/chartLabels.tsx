/** Grafik veri etiketleri: konumlar (Orta, İç Uç, İç Taban, Dış Uç, Sol/Sağ/Üst/Alt) ve veri belirtme çizgisi. */
import type { ReactNode } from "react";
import type { DataLabelPos } from "../model/types";
import type { ChartPaint } from "./ChartSvg";

const textW = (s: string, fs: number) => s.length * fs * 0.55;

/** Kutulu etiket ve çubuğa/noktaya uzanan çizgi (veri belirtme çizgisi). */
function callout(key: string, x: number, y: number, text: string, ax: number, ay: number, paint: ChartPaint, fs: number): ReactNode {
  const w = textW(text, fs) + fs * 0.8;
  const h = fs * 1.5;
  const bx = x - w / 2;
  const by = y - h / 2;
  // Çizgi kutunun en yakın kenar noktasına bağlanır.
  const ex = Math.max(bx, Math.min(bx + w, ax));
  const ey = Math.max(by, Math.min(by + h, ay));
  return (
    <g key={key} className="chart-callout">
      <line x1={ax} y1={ay} x2={ex} y2={ey} stroke={paint.text} strokeWidth={0.75} />
      <rect x={bx} y={by} width={w} height={h} rx={fs * 0.2} fill={paint.labelBg} stroke={paint.text} strokeWidth={0.75} />
      <text x={x} y={y + fs * 0.33} textAnchor="middle" fontSize={fs} fill={paint.text}>
        {text}
      </text>
    </g>
  );
}

function plain(key: string, x: number, y: number, text: string, anchor: "start" | "middle" | "end", fill: string, fs: number): ReactNode {
  return (
    <text key={key} x={x} y={y} textAnchor={anchor} fontSize={fs} fill={fill}>
      {text}
    </text>
  );
}

/** Sütun/çubuk etiketi. `r`: çubuk dikdörtgeni; `neg`: değer negatif (uç tarafı ters). */
export function barLabel(key: string, r: { x: number; y: number; w: number; h: number }, horiz: boolean, neg: boolean, pos: DataLabelPos, text: string, paint: ChartPaint): ReactNode {
  const fs = paint.fs * 0.9;
  const inside = "#fff";
  if (!horiz) {
    const cx = r.x + r.w / 2;
    const top = neg ? r.y + r.h : r.y;
    const base = neg ? r.y : r.y + r.h;
    const dir = neg ? -1 : 1; // uçtan tabana doğru yön (ekranda +y aşağı)
    switch (pos) {
      case "ctr":
        return plain(key, cx, r.y + r.h / 2 + fs * 0.35, text, "middle", inside, fs);
      case "inEnd":
        return plain(key, cx, top + dir * fs * 1.05 + (neg ? fs * 0.35 : 0), text, "middle", inside, fs);
      case "inBase":
        return plain(key, cx, base - dir * fs * 0.4 + (neg ? fs * 0.7 : 0), text, "middle", inside, fs);
      case "callout":
        return callout(key, cx + r.w * 0.4, top - dir * fs * 1.6, text, cx, top, paint, fs);
      default:
        return plain(key, cx, neg ? top + fs * 1.1 : top - fs * 0.3, text, "middle", paint.text, fs);
    }
  }
  const cy = r.y + r.h / 2 + fs * 0.35;
  const end = neg ? r.x : r.x + r.w;
  const base = neg ? r.x + r.w : r.x;
  switch (pos) {
    case "ctr":
      return plain(key, r.x + r.w / 2, cy, text, "middle", inside, fs);
    case "inEnd":
      return plain(key, end + (neg ? fs * 0.3 : -fs * 0.3), cy, text, neg ? "start" : "end", inside, fs);
    case "inBase":
      return plain(key, base + (neg ? -fs * 0.3 : fs * 0.3), cy, text, neg ? "end" : "start", inside, fs);
    case "callout":
      return callout(key, end + (neg ? -1 : 1) * (textW(text, fs) / 2 + fs * 1.4), r.y + r.h / 2 - fs * 0.6, text, end, r.y + r.h / 2, paint, fs);
    default:
      return plain(key, end + (neg ? -fs * 0.3 : fs * 0.3), cy, text, neg ? "end" : "start", paint.text, fs);
  }
}

/** Nokta (çizgi, dağılım, radar) etiketi. */
export function pointLabel(key: string, x: number, y: number, pos: DataLabelPos, text: string, paint: ChartPaint): ReactNode {
  const fs = paint.fs * 0.9;
  switch (pos) {
    case "ctr":
      return plain(key, x, y + fs * 0.35, text, "middle", paint.text, fs);
    case "l":
      return plain(key, x - fs * 0.6, y + fs * 0.35, text, "end", paint.text, fs);
    case "r":
      return plain(key, x + fs * 0.6, y + fs * 0.35, text, "start", paint.text, fs);
    case "b":
      return plain(key, x, y + fs * 1.4, text, "middle", paint.text, fs);
    case "callout":
      return callout(key, x + textW(text, fs) / 2 + fs, y - fs * 1.6, text, x, y, paint, fs);
    default:
      return plain(key, x, y - fs * 0.6, text, "middle", paint.text, fs);
  }
}

/** Pasta dilimi etiketi; `mid`: dilim orta açısı, `r`/`hole`: dış ve iç yarıçap. */
export function sliceLabel(key: string, cx: number, cy: number, r: number, hole: number, mid: number, pos: DataLabelPos, text: string, paint: ChartPaint): ReactNode {
  const fs = paint.fs;
  const at = (rr: number): [number, number] => [cx + rr * Math.cos(mid), cy + rr * Math.sin(mid)];
  const right = Math.cos(mid) >= 0;
  switch (pos) {
    case "outEnd": {
      const [x, y] = at(r + fs * 0.5);
      return plain(key, x, y + fs * 0.35, text, Math.abs(Math.cos(mid)) < 0.2 ? "middle" : right ? "start" : "end", paint.text, fs);
    }
    case "inEnd": {
      const [x, y] = at(r - fs * 1.4);
      return plain(key, x, y + fs * 0.35, text, "middle", "#fff", fs);
    }
    case "callout": {
      const [ax, ay] = at(r * 0.92);
      const [x, y] = at(r + fs * 1.6);
      return callout(key, x + (right ? 1 : -1) * textW(text, fs) * 0.3, y, text, ax, ay, paint, fs);
    }
    case "ctr": {
      const [x, y] = at(hole ? (r + hole) / 2 : r * 0.5);
      return plain(key, x, y + fs * 0.35, text, "middle", "#fff", fs);
    }
    default: {
      const [x, y] = at(hole ? (r + hole) / 2 : r * 0.65);
      return plain(key, x, y + fs * 0.35, text, "middle", "#fff", fs);
    }
  }
}

/** Dış etiketler için pasta yarıçapından bırakılacak pay. */
export const pieLabelMargin = (pos: DataLabelPos, fs: number, labels: boolean) => (labels && (pos === "outEnd" || pos === "callout") ? fs * 2.4 : 0);
