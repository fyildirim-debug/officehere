/** Radar grafiği: kategoriler eşit açılı kollarda, seriler çokgen çizgi olarak. */
import type { ReactNode } from "react";
import { formatChartNumber, niceScale, valueRange } from "../model/chart";
import type { ChartData } from "../model/types";
import type { Box, ChartPaint } from "./ChartSvg";

export function radarParts(chart: ChartData, plot: Box, paint: ChartPaint, labels: boolean): ReactNode[] {
  const out: ReactNode[] = [];
  const { fs } = paint;
  const n = Math.max(3, chart.categories.length);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const r = Math.max(8, Math.min(plot.w, plot.h) / 2 - fs * 1.6);
  const { min, max } = valueRange(chart);
  const sc = niceScale(Math.min(0, min), max, 5);
  const rel = (v: number) => (Math.max(sc.min, Math.min(sc.max, v)) - sc.min) / (sc.max - sc.min || 1);
  const ang = (i: number) => -Math.PI / 2 + (i * Math.PI * 2) / n;
  const pt = (i: number, rr: number): [number, number] => [cx + rr * Math.cos(ang(i)), cy + rr * Math.sin(ang(i))];
  // Ağ (kılavuz çokgenleri) ve kollar
  let k = 0;
  for (let v = sc.min; v <= sc.max + sc.step * 1e-6; v += sc.step, k++) {
    const rr = rel(v) * r;
    const d = Array.from({ length: n }, (_, i) => pt(i, rr)).map(([x, y], i) => `${i ? "L" : "M"}${x.toFixed(2)} ${y.toFixed(2)}`).join(" ") + " Z";
    if (chart.gridMajor !== false || k === 0) out.push(<path key={`rg${k}`} d={d} fill="none" stroke={paint.grid} strokeWidth={0.75} />);
    out.push(
      <text key={`rv${k}`} x={cx + fs * 0.3} y={cy - rr + fs * 0.35} fontSize={fs * 0.85} fill={paint.text}>
        {formatChartNumber(Math.round(v / sc.step) * sc.step)}
      </text>,
    );
  }
  for (let i = 0; i < n; i++) {
    const [x, y] = pt(i, r);
    out.push(<line key={`rs${i}`} x1={cx} y1={cy} x2={x} y2={y} stroke={paint.grid} strokeWidth={0.75} />);
    const [lx, ly] = pt(i, r + fs * 0.9);
    out.push(
      <text key={`rc${i}`} x={lx} y={ly + fs * 0.35} textAnchor={Math.abs(lx - cx) < fs ? "middle" : lx > cx ? "start" : "end"} fontSize={fs} fill={paint.text}>
        {chart.categories[i] ?? ""}
      </text>,
    );
  }
  // Seriler
  chart.series.forEach((s, j) => {
    const pts = Array.from({ length: n }, (_, i) => pt(i, rel(s.values[i] ?? 0) * r));
    const d = pts.map(([x, y], i) => `${i ? "L" : "M"}${x.toFixed(2)} ${y.toFixed(2)}`).join(" ") + " Z";
    out.push(<path key={`rp${j}`} d={d} fill="none" stroke={paint.color(j)} strokeWidth={Math.max(1.5, fs * 0.19)} strokeLinejoin="round" />);
    if (labels)
      pts.forEach(([x, y], i) =>
        out.push(
          <text key={`rd${j}-${i}`} x={x} y={y - fs * 0.4} textAnchor="middle" fontSize={fs * 0.85} fill={paint.text}>
            {formatChartNumber(s.values[i] ?? 0)}
          </text>,
        ),
      );
  });
  return out;
}
