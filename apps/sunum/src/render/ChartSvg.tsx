/**
 * Grafik çizici (SVG): Sütun, Çubuk, Çizgi, Alan, Pasta, Halka. Renkler tema
 * vurgu renklerinden; yazılar tema gövde yazı tipiyle çizilir.
 */
import { memo, type ReactNode } from "react";
import { fontStack } from "@officehere/core";
import { chartTitle, formatChartNumber, niceScale, seriesColor, valueRange } from "../model/chart";
import { cssColor } from "../model/color";
import type { ChartData, Theme } from "../model/types";


interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export const ChartSvg = memo(function ChartSvg({ chart, w, h, theme }: { chart: ChartData; w: number; h: number; theme: Theme }) {
  const W = Math.max(w, 10);
  const H = Math.max(h, 10);
  const fs = Math.max(7, Math.min(16, Math.min(W / 32, H / 18)));
  const font = fontStack(theme.fonts.minor).replace(/"/g, "'");
  const textColor = cssColor({ scheme: "dk1", lumMod: 0.65, lumOff: 0.35 }, theme);
  const gridColor = cssColor({ scheme: "dk1", lumMod: 0.15, lumOff: 0.85 }, theme);
  const pad = fs * 0.8;
  let top = pad;
  let bottom = H - pad;
  let left = pad;
  let right = W - pad;
  const parts: ReactNode[] = [];

  // Başlık
  if (chart.showTitle) {
    const tfs = fs * 1.55;
    parts.push(
      <text key="title" x={W / 2} y={top + tfs} textAnchor="middle" fontSize={tfs} fill={textColor}>
        {chartTitle(chart)}
      </text>,
    );
    top += tfs * 1.6;
  }

  const pie = chart.type === "pie" || chart.type === "doughnut";
  const legendItems = pie ? chart.categories.map((c, i) => ({ label: c, color: seriesColor(i, undefined) })) : chart.series.map((s, i) => ({ label: s.name, color: seriesColor(i, s.color) }));

  // Gösterge
  if (chart.legend !== "none" && legendItems.length) {
    const sw = fs * 0.7;
    const itemW = legendItems.map((it) => sw + fs * 0.4 + it.label.length * fs * 0.52 + fs * 1.2);
    if (chart.legend === "b" || chart.legend === "t") {
      const total = itemW.reduce((a, b) => a + b, 0);
      const ly = chart.legend === "b" ? bottom - fs * 0.4 : top + fs * 0.6;
      let lx = Math.max(left, W / 2 - total / 2);
      legendItems.forEach((it, i) => {
        parts.push(
          <g key={`lg${i}`}>
            <rect x={lx} y={ly - sw * 0.85} width={sw} height={sw} fill={cssColor(it.color, theme)} />
            <text x={lx + sw + fs * 0.35} y={ly} fontSize={fs} fill={textColor}>
              {it.label}
            </text>
          </g>,
        );
        lx += itemW[i];
      });
      if (chart.legend === "b") bottom -= fs * 1.9;
      else top += fs * 1.9;
    } else {
      const maxW = Math.min(W * 0.35, Math.max(...itemW));
      const lx = chart.legend === "r" ? right - maxW : left;
      const startY = (top + bottom) / 2 - (legendItems.length * fs * 1.4) / 2 + fs;
      legendItems.forEach((it, i) => {
        const y = startY + i * fs * 1.4;
        parts.push(
          <g key={`lg${i}`}>
            <rect x={lx} y={y - sw * 0.85} width={sw} height={sw} fill={cssColor(it.color, theme)} />
            <text x={lx + sw + fs * 0.35} y={y} fontSize={fs} fill={textColor}>
              {it.label}
            </text>
          </g>,
        );
      });
      if (chart.legend === "r") right -= maxW + fs;
      else left += maxW + fs;
    }
  }

  const plot: Box = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  if (pie) parts.push(...pieParts(chart, plot, theme, fs, textColor));
  else parts.push(...axisParts(chart, plot, theme, fs, textColor, gridColor));

  return (
    <svg className="chart-svg" width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ fontFamily: font }} role="img" aria-label={chart.showTitle ? chartTitle(chart) : "Grafik"}>
      {parts}
    </svg>
  );
});

function pieParts(chart: ChartData, plot: Box, theme: Theme, fs: number, textColor: string): ReactNode[] {
  const out: ReactNode[] = [];
  const vals = (chart.series[0]?.values ?? []).map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  const r = Math.max(4, Math.min(plot.w, plot.h) / 2 - fs * 0.3);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const hole = chart.type === "doughnut" ? r * 0.75 : 0;
  let a0 = -Math.PI / 2;
  vals.forEach((v, i) => {
    if (!total || !v) return;
    const a1 = a0 + (v / total) * Math.PI * 2;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const p = (ang: number, rr: number) => `${(cx + rr * Math.cos(ang)).toFixed(2)} ${(cy + rr * Math.sin(ang)).toFixed(2)}`;
    let d: string;
    if (v === total) {
      d = `M${cx - r} ${cy} A${r} ${r} 0 1 1 ${cx + r} ${cy} A${r} ${r} 0 1 1 ${cx - r} ${cy} Z`;
      if (hole) d += ` M${cx - hole} ${cy} A${hole} ${hole} 0 1 0 ${cx + hole} ${cy} A${hole} ${hole} 0 1 0 ${cx - hole} ${cy} Z`;
    } else if (hole) d = `M${p(a0, r)} A${r} ${r} 0 ${large} 1 ${p(a1, r)} L${p(a1, hole)} A${hole} ${hole} 0 ${large} 0 ${p(a0, hole)} Z`;
    else d = `M${cx} ${cy} L${p(a0, r)} A${r} ${r} 0 ${large} 1 ${p(a1, r)} Z`;
    out.push(<path key={`sl${i}`} d={d} fill={cssColor(seriesColor(i), theme)} fillRule="evenodd" stroke="#fff" strokeWidth={Math.max(0.75, r / 120)} />);
    if (chart.dataLabels) {
      const mid = (a0 + a1) / 2;
      const lr = hole ? (r + hole) / 2 : r * 0.65;
      out.push(
        <text key={`dl${i}`} x={cx + lr * Math.cos(mid)} y={cy + lr * Math.sin(mid) + fs * 0.35} textAnchor="middle" fontSize={fs} fill="#fff">
          {formatChartNumber(chart.series[0].values[i] ?? 0)}
        </text>,
      );
    }
    a0 = a1;
  });
  void textColor;
  return out;
}

function axisParts(chart: ChartData, plot: Box, theme: Theme, fs: number, textColor: string, gridColor: string): ReactNode[] {
  const out: ReactNode[] = [];
  const horiz = chart.type === "bar";
  const { min, max } = valueRange(chart);
  const sc = niceScale(min, max, Math.max(3, Math.floor((horiz ? plot.w : plot.h) / (fs * (horiz ? 6 : 2.6)))));
  const ticks: number[] = [];
  for (let v = sc.min; v <= sc.max + sc.step * 1e-6; v += sc.step) ticks.push(Math.round(v / sc.step) * sc.step);
  const valLabelW = Math.max(...ticks.map((t) => formatChartNumber(t).length)) * fs * 0.55 + fs * 0.6;
  const cats = chart.categories;
  const n = Math.max(1, cats.length);
  // Eksen etiketleri için yer ayır.
  const area: Box = horiz
    ? { x: plot.x + Math.min(plot.w * 0.35, Math.max(...cats.map((c) => c.length), 1) * fs * 0.55 + fs * 0.6), y: plot.y, w: 0, h: plot.h - fs * 1.6 }
    : { x: plot.x + valLabelW, y: plot.y + fs * 0.4, w: 0, h: plot.h - fs * 2 };
  area.w = plot.x + plot.w - area.x;
  const vpos = (v: number) => (v - sc.min) / (sc.max - sc.min || 1);

  // Kılavuz çizgileri ve değer etiketleri
  ticks.forEach((t, i) => {
    if (horiz) {
      const x = area.x + vpos(t) * area.w;
      out.push(<line key={`g${i}`} x1={x} x2={x} y1={area.y} y2={area.y + area.h} stroke={gridColor} strokeWidth={0.75} />);
      out.push(
        <text key={`vl${i}`} x={x} y={area.y + area.h + fs * 1.25} textAnchor="middle" fontSize={fs} fill={textColor}>
          {formatChartNumber(t)}
        </text>,
      );
    } else {
      const y = area.y + area.h - vpos(t) * area.h;
      out.push(<line key={`g${i}`} x1={area.x} x2={area.x + area.w} y1={y} y2={y} stroke={gridColor} strokeWidth={0.75} />);
      out.push(
        <text key={`vl${i}`} x={area.x - fs * 0.4} y={y + fs * 0.35} textAnchor="end" fontSize={fs} fill={textColor}>
          {formatChartNumber(t)}
        </text>,
      );
    }
  });

  // Kategori etiketleri
  const band = (horiz ? area.h : area.w) / n;
  cats.forEach((c, i) => {
    if (horiz) {
      // Çubuk grafikte ilk kategori altta.
      const y = area.y + area.h - (i + 0.5) * band;
      out.push(
        <text key={`cl${i}`} x={area.x - fs * 0.4} y={y + fs * 0.35} textAnchor="end" fontSize={fs} fill={textColor}>
          {c}
        </text>,
      );
    } else {
      const x = area.x + (i + 0.5) * band;
      out.push(
        <text key={`cl${i}`} x={x} y={area.y + area.h + fs * 1.25} textAnchor="middle" fontSize={fs} fill={textColor}>
          {c}
        </text>,
      );
    }
  });

  const zero = vpos(Math.max(sc.min, Math.min(sc.max, 0)));
  const ns = Math.max(1, chart.series.length);
  if (chart.type === "col" || chart.type === "bar") {
    // Aralık genişliği %219, çakışma -%27 (varsayılan kümelenmiş sütun).
    const barW = band / (ns + (ns - 1) * 0.27 + 2.19);
    const step = barW * 1.27;
    const start = (barW * 2.19) / 2;
    chart.series.forEach((s, j) => {
      const col = cssColor(seriesColor(j, s.color), theme);
      s.values.forEach((v, i) => {
        if (v === null || v === undefined || !Number.isFinite(v)) return;
        const p = vpos(Math.max(sc.min, Math.min(sc.max, v)));
        if (horiz) {
          const y = area.y + area.h - (i + 1) * band + start + (ns - 1 - j) * step;
          const x0 = area.x + Math.min(p, zero) * area.w;
          const ww = Math.abs(p - zero) * area.w;
          out.push(<rect key={`b${j}-${i}`} x={x0} y={y} width={Math.max(0, ww)} height={barW} fill={col} />);
          if (chart.dataLabels)
            out.push(
              <text key={`d${j}-${i}`} x={x0 + ww + fs * 0.3} y={y + barW / 2 + fs * 0.35} fontSize={fs * 0.9} fill={textColor}>
                {formatChartNumber(v)}
              </text>,
            );
        } else {
          const x = area.x + i * band + start + j * step;
          const y0 = area.y + area.h - Math.max(p, zero) * area.h;
          const hh = Math.abs(p - zero) * area.h;
          out.push(<rect key={`b${j}-${i}`} x={x} y={y0} width={barW} height={Math.max(0, hh)} fill={col} />);
          if (chart.dataLabels)
            out.push(
              <text key={`d${j}-${i}`} x={x + barW / 2} y={y0 - fs * 0.3} textAnchor="middle" fontSize={fs * 0.9} fill={textColor}>
                {formatChartNumber(v)}
              </text>,
            );
        }
      });
    });
  } else {
    // Çizgi ve alan: noktalar kategori ortalarında.
    const pts = (s: ChartData["series"][number]) => s.values.map((v, i) => (v === null || v === undefined || !Number.isFinite(v) ? null : ([area.x + (i + 0.5) * band, area.y + area.h - vpos(Math.max(sc.min, Math.min(sc.max, v))) * area.h] as [number, number])));
    const baseY = area.y + area.h - zero * area.h;
    chart.series.forEach((s, j) => {
      const col = cssColor(seriesColor(j, s.color), theme);
      const p = pts(s);
      const valid = p.filter((x): x is [number, number] => !!x);
      if (!valid.length) return;
      if (chart.type === "area") {
        const d = `M${valid[0][0]} ${baseY} ${valid.map(([x, y]) => `L${x.toFixed(2)} ${y.toFixed(2)}`).join(" ")} L${valid[valid.length - 1][0]} ${baseY} Z`;
        out.push(<path key={`a${j}`} d={d} fill={col} />);
      } else {
        let d = "";
        p.forEach((q, i) => {
          if (!q) return;
          d += `${d && p[i - 1] ? "L" : "M"}${q[0].toFixed(2)} ${q[1].toFixed(2)} `;
        });
        out.push(<path key={`l${j}`} d={d} fill="none" stroke={col} strokeWidth={Math.max(1.5, fs * 0.19)} strokeLinejoin="round" strokeLinecap="round" />);
      }
      if (chart.dataLabels)
        valid.forEach(([x, y], i) =>
          out.push(
            <text key={`d${j}-${i}`} x={x} y={y - fs * 0.5} textAnchor="middle" fontSize={fs * 0.9} fill={textColor}>
              {formatChartNumber(s.values.filter((v) => v !== null && v !== undefined)[i] as number)}
            </text>,
          ),
        );
    });
  }
  // Kategori ekseni çizgisi
  if (horiz) out.push(<line key="ax" x1={area.x + zero * area.w} x2={area.x + zero * area.w} y1={area.y} y2={area.y + area.h} stroke={gridColor} strokeWidth={1} />);
  else out.push(<line key="ax" x1={area.x} x2={area.x + area.w} y1={area.y + area.h - zero * area.h} y2={area.y + area.h - zero * area.h} stroke={gridColor} strokeWidth={1} />);
  return out;
}
