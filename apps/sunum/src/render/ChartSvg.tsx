/**
 * Grafik çizici (SVG): Sütun, Çubuk, Çizgi, Alan (kümelenmiş / yığılmış /
 * %100 yığılmış), Pasta, Halka, X Y (Dağılım) ve Radar. Renkler tema vurgu
 * renklerinden (Renkleri Değiştir paleti); yazılar tema gövde yazı tipiyle.
 */
import { memo, type ReactNode } from "react";
import { fontStack } from "@officehere/core";
import { chartStyle, chartTitle, effectiveLabelPos, formatChartNumber, hasAxes, isPieType, niceScale, scatterX, seriesColor, stackedValues, valueRange } from "../model/chart";
import { cssColor } from "../model/color";
import type { ChartData, Theme } from "../model/types";
import { radarParts } from "./chartRadar";
import { barLabel, pieLabelMargin, pointLabel, sliceLabel } from "./chartLabels";

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ChartPaint {
  fs: number;
  text: string;
  grid: string;
  minor: string;
  outline: string | null;
  color: (i: number) => string;
  /** Veri belirtme çizgisi kutusunun dolgusu. */
  labelBg: string;
}

export const ChartSvg = memo(function ChartSvg({ chart, w, h, theme }: { chart: ChartData; w: number; h: number; theme: Theme }) {
  const W = Math.max(w, 10);
  const H = Math.max(h, 10);
  const fs = Math.max(7, Math.min(16, Math.min(W / 32, H / 18)));
  const st = chartStyle(chart);
  const dark = st.bg === "dark";
  const font = fontStack(theme.fonts.minor).replace(/"/g, "'");
  const textColor = dark ? cssColor({ scheme: "lt1", lumMod: 0.85 }, theme) : cssColor({ scheme: "dk1", lumMod: 0.65, lumOff: 0.35 }, theme);
  const gridColor = dark ? cssColor({ scheme: "lt1", lumMod: 0.25, lumOff: 0.15 }, theme) : cssColor({ scheme: "dk1", lumMod: 0.15, lumOff: 0.85 }, theme);
  const pie = isPieType(chart.type);
  const paint: ChartPaint = {
    fs,
    text: textColor,
    grid: gridColor,
    minor: dark ? cssColor({ scheme: "lt1", lumMod: 0.15, lumOff: 0.1 }, theme) : cssColor({ scheme: "dk1", lumMod: 0.05, lumOff: 0.95 }, theme),
    outline: st.outline ? cssColor({ scheme: "lt1" }, theme) : null,
    color: (i) => cssColor(seriesColor(i, pie ? undefined : chart.series[i]?.color, chart.palette), theme),
    labelBg: dark ? cssColor({ scheme: "dk1", lumMod: 0.75, lumOff: 0.25 }, theme) : cssColor({ scheme: "lt1" }, theme),
  };
  const labels = !!chart.dataLabels || st.labels;
  const pad = fs * 0.8;
  let top = pad;
  let bottom = H - pad;
  let left = pad;
  let right = W - pad;
  const parts: ReactNode[] = [];
  const bg = chart.fill ? cssColor(chart.fill, theme) : st.bg === "dark" ? cssColor({ scheme: "dk1", lumMod: 0.75, lumOff: 0.25 }, theme) : st.bg === "light" ? cssColor({ scheme: "lt1", lumMod: 0.95 }, theme) : null;
  if (bg || chart.line) parts.push(<rect key="bg" x={0.5} y={0.5} width={W - 1} height={H - 1} fill={bg ?? "none"} stroke={chart.line ? cssColor(chart.line, theme) : "none"} />);

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

  const legendItems = pie ? chart.categories.map((c, i) => ({ label: c, color: paint.color(i) })) : chart.series.map((s, i) => ({ label: s.name, color: paint.color(i) }));

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
            <rect x={lx} y={ly - sw * 0.85} width={sw} height={sw} fill={it.color} />
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
            <rect x={lx} y={y - sw * 0.85} width={sw} height={sw} fill={it.color} />
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

  // Eksen başlıkları
  const at = hasAxes(chart.type) ? chart.axisTitles : undefined;
  const horiz = chart.type === "bar";
  const catTitle = horiz ? at?.val : at?.cat;
  const valTitle = horiz ? at?.cat : at?.val;
  if (catTitle) {
    parts.push(
      <text key="atc" x={(left + right) / 2} y={bottom - fs * 0.2} textAnchor="middle" fontSize={fs * 1.05} fill={textColor}>
        {catTitle}
      </text>,
    );
    bottom -= fs * 1.6;
  }
  if (valTitle) {
    const cy = (top + bottom) / 2;
    parts.push(
      <text key="atv" x={left + fs} y={cy} textAnchor="middle" fontSize={fs * 1.05} fill={textColor} transform={`rotate(-90 ${left + fs} ${cy})`}>
        {valTitle}
      </text>,
    );
    left += fs * 1.6;
  }

  const plot: Box = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  if (pie) parts.push(...pieParts(chart, plot, paint, labels));
  else if (chart.type === "radar") parts.push(...radarParts(chart, plot, paint, labels));
  else parts.push(...axisParts(chart, plot, paint, labels, st.gap));

  return (
    <svg className="chart-svg" width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ fontFamily: font }} role="img" aria-label={chart.showTitle ? chartTitle(chart) : "Grafik"}>
      {parts}
    </svg>
  );
});

function pieParts(chart: ChartData, plot: Box, paint: ChartPaint, labels: boolean): ReactNode[] {
  const out: ReactNode[] = [];
  const fs = paint.fs;
  const vals = (chart.series[0]?.values ?? []).map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  const lpos = effectiveLabelPos(chart);
  const r = Math.max(4, Math.min(plot.w, plot.h) / 2 - fs * 0.3 - pieLabelMargin(lpos, fs, labels));
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
    out.push(<path key={`sl${i}`} d={d} fill={paint.color(i)} fillRule="evenodd" stroke={paint.outline ?? "#fff"} strokeWidth={Math.max(0.75, r / 120)} />);
    if (labels) out.push(sliceLabel(`dl${i}`, cx, cy, r, hole, (a0 + a1) / 2, lpos, formatChartNumber(chart.series[0].values[i] ?? 0), paint));
    a0 = a1;
  });
  return out;
}

function axisParts(chart: ChartData, plot: Box, paint: ChartPaint, labels: boolean, gap: number): ReactNode[] {
  const out: ReactNode[] = [];
  const { fs, text: textColor, grid: gridColor } = paint;
  const horiz = chart.type === "bar";
  const scatter = chart.type === "scatter";
  const stacked = !!chart.grouping && chart.grouping !== "clustered" && !scatter;
  const pct = chart.grouping === "percentStacked" && !scatter;
  const { min, max } = valueRange(chart);
  const sc = pct ? { min: 0, max: 100, step: 20 } : niceScale(min, max, Math.max(3, Math.floor((horiz ? plot.w : plot.h) / (fs * (horiz ? 6 : 2.6)))));
  const ticks: number[] = [];
  for (let v = sc.min; v <= sc.max + sc.step * 1e-6; v += sc.step) ticks.push(Math.round(v / sc.step) * sc.step);
  const fmt = (t: number) => (pct ? `%${formatChartNumber(t)}` : formatChartNumber(t));
  const valLabelW = Math.max(...ticks.map((t) => fmt(t).length)) * fs * 0.55 + fs * 0.6;
  const cats = chart.categories;
  const n = Math.max(1, cats.length);
  // X ekseni (dağılım): sayısal ölçek.
  const xs = scatter ? scatterX(chart) : [];
  const xsc = scatter ? niceScale(Math.min(...xs, 0), Math.max(...xs, 1), 6) : null;
  const area: Box = horiz
    ? { x: plot.x + Math.min(plot.w * 0.35, Math.max(...cats.map((c) => c.length), 1) * fs * 0.55 + fs * 0.6), y: plot.y, w: 0, h: plot.h - fs * 1.6 }
    : { x: plot.x + valLabelW, y: plot.y + fs * 0.4, w: 0, h: plot.h - fs * 2 };
  area.w = plot.x + plot.w - area.x - (chart.type === "combo" && chart.secondaryLine ? valLabelW : 0);
  const vpos = (v: number) => (v - sc.min) / (sc.max - sc.min || 1);
  const showMajor = chart.gridMajor !== false;

  // İkincil kılavuz çizgileri
  if (chart.gridMinor) {
    for (let k = 0; k < ticks.length - 1; k++)
      for (let m = 1; m < 5; m++) {
        const t = ticks[k] + (sc.step * m) / 5;
        if (horiz) {
          const x = area.x + vpos(t) * area.w;
          out.push(<line key={`mg${k}-${m}`} x1={x} x2={x} y1={area.y} y2={area.y + area.h} stroke={paint.minor} strokeWidth={0.5} />);
        } else {
          const y = area.y + area.h - vpos(t) * area.h;
          out.push(<line key={`mg${k}-${m}`} x1={area.x} x2={area.x + area.w} y1={y} y2={y} stroke={paint.minor} strokeWidth={0.5} />);
        }
      }
  }
  // Ana kılavuz çizgileri ve değer etiketleri
  ticks.forEach((t, i) => {
    if (horiz) {
      const x = area.x + vpos(t) * area.w;
      if (showMajor) out.push(<line key={`g${i}`} x1={x} x2={x} y1={area.y} y2={area.y + area.h} stroke={gridColor} strokeWidth={0.75} />);
      out.push(
        <text key={`vl${i}`} x={x} y={area.y + area.h + fs * 1.25} textAnchor="middle" fontSize={fs} fill={textColor}>
          {fmt(t)}
        </text>,
      );
    } else {
      const y = area.y + area.h - vpos(t) * area.h;
      if (showMajor) out.push(<line key={`g${i}`} x1={area.x} x2={area.x + area.w} y1={y} y2={y} stroke={gridColor} strokeWidth={0.75} />);
      out.push(
        <text key={`vl${i}`} x={area.x - fs * 0.4} y={y + fs * 0.35} textAnchor="end" fontSize={fs} fill={textColor}>
          {fmt(t)}
        </text>,
      );
    }
  });

  // Kategori (ya da X değeri) etiketleri
  const band = (horiz ? area.h : area.w) / n;
  const xpos = (x: number) => (xsc ? area.x + ((x - xsc.min) / (xsc.max - xsc.min || 1)) * area.w : 0);
  if (xsc) {
    for (let v = xsc.min, k = 0; v <= xsc.max + xsc.step * 1e-6; v += xsc.step, k++)
      out.push(
        <text key={`xl${k}`} x={xpos(v)} y={area.y + area.h + fs * 1.25} textAnchor="middle" fontSize={fs} fill={textColor}>
          {formatChartNumber(Math.round(v / xsc.step) * xsc.step)}
        </text>,
      );
  } else
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
        out.push(
          <text key={`cl${i}`} x={area.x + (i + 0.5) * band} y={area.y + area.h + fs * 1.25} textAnchor="middle" fontSize={fs} fill={textColor}>
            {c}
          </text>,
        );
      }
    });

  const zero = vpos(Math.max(sc.min, Math.min(sc.max, 0)));
  const ns = Math.max(1, chart.series.length);
  const stack = stacked ? stackedValues(chart) : null;
  const lpos = effectiveLabelPos(chart);
  // Birleşik: son seri çizgi, diğerleri kümelenmiş sütun.
  const combo = chart.type === "combo";
  const barCount = combo ? Math.max(1, ns - 1) : ns;
  if (chart.type === "col" || chart.type === "bar" || combo) {
    const g = gap / 100;
    // Kümelenmişte seriler yan yana (çakışma -%27), yığılmışta üst üste.
    const barW = stacked ? band / (1 + g) : band / (barCount + (barCount - 1) * 0.27 + g);
    const step = stacked ? 0 : barW * 1.27;
    const start = (barW * g) / 2;
    chart.series.forEach((s, j) => {
      if (combo && j >= barCount && ns > 1) return;
      const col = paint.color(j);
      s.values.forEach((v, i) => {
        if (v === null || v === undefined || !Number.isFinite(v)) return;
        const seg = stack ? stack[j][i] : null;
        const a = vpos(Math.max(sc.min, Math.min(sc.max, seg ? seg.lo : Math.min(0, v))));
        const b = vpos(Math.max(sc.min, Math.min(sc.max, seg ? seg.hi : Math.max(0, v))));
        let rect: Box;
        if (horiz) {
          const y = area.y + area.h - (i + 1) * band + start + (stacked ? 0 : (ns - 1 - j) * step);
          rect = { x: area.x + a * area.w, y, w: Math.max(0, (b - a) * area.w), h: barW };
        } else {
          const y0 = area.y + area.h - b * area.h;
          rect = { x: area.x + i * band + start + j * step, y: y0, w: barW, h: Math.max(0, (b - a) * area.h) };
        }
        out.push(<rect key={`b${j}-${i}`} x={rect.x} y={rect.y} width={rect.w} height={rect.h} fill={col} stroke={paint.outline ?? undefined} strokeWidth={paint.outline ? 0.75 : undefined} />);
        if (labels) out.push(barLabel(`d${j}-${i}`, rect, horiz, v < 0, lpos, formatChartNumber(v), paint));
      });
    });
  }
  if (chart.type !== "col" && chart.type !== "bar" && !(combo && ns < 2)) {
    // İkincil eksen (birleşik): çizgi serisi kendi ölçeğiyle, değerleri sağda.
    const secondary = combo && !!chart.secondaryLine;
    const last = chart.series[ns - 1];
    const sc2 = secondary ? niceScale(Math.min(0, ...last.values.map((v) => v ?? 0)), Math.max(0, ...last.values.map((v) => v ?? 0)), Math.max(3, Math.floor(plot.h / (fs * 2.6)))) : null;
    if (sc2) {
      for (let v = sc2.min, k = 0; v <= sc2.max + sc2.step * 1e-6; v += sc2.step, k++) {
        const y = area.y + area.h - ((v - sc2.min) / (sc2.max - sc2.min || 1)) * area.h;
        out.push(
          <text key={`v2${k}`} x={area.x + area.w + fs * 0.4} y={y + fs * 0.35} textAnchor="start" fontSize={fs} fill={textColor}>
            {formatChartNumber(Math.round(v / sc2.step) * sc2.step)}
          </text>,
        );
      }
    }
    // Çizgi, alan, dağılım ve birleşik grafiğin çizgisi: noktalar kategori ortalarında (dağılımda X değerinde).
    const yOf = (v: number, j: number) =>
      sc2 && j === ns - 1 ? area.y + area.h - ((Math.max(sc2.min, Math.min(sc2.max, v)) - sc2.min) / (sc2.max - sc2.min || 1)) * area.h : area.y + area.h - vpos(Math.max(sc.min, Math.min(sc.max, v))) * area.h;
    const xOf = (i: number) => (scatter ? xpos(xs[i]) : area.x + (i + 0.5) * band);
    const baseY = area.y + area.h - zero * area.h;
    const order = combo ? [ns - 1] : chart.type === "area" && stacked ? chart.series.map((_, j) => j).reverse() : chart.series.map((_, j) => j);
    const ss = chart.scatterStyle;
    const drawLine = !scatter || !!ss;
    const smooth = scatter && (ss === "smooth" || ss === "smoothMarker");
    const markers = scatter && ss !== "line" && ss !== "smooth";
    for (const j of order) {
      const s = chart.series[j];
      const col = paint.color(j);
      const p = s.values.map((v, i) => (v === null || v === undefined || !Number.isFinite(v) ? null : ([xOf(i), yOf(stack ? stack[j][i].hi : v, j)] as [number, number])));
      const valid = p.filter((x): x is [number, number] => !!x);
      if (!valid.length) continue;
      if (chart.type === "area") {
        const lower = stack ? s.values.map((_, i) => [xOf(i), yOf(stack[j][i].lo, j)] as [number, number]).reverse() : [[valid[valid.length - 1][0], baseY] as [number, number], [valid[0][0], baseY] as [number, number]];
        const d = `M${valid.map(([x, y]) => `${x.toFixed(2)} ${y.toFixed(2)}`).join(" L")} L${lower.map(([x, y]) => `${x.toFixed(2)} ${y.toFixed(2)}`).join(" L")} Z`;
        out.push(<path key={`a${j}`} d={d} fill={col} stroke={paint.outline ?? undefined} />);
      } else {
        if (drawLine) {
          const d = smooth ? smoothPath(valid) : linePath(p);
          out.push(<path key={`l${j}`} d={d} fill="none" stroke={col} strokeWidth={Math.max(1.5, fs * 0.19)} strokeLinejoin="round" strokeLinecap="round" />);
        }
        if (markers) valid.forEach(([x, y], i) => out.push(<circle key={`m${j}-${i}`} cx={x} cy={y} r={Math.max(2.5, fs * 0.33)} fill={col} />));
      }
      if (labels) p.forEach((q, i) => q && out.push(chart.type === "area" ? pointLabel(`d${j}-${i}`, q[0], q[1] + fs, "ctr", formatChartNumber(s.values[i] as number), paint) : pointLabel(`d${j}-${i}`, q[0], q[1], combo ? (lpos === "ctr" || lpos === "callout" ? lpos : "t") : lpos, formatChartNumber(s.values[i] as number), paint)));
    }
  }
  // Kategori ekseni çizgisi
  if (horiz) out.push(<line key="ax" x1={area.x + zero * area.w} x2={area.x + zero * area.w} y1={area.y} y2={area.y + area.h} stroke={gridColor} strokeWidth={1} />);
  else out.push(<line key="ax" x1={area.x} x2={area.x + area.w} y1={area.y + area.h - zero * area.h} y2={area.y + area.h - zero * area.h} stroke={gridColor} strokeWidth={1} />);
  return out;
}

/** Kesik çizgi yolu (boş değerlerde ara verir). */
function linePath(p: ([number, number] | null)[]): string {
  let d = "";
  p.forEach((q, i) => {
    if (!q) return;
    d += `${d && p[i - 1] ? "L" : "M"}${q[0].toFixed(2)} ${q[1].toFixed(2)} `;
  });
  return d;
}

/** Düzgün eğri (Catmull-Rom → kübik Bézier). */
function smoothPath(pts: [number, number][]): string {
  if (pts.length < 3) return linePath(pts);
  let d = `M${pts[0][0].toFixed(2)} ${pts[0][1].toFixed(2)}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[Math.max(0, i - 1)];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[Math.min(pts.length - 1, i + 2)];
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += ` C${c1[0].toFixed(2)} ${c1[1].toFixed(2)} ${c2[0].toFixed(2)} ${c2[1].toFixed(2)} ${p2[0].toFixed(2)} ${p2[1].toFixed(2)}`;
  }
  return d;
}
