/**
 * SVG grafik çizici (Sütun, Çubuk, Çizgi, Alan, Pasta, Halka, Dağılım).
 * Görünüm ofis grafiklerine yakındır: ince ızgara, gri eksen yazıları, gösterge.
 */
import { useState } from "react";
import { formatGeneral, formatValue } from "@officehere/formula";
import type { ChartDef } from "../model/types";
import { PALETTES, niceScale, type ChartData } from "./chartData";

/** Grafik stilleri (Grafik Tasarımı > Grafik Stilleri): arka plan, yazı ve ızgara renkleri. */
export const CHART_STYLES: { bg: string; ink: string; grid: string; border?: string }[] = [
  { bg: "#FFFFFF", ink: "#595959", grid: "#D9D9D9" },
  { bg: "#FFFFFF", ink: "#595959", grid: "#D9D9D9" },
  { bg: "#F2F2F2", ink: "#404040", grid: "#FFFFFF" },
  { bg: "#FFFFFF", ink: "#404040", grid: "#EDEDED", border: "#A6A6A6" },
  { bg: "#262626", ink: "#D9D9D9", grid: "#404040" },
  { bg: "#DEEBF7", ink: "#1F3864", grid: "#FFFFFF" },
  { bg: "#FBE5D6", ink: "#843C0C", grid: "#FFFFFF" },
  { bg: "#0E2841", ink: "#E8E8E8", grid: "#2B4A6B" },
];

export function themeOf(def: ChartDef): { bg: string; ink: string; grid: string; border?: string } {
  return CHART_STYLES[def.style] ?? CHART_STYLES[1];
}
const FONT = '"Segoe UI", "Aptos", "Carlito", sans-serif';

function fmt(v: number, numFmt?: string): string {
  if (numFmt && numFmt !== "General") return formatValue(v, numFmt).text.trim();
  return formatGeneral(Number(v.toPrecision(10)));
}

function truncate(s: string, max: number): string {
  return s.length > max ? s.slice(0, Math.max(1, max - 1)) + "…" : s;
}

/** Seri rengi: Veri Serisini Biçimlendir'de seçilen renk ya da paletten. */
export function seriesColor(def: ChartDef, i: number): string {
  const p = PALETTES[def.palette % PALETTES.length];
  return def.seriesFmt?.[i]?.color ?? p[i % p.length];
}

export function ChartView({ def, data, width, height }: { def: ChartDef; data: ChartData; width: number; height: number }) {
  const [hover, setHover] = useState<{ x: number; y: number; text: string } | null>(null);
  const isPieDef = def.type === "pie" || def.type === "doughnut";
  const colors = isPieDef ? PALETTES[def.palette % PALETTES.length] : Array.from({ length: Math.max(1, data.series.length) }, (_, i) => seriesColor(def, i));
  const isPie = def.type === "pie" || def.type === "doughnut";
  const titleH = def.showTitle ? 30 : 8;
  const legendItems = isPie ? data.categories.map((c, i) => ({ name: c, color: colors[i % colors.length] })) : data.series.map((s, i) => ({ name: s.name, color: colors[i % colors.length] }));
  const showLegend = def.showLegend && (legendItems.length > 1 || isPie);
  const legendRight = def.legendPos === "right" || def.legendPos === "left";
  const legendW = showLegend && legendRight ? Math.min(width * 0.32, Math.max(...legendItems.map((l) => l.name.length)) * 6.5 + 30) : 0;
  const legendH = showLegend && !legendRight ? 24 : 0;
  const plot = {
    x: 10 + (def.legendPos === "left" ? legendW : 0),
    y: titleH + (def.legendPos === "top" ? legendH : 0),
    w: Math.max(20, width - 20 - legendW),
    h: Math.max(20, height - titleH - legendH - 10),
  };
  const hoverProps = (text: string) => ({
    onMouseEnter: (e: React.MouseEvent) => {
      const r = (e.currentTarget as SVGElement).ownerSVGElement!.getBoundingClientRect();
      setHover({ x: e.clientX - r.left, y: e.clientY - r.top, text });
    },
    onMouseMove: (e: React.MouseEvent) => {
      const r = (e.currentTarget as SVGElement).ownerSVGElement!.getBoundingClientRect();
      setHover({ x: e.clientX - r.left, y: e.clientY - r.top, text });
    },
    onMouseLeave: () => setHover(null),
  });

  let body: React.ReactNode;
  if (!data.series.length || !data.series.some((s) => s.values.some((v) => v !== null))) {
    body = (
      <text x={width / 2} y={height / 2} textAnchor="middle" fontSize={12} fill={themeOf(def).ink} fontFamily={FONT}>
        Grafik verisi yok
      </text>
    );
  } else if (isPie) body = <PieBody def={def} data={data} plot={plot} colors={colors} hoverProps={hoverProps} />;
  else body = <AxisBody def={def} data={data} plot={plot} colors={colors} hoverProps={hoverProps} />;

  return (
    <div className="tb-chart-svgwrap">
      <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} role="img" aria-label={def.title || "Grafik"} style={{ fontFamily: FONT, display: "block" }}>
        <rect data-el="area" x={0.5} y={0.5} width={width - 1} height={height - 1} fill={def.fill ?? themeOf(def).bg} stroke={def.border === "none" ? "none" : (def.border ?? themeOf(def).border ?? "none")} />
        {def.showTitle && (
          <text data-el="title" x={width / 2} y={22} textAnchor="middle" fontSize={15} fill={themeOf(def).ink}>
            {truncate(def.title || "Grafik Başlığı", Math.floor(width / 8))}
          </text>
        )}
        {body}
        {showLegend && <Legend items={legendItems} def={def} width={width} height={height} plot={plot} legendW={legendW} titleH={titleH} />}
      </svg>
      {hover && (
        <div className="tb-chart-tip" style={{ left: Math.min(hover.x + 10, width - 140), top: Math.max(0, hover.y - 34) }}>
          {hover.text}
        </div>
      )}
    </div>
  );
}

type HoverProps = (text: string) => Record<string, unknown>;
interface BodyProps {
  def: ChartDef;
  data: ChartData;
  plot: { x: number; y: number; w: number; h: number };
  colors: string[];
  hoverProps: HoverProps;
}

function Legend({ items, def, width, height, plot, legendW, titleH }: { items: { name: string; color: string }[]; def: ChartDef; width: number; height: number; plot: BodyProps["plot"]; legendW: number; titleH: number }) {
  if (def.legendPos === "right" || def.legendPos === "left") {
    const x0 = def.legendPos === "right" ? width - legendW + 4 : 8;
    const total = items.length * 18;
    let y = Math.max(titleH, plot.y + (plot.h - total) / 2);
    return (
      <g fontSize={11} fill={themeOf(def).ink} data-el="legend">
        {items.map((it, i) => {
          const yy = y + i * 18;
          return (
            <g key={i}>
              <rect x={x0} y={yy + 3} width={9} height={9} fill={it.color} />
              <text x={x0 + 14} y={yy + 11}>
                {truncate(it.name, Math.floor((legendW - 20) / 6.2))}
              </text>
            </g>
          );
        })}
      </g>
    );
  }
  const widths = items.map((it) => Math.min(120, it.name.length * 6.3 + 22));
  const total = widths.reduce((a, b) => a + b, 0);
  let x = Math.max(4, (width - total) / 2);
  const y = def.legendPos === "top" ? titleH + 4 : height - 18;
  return (
    <g fontSize={11} fill={themeOf(def).ink} data-el="legend">
      {items.map((it, i) => {
        const xx = x;
        x += widths[i];
        return (
          <g key={i}>
            <rect x={xx} y={y + 2} width={9} height={9} fill={it.color} />
            <text x={xx + 13} y={y + 10}>
              {truncate(it.name, 17)}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function PieBody({ def, data, plot, colors, hoverProps }: BodyProps) {
  const s = data.series[0];
  const vals = s.values.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0) || 1;
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const r = Math.max(10, Math.min(plot.w, plot.h) / 2 - 6);
  const inner = def.type === "doughnut" ? r * 0.5 : 0;
  let a0 = -Math.PI / 2;
  return (
    <g>
      {vals.map((v, i) => {
        const a1 = a0 + (v / total) * Math.PI * 2;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = (ang: number, rad: number) => `${cx + rad * Math.cos(ang)},${cy + rad * Math.sin(ang)}`;
        let d: string;
        if (v / total >= 0.9999) {
          d = inner ? `M${p(0, r)}A${r},${r} 0 1 1 ${p(Math.PI, r)}A${r},${r} 0 1 1 ${p(0, r)}M${p(0, inner)}A${inner},${inner} 0 1 0 ${p(Math.PI, inner)}A${inner},${inner} 0 1 0 ${p(0, inner)}Z` : `M${p(0, r)}A${r},${r} 0 1 1 ${p(Math.PI, r)}A${r},${r} 0 1 1 ${p(0, r)}Z`;
        } else if (inner) d = `M${p(a0, r)}A${r},${r} 0 ${large} 1 ${p(a1, r)}L${p(a1, inner)}A${inner},${inner} 0 ${large} 0 ${p(a0, inner)}Z`;
        else d = `M${cx},${cy}L${p(a0, r)}A${r},${r} 0 ${large} 1 ${p(a1, r)}Z`;
        const mid = (a0 + a1) / 2;
        const label = def.showDataLabels && v > 0 ? `${Math.round((v / total) * 100)}%` : null;
        const lr = inner ? (r + inner) / 2 : r * 0.65;
        a0 = a1;
        return (
          <g key={i}>
            <path d={d} fill={colors[i % colors.length]} stroke={themeOf(def).bg} strokeWidth={1.5} className="tb-chart-mark" {...hoverProps(`${data.categories[i] ?? ""}: ${fmt(v, data.numFmt)} (${Math.round((v / total) * 1000) / 10}%)`)} />
            {label && (
              <text x={cx + lr * Math.cos(mid)} y={cy + lr * Math.sin(mid) + 4} textAnchor="middle" fontSize={11} fill="#FFFFFF" pointerEvents="none">
                {label}
              </text>
            )}
          </g>
        );
      })}
    </g>
  );
}

function AxisBody({ def, data, plot, colors, hoverProps }: BodyProps) {
  const horizontal = def.type === "bar" || def.type === "barStacked";
  const stacked = def.type === "columnStacked" || def.type === "barStacked" || def.type === "areaStacked";
  const scatter = def.type === "scatter";
  const n = data.categories.length;
  // Değer aralığı
  let vmin = 0;
  let vmax = 0;
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of data.series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) pos += v;
        else neg += v;
      }
      vmax = Math.max(vmax, pos);
      vmin = Math.min(vmin, neg);
    }
  } else {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
    vmin = Math.min(0, ...all);
    vmax = Math.max(0, ...all);
    if (def.type.startsWith("line") || scatter) {
      const lo = Math.min(...all);
      const hi = Math.max(...all);
      if (lo > 0 && hi > 0 && lo > hi * 0.6) vmin = lo;
      vmin = Math.min(vmin, lo);
      vmax = Math.max(vmax, hi);
    }
  }
  const va = def.valAxis ?? {};
  const sc = niceScale(va.min ?? vmin, va.max ?? vmax, horizontal ? Math.max(3, Math.floor(plot.w / 80)) : Math.max(3, Math.floor(plot.h / 45)));
  // Biçim bölmesinde sabitlenen sınırlar ve ana birim
  if (va.min !== undefined) sc.min = va.min;
  if (va.max !== undefined) sc.max = va.max;
  if (va.major !== undefined && va.major > 0 && (sc.max - sc.min) / va.major <= 200) sc.step = va.major;
  if (sc.max <= sc.min) sc.max = sc.min + (sc.step || 1);
  const ticks: number[] = [];
  for (let t = sc.min; t <= sc.max + sc.step / 1000; t += sc.step) ticks.push(Number(t.toPrecision(12)));
  const axisFmt = va.numFmt ?? data.numFmt;
  const tickLabels = ticks.map((t) => fmt(t, axisFmt));
  const labelW = Math.max(...tickLabels.map((l) => l.length)) * 6.5 + 8;
  // X ekseni (dağılım için sayısal)
  let xs: { min: number; max: number; step: number } | null = null;
  if (scatter) {
    const allX = data.series.flatMap((s) => (s.xs ?? s.values.map((_, i) => i + 1)).filter((v): v is number => v !== null));
    xs = niceScale(Math.min(0, ...allX), Math.max(...allX), 5);
  }
  const catLabelH = horizontal ? 0 : 18;
  const area = horizontal
    ? { x: plot.x + Math.min(plot.w * 0.35, Math.max(...data.categories.map((c) => c.length)) * 6.3 + 10), y: plot.y + 4, w: 0, h: plot.h - 22 }
    : { x: plot.x + labelW, y: plot.y + 6, w: plot.w - labelW - 6, h: plot.h - catLabelH - 8 };
  if (horizontal) area.w = plot.x + plot.w - area.x - 10;
  // Eksen başlıkları için yer aç
  if (def.yTitle) {
    area.x += 14;
    area.w -= 14;
  }
  if (def.xTitle) area.h -= 16;
  const vToPos = (v: number) => {
    const c = Math.max(sc.min, Math.min(sc.max, v));
    return horizontal ? area.x + ((c - sc.min) / (sc.max - sc.min)) * area.w : area.y + area.h - ((c - sc.min) / (sc.max - sc.min)) * area.h;
  };
  const zero = vToPos(Math.max(sc.min, Math.min(sc.max, 0)));
  const band = (horizontal ? area.h : area.w) / Math.max(1, n);
  const els: React.ReactNode[] = [];
  els.push(<rect key="plotbg" data-el="plot" x={area.x} y={area.y} width={Math.max(0, area.w)} height={Math.max(0, area.h)} fill="transparent" />);
  // İkincil kılavuz çizgileri (ana birimin beşte biri)
  if (va.minorGrid)
    for (let t = sc.min; t < sc.max - sc.step / 1000; t += sc.step)
      for (let k = 1; k < 5; k++) {
        const p = vToPos(t + (sc.step * k) / 5);
        els.push(horizontal ? <line key={`mg${t}-${k}`} x1={p} x2={p} y1={area.y} y2={area.y + area.h} stroke={themeOf(def).grid} strokeWidth={0.5} strokeDasharray="2 2" /> : <line key={`mg${t}-${k}`} x1={area.x} x2={area.x + area.w} y1={p} y2={p} stroke={themeOf(def).grid} strokeWidth={0.5} strokeDasharray="2 2" />);
      }
  // Izgara ve değer ekseni etiketleri
  ticks.forEach((t, i) => {
    const p = vToPos(t);
    if (horizontal) {
      if (def.showGridlines) els.push(<line key={`g${i}`} x1={p} x2={p} y1={area.y} y2={area.y + area.h} stroke={themeOf(def).grid} strokeWidth={1} />);
      els.push(
        <text key={`t${i}`} data-el="axis:val" x={p} y={area.y + area.h + 14} textAnchor="middle" fontSize={10.5} fill={themeOf(def).ink}>
          {tickLabels[i]}
        </text>,
      );
    } else {
      if (def.showGridlines) els.push(<line key={`g${i}`} x1={area.x} x2={area.x + area.w} y1={p} y2={p} stroke={themeOf(def).grid} strokeWidth={1} />);
      els.push(
        <text key={`t${i}`} data-el="axis:val" x={area.x - 6} y={p + 4} textAnchor="end" fontSize={10.5} fill={themeOf(def).ink}>
          {tickLabels[i]}
        </text>,
      );
    }
  });
  // Kategori etiketleri
  if (!scatter) {
    const every = Math.max(1, Math.ceil((n * (horizontal ? 14 : 40)) / Math.max(1, horizontal ? area.h : area.w)));
    data.categories.forEach((c, i) => {
      if (i % every) return;
      const mid = (horizontal ? area.y : area.x) + band * (i + 0.5);
      if (horizontal)
        els.push(
          <text key={`c${i}`} data-el="axis:cat" x={area.x - 6} y={mid + 4} textAnchor="end" fontSize={10.5} fill={themeOf(def).ink}>
            {truncate(c, 18)}
          </text>,
        );
      else
        els.push(
          <text key={`c${i}`} data-el="axis:cat" x={mid} y={area.y + area.h + 14} textAnchor="middle" fontSize={10.5} fill={themeOf(def).ink}>
            {truncate(c, Math.max(3, Math.floor((band * every) / 6.5)))}
          </text>,
        );
    });
  } else if (xs) {
    for (let t = xs.min; t <= xs.max + xs.step / 1000; t += xs.step) {
      const p = area.x + ((t - xs.min) / (xs.max - xs.min)) * area.w;
      els.push(
        <text key={`x${t}`} x={p} y={area.y + area.h + 14} textAnchor="middle" fontSize={10.5} fill={themeOf(def).ink}>
          {fmt(Number(t.toPrecision(12)))}
        </text>,
      );
    }
  }
  // Eksen çizgisi
  if (horizontal) els.push(<line key="axis" data-el="axis:cat" x1={zero} x2={zero} y1={area.y} y2={area.y + area.h} stroke="#BFBFBF" strokeWidth={3} strokeOpacity={0.35} />);
  else els.push(<line key="axis" data-el="axis:cat" x1={area.x} x2={area.x + area.w} y1={zero} y2={zero} stroke="#BFBFBF" />);

  const ns = data.series.length;
  if (def.type.startsWith("column") || def.type.startsWith("bar")) {
    const groupW = band * 0.68;
    const barW = stacked ? groupW : groupW / ns;
    const posAcc = new Array(n).fill(0);
    const negAcc = new Array(n).fill(0);
    data.series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (v === null) return;
        let from = 0;
        let to = v;
        if (stacked) {
          if (v >= 0) {
            from = posAcc[i];
            posAcc[i] += v;
            to = posAcc[i];
          } else {
            from = negAcc[i];
            negAcc[i] += v;
            to = negAcc[i];
          }
        }
        const p1 = vToPos(from);
        const p2 = vToPos(to);
        const start = (horizontal ? area.y : area.x) + band * i + (band - groupW) / 2 + (stacked ? 0 : barW * si);
        const color = colors[si % colors.length];
        const sf = def.seriesFmt?.[si];
        const stroke = sf?.border ? { stroke: sf.border, strokeWidth: ((sf.lineWidth ?? 0.75) * 4) / 3 } : {};
        const tip = `${s.name} — ${data.categories[i]}: ${fmt(v, data.numFmt)}`;
        if (horizontal) els.push(<rect key={`b${si}-${i}`} data-el={`series:${si}`} className="tb-chart-mark" x={Math.min(p1, p2)} y={start} width={Math.max(0.5, Math.abs(p2 - p1))} height={Math.max(1, barW - 1)} fill={color} {...stroke} {...hoverProps(tip)} />);
        else els.push(<rect key={`b${si}-${i}`} data-el={`series:${si}`} className="tb-chart-mark" x={start} y={Math.min(p1, p2)} width={Math.max(1, barW - 1)} height={Math.max(0.5, Math.abs(p2 - p1))} fill={color} {...stroke} {...hoverProps(tip)} />);
        if (def.showDataLabels) {
          // Konum: dış uç (varsayılan), iç uç, orta, iç taban
          const pos = def.labelPos ?? (stacked ? "ctr" : "outEnd");
          const hi = v >= 0 ? Math.max(p1, p2) : Math.min(p1, p2);
          const lo = v >= 0 ? Math.min(p1, p2) : Math.max(p1, p2);
          const inside = pos !== "outEnd";
          const at = horizontal
            ? pos === "outEnd" ? hi + (v >= 0 ? 4 : -4) : pos === "inEnd" ? hi + (v >= 0 ? -4 : 4) : pos === "ctr" ? (p1 + p2) / 2 : lo + (v >= 0 ? 4 : -4)
            : pos === "outEnd" ? Math.min(p1, p2) - 3 : pos === "inEnd" ? Math.min(p1, p2) + 12 : pos === "ctr" ? (p1 + p2) / 2 + 4 : Math.max(p1, p2) - 4;
          const anchor = horizontal ? (pos === "ctr" ? "middle" : (pos === "outEnd" || pos === "inBase") === v >= 0 ? "start" : "end") : "middle";
          els.push(
            <text key={`l${si}-${i}`} data-el="labels" x={horizontal ? at : start + barW / 2} y={horizontal ? start + barW / 2 + 4 : at} textAnchor={anchor} fontSize={10} fill={inside ? "#FFFFFF" : themeOf(def).ink} pointerEvents="none">
              {fmt(v, data.numFmt)}
            </text>,
          );
        }
      });
    });
  } else if (def.type.startsWith("line") || def.type.startsWith("area")) {
    const acc = new Array(n).fill(0);
    const area0 = def.type.startsWith("area");
    const layers: React.ReactNode[] = [];
    data.series.forEach((s, si) => {
      const color = colors[si % colors.length];
      const pts: [number, number, number][] = [];
      s.values.forEach((v, i) => {
        if (v === null && !area0) return;
        let val = v ?? 0;
        if (stacked) {
          acc[i] += val;
          val = acc[i];
        }
        pts.push([area.x + band * (i + 0.5), vToPos(val), v ?? 0]);
      });
      if (!pts.length) return;
      const path = pts.map((p, k) => `${k ? "L" : "M"}${p[0]},${p[1]}`).join("");
      if (area0) {
        const baseline = stacked && si > 0 ? null : zero;
        let d = path;
        if (baseline !== null) d += `L${pts[pts.length - 1][0]},${baseline}L${pts[0][0]},${baseline}Z`;
        else {
          // Yığılmış: önceki serinin üst çizgisine kapat
          const prev: [number, number][] = [];
          const accPrev = acc.map((a, i) => a - (s.values[i] ?? 0));
          s.values.forEach((_, i) => prev.push([area.x + band * (i + 0.5), vToPos(accPrev[i])]));
          d += prev.reverse().map((p) => `L${p[0]},${p[1]}`).join("") + "Z";
        }
        layers.unshift(<path key={`a${si}`} data-el={`series:${si}`} d={d} fill={color} fillOpacity={stacked ? 1 : 0.85} className="tb-chart-mark" {...hoverProps(s.name)} />);
      } else {
        layers.push(<path key={`p${si}`} data-el={`series:${si}`} d={path} fill="none" stroke={color} strokeWidth={((def.seriesFmt?.[si]?.lineWidth ?? 2.25) * 4) / 3} strokeLinejoin="round" />);
        pts.forEach((p, k) => {
          const showMarker = def.type === "lineMarkers";
          layers.push(<circle key={`m${si}-${k}`} data-el={`series:${si}`} cx={p[0]} cy={p[1]} r={showMarker ? 4 : 6} fill={showMarker ? color : "transparent"} stroke={showMarker ? "#FFFFFF" : "none"} strokeWidth={1} className="tb-chart-mark" {...hoverProps(`${s.name} — ${data.categories[k]}: ${fmt(p[2], data.numFmt)}`)} />);
        });
      }
      if (def.showDataLabels)
        pts.forEach((p, k) =>
          layers.push(
            <text key={`dl${si}-${k}`} x={p[0]} y={p[1] - 7} textAnchor="middle" fontSize={10} fill={themeOf(def).ink} pointerEvents="none">
              {fmt(p[2], data.numFmt)}
            </text>,
          ),
        );
    });
    els.push(...layers);
  } else if (scatter && xs) {
    const sx = xs;
    data.series.forEach((s, si) => {
      const color = colors[si % colors.length];
      s.values.forEach((v, i) => {
        const xv = s.xs?.[i] ?? i + 1;
        if (v === null || xv === null) return;
        const px = area.x + ((xv - sx.min) / (sx.max - sx.min)) * area.w;
        els.push(<circle key={`s${si}-${i}`} data-el={`series:${si}`} cx={px} cy={vToPos(v)} r={4.5} fill={color} stroke="#FFFFFF" strokeWidth={1} className="tb-chart-mark" {...hoverProps(`${s.name} (${fmt(xv)}; ${fmt(v, data.numFmt)})`)} />);
      });
    });
  }
  if (def.yTitle)
    els.push(
      <text key="yt" x={plot.x + 2} y={area.y + area.h / 2} fontSize={11} fill={themeOf(def).ink} transform={`rotate(-90 ${plot.x + 8} ${area.y + area.h / 2})`} textAnchor="middle">
        {def.yTitle}
      </text>,
    );
  if (def.xTitle)
    els.push(
      <text key="xt" x={area.x + area.w / 2} y={plot.y + plot.h - 2} fontSize={11} fill={themeOf(def).ink} textAnchor="middle">
        {def.xTitle}
      </text>,
    );
  return <g>{els}</g>;
}
