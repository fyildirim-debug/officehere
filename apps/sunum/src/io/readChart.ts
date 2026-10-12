/** Grafik parçası okuyucusu (c:chartSpace → ChartData; önbellekli değerler kullanılır). */
import type { ChartData, ChartSeries, ChartType } from "../model/types";
import { readFill, type ReadCtx } from "./readStyle";
import { attr, descendants, kid, kids, numAttr, parseXml } from "./xml";

const TYPES: Record<string, ChartType> = {
  barChart: "col",
  bar3DChart: "col",
  lineChart: "line",
  line3DChart: "line",
  stockChart: "line",
  radarChart: "radar",
  scatterChart: "scatter",
  areaChart: "area",
  area3DChart: "area",
  pieChart: "pie",
  pie3DChart: "pie",
  ofPieChart: "pie",
  doughnutChart: "doughnut",
};

/** strCache/numCache/strLit/numLit noktaları (idx sırasıyla). */
function cachePoints(el: Element | undefined): { count: number; pts: Map<number, string> } {
  const pts = new Map<number, string>();
  if (!el) return { count: 0, pts };
  const cache = descendants(el, "strCache")[0] ?? descendants(el, "numCache")[0] ?? descendants(el, "strLit")[0] ?? descendants(el, "numLit")[0];
  if (!cache) return { count: 0, pts };
  for (const pt of kids(cache, "pt")) pts.set(numAttr(pt, "idx") ?? 0, kid(pt, "v")?.textContent ?? "");
  const count = numAttr(kid(cache, "ptCount"), "val") ?? (pts.size ? Math.max(...pts.keys()) + 1 : 0);
  return { count, pts };
}

function richText(tx: Element | undefined): string {
  if (!tx) return "";
  const rich = kid(tx, "rich");
  if (rich) return kids(rich, "p").map((p) => descendants(p, "t").map((t) => t.textContent ?? "").join("")).join(" ").trim();
  const { pts } = cachePoints(tx);
  return pts.get(0) ?? kid(tx, "v")?.textContent ?? "";
}

export function readChartXml(xml: string, ctx: ReadCtx): ChartData | null {
  const doc = parseXml(xml).documentElement;
  const chart = kid(doc, "chart");
  const plot = kid(chart, "plotArea");
  if (!plot) return null;
  const typeEl = kids(plot).find((k) => TYPES[k.localName]);
  if (!typeEl) return null;
  let type = TYPES[typeEl.localName];
  if (type === "col" && attr(kid(typeEl, "barDir"), "val") === "bar") type = "bar";
  let serEls = kids(typeEl, "ser");
  // Sütun + çizgi aynı eksenlerde: Birleşik grafik.
  const lineEl = typeEl.localName === "barChart" ? kid(plot, "lineChart") : undefined;
  if (lineEl && type === "col") {
    type = "combo";
    serEls = [...serEls, ...kids(lineEl, "ser")];
  }
  let categories: string[] = [];
  const series: ChartSeries[] = [];
  for (const [j, ser] of serEls.entries()) {
    const cat = cachePoints(kid(ser, "cat") ?? kid(ser, "xVal"));
    const val = cachePoints(kid(ser, "val") ?? kid(ser, "yVal"));
    const n = Math.max(cat.count, val.count);
    if (j === 0 || categories.length < n) categories = Array.from({ length: n }, (_, i) => cat.pts.get(i) ?? `${i + 1}`);
    const values = Array.from({ length: n }, (_, i) => {
      const v = val.pts.get(i);
      if (v === undefined || v === "") return null;
      const num = Number(v);
      return Number.isFinite(num) ? num : null;
    });
    const fill = readFill(kid(ser, "spPr"), ctx);
    const ln = kid(kid(ser, "spPr"), "ln");
    const lnFill = ln ? readFill(ln, ctx) : undefined;
    const lineLike = type === "line" || type === "radar" || (type === "combo" && j === serEls.length - 1);
    const color = lineLike ? (lnFill?.type === "solid" ? lnFill.color : undefined) : fill?.type === "solid" ? fill.color : undefined;
    series.push({ name: richText(kid(ser, "tx")) || `Seri ${j + 1}`, values, color: type === "pie" || type === "doughnut" ? undefined : color });
  }
  const titleEl = kid(chart, "title");
  const autoDeleted = attr(kid(chart, "autoTitleDeleted"), "val");
  const showTitle = !!titleEl || (autoDeleted === "0" && series.length === 1);
  const legendEl = kid(chart, "legend");
  const pos = attr(kid(legendEl, "legendPos"), "val");
  // CT_LegendPos varsayılanı "r" (sağ); öğe yoksa da sağ kabul edilir.
  const legend: ChartData["legend"] = !legendEl ? "none" : pos === "t" || pos === "l" || pos === "b" ? pos : "r";
  const dl = descendants(typeEl, "dLbls").some((d) => attr(kid(d, "showVal"), "val") === "1");
  const g = attr(kid(typeEl, "grouping"), "val");
  const grouping: ChartData["grouping"] = g === "stacked" || g === "percentStacked" ? g : undefined;
  // Eksenler: kategori ekseni (dağılımda ilk değer ekseni) ve değer ekseni.
  const valAxes = kids(plot, "valAx");
  const catAx = kid(plot, "catAx") ?? kid(plot, "dateAx") ?? (type === "scatter" ? valAxes[0] : undefined);
  const valAx = type === "scatter" ? valAxes[1] : valAxes[0];
  const catT = catAx ? richText(kid(kid(catAx, "title"), "tx")) : "";
  const valT = valAx ? richText(kid(kid(valAx, "title"), "tx")) : "";
  const axisTitles = catT || valT ? { cat: catT || undefined, val: valT || undefined } : undefined;
  const styleVal = numAttr(kid(doc, "style"), "val");
  const space = readFill(kid(doc, "spPr"), ctx);
  const spLn = kid(kid(doc, "spPr"), "ln");
  const spLnFill = spLn ? readFill(spLn, ctx) : undefined;
  return {
    type,
    title: titleEl ? richText(kid(titleEl, "tx")) || undefined : undefined,
    showTitle,
    categories,
    series: series.length ? series : [{ name: "Seri 1", values: categories.map(() => null) }],
    legend,
    dataLabels: dl || undefined,
    grouping,
    axisTitles,
    gridMajor: valAx && !kid(valAx, "majorGridlines") ? false : undefined,
    gridMinor: valAx && kid(valAx, "minorGridlines") ? true : undefined,
    style: styleVal && styleVal >= 1 && styleVal <= 8 ? styleVal : undefined,
    fill: space?.type === "solid" ? space.color : undefined,
    line: spLnFill?.type === "solid" ? spLnFill.color : undefined,
  };
}
