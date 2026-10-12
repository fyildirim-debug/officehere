// Grafik parçası (word/charts/chartN.xml, DrawingML Chart, ECMA-376 §21.2): değer önbellekli
// sütun/çubuk/çizgi/pasta grafikleri yazılır ve okunur.
import type { ChartData, ChartType } from "../../editor/chart";
import { pieColor } from "../../editor/chart";
import { XML_HEAD, cleanText, descendants, esc, kid, kids } from "./xml";

export const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const col = (i: number) => String.fromCharCode(66 + i); // B, C, D…
const hex = (c: string) => (/^#?[0-9a-f]{6}$/i.test(c) ? c.replace("#", "").toUpperCase() : "156082");
const fill = (c: string) => `<a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill>`;

function strCache(ref: string, vals: string[]) {
  return `<c:strRef><c:f>${ref}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(cleanText(v))}</c:v></c:pt>`).join("")}</c:strCache></c:strRef>`;
}

function numCache(ref: string, vals: number[]) {
  return `<c:numRef><c:f>${ref}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : "")).join("")}</c:numCache></c:numRef>`;
}

/** Grafik verisi → c:chartSpace. */
export function chartSpaceXml(c: ChartData): string {
  const n = c.categories.length;
  const serList = c.series.map((s, i) => {
      const tx = `<c:tx>${strCache(`Sayfa1!$${col(i)}$1`, [s.name])}</c:tx>`;
      const cat = `<c:cat>${strCache(`Sayfa1!$A$2:$A$${n + 1}`, c.categories)}</c:cat>`;
      const val = `<c:val>${numCache(`Sayfa1!$${col(i)}$2:$${col(i)}$${n + 1}`, c.categories.map((_, k) => s.values[k] ?? NaN))}</c:val>`;
      const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
      if (c.type === "line") return `<c:ser>${head}<c:spPr><a:ln w="28575" cap="rnd">${fill(s.color)}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(s.color)}</c:spPr></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
      if (c.type === "pie") {
        const dpts = c.categories.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr>${fill(pieColor(c, k))}<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("");
        return `<c:ser>${head}${dpts}${cat}${val}</c:ser>`;
      }
      return `<c:ser>${head}<c:spPr>${fill(s.color)}</c:spPr><c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
  });
  const sers = serList.join("");
  const axIds = `<c:axId val="50010001"/><c:axId val="50010002"/>`;
  let plot = "";
  // Pasta grafiği yalnız ilk seriyi gösterir.
  if (c.type === "pie") plot = `<c:pieChart><c:varyColors val="1"/>${serList[0] ?? ""}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (c.type === "line") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>`;
  else plot = `<c:barChart><c:barDir val="${c.type === "bar" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="219"/><c:overlap val="-27"/>${axIds}</c:barChart>`;
  const horizontal = c.type === "bar";
  const axes =
    c.type === "pie"
      ? ""
      : `<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? "l" : "b"}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
        `<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? "b" : "l"}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const title = c.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="tr-TR" sz="1400" b="0"/><a:t>${esc(cleanText(c.title))}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="1"/>`;
  return (
    `${XML_HEAD}<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:date1904 val="0"/><c:lang val="tr-TR"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`
  );
}

const valOf = (el: Element | null | undefined) => el?.getAttribute("val") ?? null;

function cacheValues(el: Element | null): string[] {
  if (!el) return [];
  const cache = descendants(el, "strCache")[0] ?? descendants(el, "numCache")[0] ?? descendants(el, "strLit")[0] ?? descendants(el, "numLit")[0];
  if (!cache) {
    const v = descendants(el, "v")[0];
    return v ? [v.textContent ?? ""] : [];
  }
  const count = Number(valOf(kid(cache, "ptCount")) ?? 0);
  const out: string[] = Array.from({ length: count }, () => "");
  for (const pt of kids(cache, "pt")) out[Number(pt.getAttribute("idx") ?? 0)] = descendants(pt, "v")[0]?.textContent ?? "";
  return out;
}

function colorOf(spPr: Element | null): string | null {
  if (!spPr) return null;
  const clr = descendants(spPr, "srgbClr")[0];
  const v = clr?.getAttribute("val");
  return v && /^[0-9a-f]{6}$/i.test(v) ? `#${v.toUpperCase()}` : null;
}

const DEFAULT_COLORS = ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"];

/** c:chartSpace → grafik verisi (desteklenmeyen tür sütun olarak alınır). */
export function readChartXml(doc: Document, width: number, height: number): ChartData | null {
  const root = doc.documentElement;
  const plotArea = descendants(root, "plotArea")[0];
  if (!plotArea) return null;
  const kinds = ["barChart", "bar3DChart", "lineChart", "line3DChart", "pieChart", "pie3DChart", "doughnutChart", "areaChart"];
  const ch = kids(plotArea).find((k) => kinds.includes(k.localName));
  if (!ch) return null;
  let type: ChartType = "column";
  if (ch.localName.startsWith("line")) type = "line";
  else if (ch.localName.startsWith("pie") || ch.localName === "doughnutChart") type = "pie";
  else if (valOf(kid(ch, "barDir")) === "bar") type = "bar";
  const sers = kids(ch, "ser");
  let categories: string[] = [];
  const series = sers.map((s, i) => {
    const cats = cacheValues(kid(s, "cat"));
    if (cats.length > categories.length) categories = cats;
    const vals = cacheValues(kid(s, "val")).map((v) => (v === "" ? NaN : Number(v)));
    const name = cacheValues(kid(s, "tx"))[0] || `Seri ${i + 1}`;
    const color = colorOf(kid(s, "spPr")) ?? colorOf(kid(kid(s, "marker"), "spPr")) ?? (type === "pie" ? colorOf(kid(kids(s, "dPt")[0] ?? null, "spPr")) : null) ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length];
    return { name, values: vals, color };
  });
  if (!categories.length) categories = (series[0]?.values ?? []).map((_, i) => String(i + 1));
  const titleEl = kid(kid(root, "chart"), "title");
  const deleted = valOf(kid(kid(root, "chart"), "autoTitleDeleted")) === "1";
  const title = titleEl ? descendants(titleEl, "t").map((t) => t.textContent ?? "").join("") || (sers.length === 1 ? series[0].name : "Grafik Başlığı") : deleted ? "" : sers.length === 1 ? series[0].name : "";
  return { type, title, categories, series: series.map((s) => ({ ...s, values: categories.map((_, k) => (Number.isFinite(s.values[k]) ? s.values[k] : 0)) })), width, height };
}
