/**
 * Grafiklerin DrawingML (ECMA-376 Bölüm 1, §21.2 "DrawingML - Charts")
 * olarak XLSX paketine yazılması ve geri okunması. exceljs grafik
 * desteklemediğinden paket JSZip ile sonradan işlenir.
 */
import JSZip from "jszip";
import { formatQualifiedRange, parseQualifiedRange } from "@officehere/formula";
import { PALETTES, rawChartData, seriesRefs } from "../charts/chartData";
import type { ChartDef, ChartSeriesDef, ChartType, Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import type { PackageParts } from "./xlsxPost";
import { NS_R, esc, nextRid, relsPathOf, resolveTarget, sheetPaths, unescapeXml } from "./ooxmlZip";
import { loadZip } from "./zipCache";

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

const EMU = 9525;

function absRef(sheetName: string, r: Rect): string {
  const a = formatQualifiedRange(sheetName, r);
  return a.replace(/(!|^)([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/, (_m, p, c1, r1, c2, r2) => `${p}$${c1}$${r1}${c2 ? `:$${c2}$${r2}` : ""}`);
}

/* ------------------------------ Yazma ------------------------------ */

function chartTypeXml(def: ChartDef): { tag: string; extra: string; grouping: string; dir?: string } {
  switch (def.type) {
    case "column":
      return { tag: "c:barChart", dir: "col", grouping: "clustered", extra: '<c:gapWidth val="219"/><c:overlap val="-27"/>' };
    case "columnStacked":
      return { tag: "c:barChart", dir: "col", grouping: "stacked", extra: '<c:gapWidth val="150"/><c:overlap val="100"/>' };
    case "bar":
      return { tag: "c:barChart", dir: "bar", grouping: "clustered", extra: '<c:gapWidth val="182"/>' };
    case "barStacked":
      return { tag: "c:barChart", dir: "bar", grouping: "stacked", extra: '<c:gapWidth val="150"/><c:overlap val="100"/>' };
    case "line":
    case "lineMarkers":
      return { tag: "c:lineChart", grouping: "standard", extra: '<c:marker val="1"/>' };
    case "area":
      return { tag: "c:areaChart", grouping: "standard", extra: "" };
    case "areaStacked":
      return { tag: "c:areaChart", grouping: "stacked", extra: "" };
    case "pie":
      return { tag: "c:pieChart", grouping: "", extra: '<c:firstSliceAng val="0"/>' };
    case "doughnut":
      return { tag: "c:doughnutChart", grouping: "", extra: '<c:firstSliceAng val="0"/><c:holeSize val="50"/>' };
    case "scatter":
      return { tag: "c:scatterChart", grouping: "", extra: "" };
  }
}

function strCache(values: string[]): string {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
}

function numCache(values: (number | null)[]): string {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join("")}</c:numCache>`;
}

function solid(hex: string): string {
  return `<a:solidFill><a:srgbClr val="${hex.replace("#", "")}"/></a:solidFill>`;
}

export function chartXml(wb: WorkbookModel, def: ChartDef): string {
  const refs = seriesRefs(wb, def);
  const data = rawChartData(wb, def);
  const t = chartTypeXml(def);
  const colors = PALETTES[def.palette % PALETTES.length];
  const isPie = def.type === "pie" || def.type === "doughnut";
  const isLine = def.type === "line" || def.type === "lineMarkers";
  const scatter = def.type === "scatter";
  const sers = (refs?.series ?? []).map((s, i) => {
    const color = def.seriesFmt?.[i]?.color ?? colors[i % colors.length];
    const sf = def.seriesFmt?.[i];
    const tx = s.name ? `<c:tx><c:strRef><c:f>${esc(absRef(refs!.sheetName, s.name))}</c:f>${strCache([data.series[i]?.name ?? ""])}</c:strRef></c:tx>` : `<c:tx><c:v>${esc(data.series[i]?.name ?? `Seri${i + 1}`)}</c:v></c:tx>`;
    let sp: string;
    if (isLine) sp = `<c:spPr><a:ln w="${Math.round((sf?.lineWidth ?? 2.25) * 12700)}" cap="rnd"><a:solidFill><a:srgbClr val="${color.slice(1)}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="${def.type === "lineMarkers" ? "circle" : "none"}"/>${def.type === "lineMarkers" ? `<c:size val="5"/><c:spPr>${solid(color)}</c:spPr>` : ""}</c:marker>`;
    else if (scatter) sp = `<c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${solid(color)}</c:spPr></c:marker>`;
    else if (isPie) sp = (data.categories.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr>${solid(colors[k % colors.length])}<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join(""));
    else sp = `<c:spPr>${solid(color)}${sf?.border ? `<a:ln w="${Math.round((sf.lineWidth ?? 0.75) * 12700)}">${solid(sf.border)}</a:ln>` : "<a:ln><a:noFill/></a:ln>"}</c:spPr>${def.type.startsWith("column") || def.type.startsWith("bar") ? '<c:invertIfNegative val="0"/>' : ""}`;
    const barLike = def.type.startsWith("column") || def.type.startsWith("bar");
    const dl = def.showDataLabels ? `<c:dLbls>${barLike && def.labelPos ? `<c:dLblPos val="${def.labelPos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${isPie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${isPie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : "";
    const catTag = scatter ? "c:xVal" : "c:cat";
    const valTag = scatter ? "c:yVal" : "c:val";
    const catRef = scatter ? s.xs : refs!.cat;
    const cat = catRef ? (scatter ? `<${catTag}><c:numRef><c:f>${esc(absRef(refs!.sheetName, catRef))}</c:f>${numCache(data.series[i]?.xs ?? [])}</c:numRef></${catTag}>` : `<${catTag}><c:strRef><c:f>${esc(absRef(refs!.sheetName, catRef))}</c:f>${strCache(data.categories)}</c:strRef></${catTag}>`) : "";
    const val = `<${valTag}><c:numRef><c:f>${esc(absRef(refs!.sheetName, s.values))}</c:f>${numCache(data.series[i]?.values ?? [])}</c:numRef></${valTag}>`;
    const smooth = isLine || scatter ? '<c:smooth val="0"/>' : "";
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${dl}${cat}${val}${smooth}</c:ser>`;
  });
  const axIds = scatter || !isPie ? '<c:axId val="111111111"/><c:axId val="222222222"/>' : "";
  const typeEl = `<${t.tag}>${t.dir ? `<c:barDir val="${t.dir}"/>` : ""}${t.grouping ? `<c:grouping val="${t.grouping}"/>` : ""}${scatter ? '<c:scatterStyle val="lineMarker"/>' : ""}<c:varyColors val="${isPie ? 1 : 0}"/>${(isPie ? sers.slice(0, 1) : sers).join("")}${t.extra}${axIds}</${t.tag}>`;
  const txPr = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr>';
  const va = def.valAxis ?? {};
  const grid = (def.showGridlines ? '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : "") + (va.minorGrid ? '<c:minorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:minorGridlines>' : "");
  const valScaling = `<c:scaling><c:orientation val="minMax"/>${va.max !== undefined ? `<c:max val="${va.max}"/>` : ""}${va.min !== undefined ? `<c:min val="${va.min}"/>` : ""}</c:scaling>`;
  const valNumFmt = va.numFmt ? `<c:numFmt formatCode="${esc(va.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>';
  const majorUnit = va.major ? `<c:majorUnit val="${va.major}"/>` : "";
  const axPos = def.type === "bar" || def.type === "barStacked" ? ["l", "b"] : ["b", "l"];
  let axes = "";
  if (!isPie) {
    const xTitle = def.xTitle ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(def.xTitle)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : "";
    const firstAx = scatter
      ? `<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${xTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${txPr}<c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${axPos[0]}"/>${xTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>${txPr}<c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const yTitle = def.yTitle ? `<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:r><a:t>${esc(def.yTitle)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : "";
    axes = `${firstAx}<c:valAx><c:axId val="222222222"/>${valScaling}<c:delete val="0"/><c:axPos val="${axPos[1]}"/>${grid}${yTitle}${valNumFmt}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPr}<c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="${scatter ? "midCat" : "between"}"/>${majorUnit}</c:valAx>`;
  }
  const title = def.showTitle
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="tr-TR" sz="1400" b="0"/><a:t>${esc(def.title || "Grafik Başlığı")}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legendPos = { right: "r", bottom: "b", top: "t", left: "l" }[def.legendPos];
  const legend = def.showLegend ? `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/>${txPr}</c:legend>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:date1904 val="0"/><c:lang val="tr-TR"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${typeEl}${axes}</c:plotArea>${legend}<c:plotVisOnly val="${def.showHidden ? 0 : 1}"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr>${solid("#FFFFFF")}<a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr></c:chartSpace>`;
}

function anchorXml(sheet: SheetModel, def: ChartDef, rid: string, id: number): string {
  const x1 = sheet.cols.offset(def.anchor.col) + def.anchor.dx;
  const y1 = sheet.rows.offset(def.anchor.row) + def.anchor.dy;
  const x2 = x1 + def.width;
  const y2 = y1 + def.height;
  const c2 = sheet.cols.indexAt(x2);
  const r2 = sheet.rows.indexAt(y2);
  const from = `<xdr:from><xdr:col>${def.anchor.col}</xdr:col><xdr:colOff>${Math.round(def.anchor.dx * EMU)}</xdr:colOff><xdr:row>${def.anchor.row}</xdr:row><xdr:rowOff>${Math.round(def.anchor.dy * EMU)}</xdr:rowOff></xdr:from>`;
  const to = `<xdr:to><xdr:col>${c2}</xdr:col><xdr:colOff>${Math.round((x2 - sheet.cols.offset(c2)) * EMU)}</xdr:colOff><xdr:row>${r2}</xdr:row><xdr:rowOff>${Math.round((y2 - sheet.rows.offset(r2)) * EMU)}</xdr:rowOff></xdr:to>`;
  return `<xdr:twoCellAnchor editAs="oneCell">${from}${to}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(def.name ?? `Grafik ${id - 999}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

export async function writeChartsInto(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  let chartNo = 1;
  while (zip.file(`xl/charts/chart${chartNo}.xml`)) chartNo++;
  let drawingNo = 1;
  while (zip.file(`xl/drawings/drawing${drawingNo}.xml`)) drawingNo++;
  for (const sheet of wb.sheets) {
    if (!sheet.charts.length) continue;
    const sheetPath = paths.get(sheet.name);
    if (!sheetPath) continue;
    const sheetRelsPath = relsPathOf(sheetPath);
    let sheetRels = zip.file(sheetRelsPath) ? await zip.file(sheetRelsPath)!.async("string") : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const sheetPart = await parts.part(sheetPath);
    if (!sheetPart) continue;
    // Mevcut çizim (resimler) var mı?
    let drawingPath: string | null = null;
    const relMatch = [...sheetRels.matchAll(/<Relationship\b[^>]*>/g)].find((m) => m[0].includes(REL_DRAWING));
    if (relMatch) {
      const target = /Target="([^"]+)"/.exec(relMatch[0])![1];
      drawingPath = resolveTarget(sheetPath, target);
    }
    let drawingXml: string;
    if (drawingPath && zip.file(drawingPath)) drawingXml = await zip.file(drawingPath)!.async("string");
    else {
      drawingPath = `xl/drawings/drawing${drawingNo++}.xml`;
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"></xdr:wsDr>`;
      const rid = nextRid(sheetRels);
      sheetRels = sheetRels.replace("</Relationships>", `<Relationship Id="${rid}" Type="${REL_DRAWING}" Target="../drawings/${drawingPath.split("/").pop()}"/></Relationships>`);
      // Şema sırasındaki yeri XmlPart belirler (pageSetup/headerFooter/sayfa sonlarından sonra)
      sheetPart.set("drawing", `<drawing r:id="${rid}"/>`);
      sheetPart.ensureNamespace("r", NS_R);
      ct = ct.replace("</Types>", `<Override PartName="/${drawingPath}" ContentType="${CT_DRAWING}"/></Types>`);
    }
    if (!/xmlns:r=/.test(drawingXml)) drawingXml = drawingXml.replace("<xdr:wsDr", `<xdr:wsDr xmlns:r="${NS_R}"`);
    const drawingRelsPath = relsPathOf(drawingPath);
    let drawingRels = zip.file(drawingRelsPath) ? await zip.file(drawingRelsPath)!.async("string") : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    let anchors = "";
    let nvId = 1000;
    for (const def of sheet.charts) {
      const chartPath = `xl/charts/chart${chartNo++}.xml`;
      zip.file(chartPath, chartXml(wb, def));
      ct = ct.replace("</Types>", `<Override PartName="/${chartPath}" ContentType="${CT_CHART}"/></Types>`);
      const rid = nextRid(drawingRels);
      drawingRels = drawingRels.replace("</Relationships>", `<Relationship Id="${rid}" Type="${REL_CHART}" Target="../charts/${chartPath.split("/").pop()}"/></Relationships>`);
      anchors += anchorXml(sheet, def, rid, nvId++);
    }
    drawingXml = drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>");
    zip.file(drawingPath, drawingXml);
    zip.file(drawingRelsPath, drawingRels);
    zip.file(sheetRelsPath, sheetRels);
  }
  zip.file("[Content_Types].xml", ct);
}

/* ------------------------------ Okuma ------------------------------ */

/** Seri renkleri/kenarlıkları, değer ekseni ölçeği ve sayı biçimi, etiket konumu. */
function readFormatting(cXml: string, def: ChartDef, valAx: string): void {
  const fmts: ChartDef["seriesFmt"] = [];
  for (const sm of cXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const sp = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(sm[1].replace(/<c:dPt>[\s\S]*?<\/c:dPt>|<c:marker>[\s\S]*?<\/c:marker>|<c:dLbls>[\s\S]*?<\/c:dLbls>/g, ""))?.[1] ?? "";
    const fill = /^\s*<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(sp)?.[1];
    const ln = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/.exec(sp);
    const lnColor = ln ? /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(ln[2])?.[1] : undefined;
    const w = ln ? Number(/\bw="(\d+)"/.exec(ln[1])?.[1] ?? NaN) / 12700 : NaN;
    const isLine = def.type === "line" || def.type === "lineMarkers";
    const f: { color?: string; border?: string; lineWidth?: number } = {};
    if (isLine && lnColor) f.color = `#${lnColor.toUpperCase()}`;
    else if (fill) f.color = `#${fill.toUpperCase()}`;
    if (!isLine && lnColor) f.border = `#${lnColor.toUpperCase()}`;
    if (Number.isFinite(w) && (isLine ? Math.abs(w - 2.25) > 0.01 : !!f.border)) f.lineWidth = Math.round(w * 100) / 100;
    // Paletin kendi rengi özel biçim sayılmaz (palet değişince seri rengi de değişir)
    const pal = PALETTES[def.palette % PALETTES.length];
    if (f.color && f.color === pal[fmts.length % pal.length].toUpperCase()) delete f.color;
    fmts.push(Object.keys(f).length ? f : null);
  }
  if (fmts.some(Boolean)) def.seriesFmt = fmts;
  if (valAx) {
    const va: NonNullable<ChartDef["valAxis"]> = {};
    const num = (re: RegExp) => {
      const m = re.exec(valAx);
      return m ? Number(m[1]) : undefined;
    };
    va.min = num(/<c:min val="(-?[\d.E+-]+)"/);
    va.max = num(/<c:max val="(-?[\d.E+-]+)"/);
    va.major = num(/<c:majorUnit val="([\d.E+-]+)"/);
    const nf = /<c:numFmt formatCode="([^"]*)" sourceLinked="0"/.exec(valAx)?.[1];
    if (nf) va.numFmt = unescapeXml(nf);
    if (/<c:minorGridlines/.test(valAx)) va.minorGrid = true;
    for (const k of Object.keys(va) as (keyof typeof va)[]) if (va[k] === undefined) delete va[k];
    if (Object.keys(va).length) def.valAxis = va;
  }
  const pos = /<c:dLblPos val="(outEnd|inEnd|ctr|inBase)"/.exec(cXml)?.[1];
  if (pos) def.labelPos = pos as ChartDef["labelPos"];
}

/** <c:ser> öğelerinden seri başvuruları; okunamayan varsa null (aralık tahminine dönülür). */
export function readSeries(cXml: string): ChartSeriesDef[] | null {
  const out: ChartSeriesDef[] = [];
  const block = (s: string, tag: string) => new RegExp(`<c:${tag}>([\\s\\S]*?)</c:${tag}>`).exec(s)?.[1];
  const fOf = (s: string | undefined) => {
    const m = s ? /<c:f>([^<]+)<\/c:f>/.exec(s) : null;
    return m ? unescapeXml(m[1]) : undefined;
  };
  const ok = (t: string | undefined) => t === undefined || !!parseQualifiedRange(t.replace(/\$/g, ""));
  for (const sm of cXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const s = sm[1];
    const tx = block(s, "tx");
    const values = fOf(block(s, "val") ?? block(s, "yVal"));
    if (!values || !ok(values)) return null;
    const categories = fOf(block(s, "cat"));
    const xValues = fOf(block(s, "xVal"));
    const nameRef = tx ? fOf(tx) : undefined;
    const name = tx && !nameRef ? unescapeXml(/<c:v>([^<]*)<\/c:v>/.exec(tx)?.[1] ?? "") : undefined;
    if (!ok(categories) || !ok(xValues) || !ok(nameRef)) return null;
    out.push({ values, ...(categories ? { categories } : {}), ...(xValues ? { xValues } : {}), ...(nameRef ? { nameRef } : {}), ...(name ? { name } : {}) });
  }
  return out.length ? out : null;
}

function detectType(xml: string): ChartType {
  const bar = /<c:barChart>/.test(xml) || /<c:bar3DChart>/.test(xml);
  const grouping = /<c:grouping val="(\w+)"/.exec(xml)?.[1] ?? "clustered";
  const stacked = grouping === "stacked" || grouping === "percentStacked";
  if (bar) {
    const dir = /<c:barDir val="(\w+)"/.exec(xml)?.[1] ?? "col";
    if (dir === "bar") return stacked ? "barStacked" : "bar";
    return stacked ? "columnStacked" : "column";
  }
  if (/<c:lineChart>|<c:line3DChart>/.test(xml)) return /<c:symbol val="none"\/>/.test(xml) ? "line" : "lineMarkers";
  if (/<c:areaChart>|<c:area3DChart>/.test(xml)) return stacked ? "areaStacked" : "area";
  if (/<c:doughnutChart>/.test(xml)) return "doughnut";
  if (/<c:pieChart>|<c:pie3DChart>|<c:ofPieChart>/.test(xml)) return "pie";
  if (/<c:scatterChart>|<c:bubbleChart>/.test(xml)) return "scatter";
  return "column";
}

export async function readCharts(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await loadZip(data);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const sheetPath = paths.get(sheet.name);
    if (!sheetPath) continue;
    const relsFile = zip.file(relsPathOf(sheetPath));
    if (!relsFile) continue;
    const rels = await relsFile.async("string");
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].find((m) => m[0].includes(REL_DRAWING));
    if (!rel) continue;
    const drawingPath = resolveTarget(sheetPath, /Target="([^"]+)"/.exec(rel[0])![1]);
    const dFile = zip.file(drawingPath);
    if (!dFile) continue;
    const dXml = await dFile.async("string");
    const dRelsFile = zip.file(relsPathOf(drawingPath));
    const dRels = dRelsFile ? await dRelsFile.async("string") : "";
    const relTargets = new Map<string, string>();
    for (const m of dRels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1];
      const t = /Target="([^"]+)"/.exec(m[0])?.[1];
      if (id && t) relTargets.set(id, resolveTarget(drawingPath, t));
    }
    for (const am of dXml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g)) {
      const a = am[0];
      const rid = /<c:chart\b[^>]*r:id="([^"]+)"/.exec(a)?.[1];
      if (!rid) continue;
      const chartPath = relTargets.get(rid);
      const cFile = chartPath ? zip.file(chartPath) : null;
      if (!cFile) continue;
      const cXml = await cFile.async("string");
      const num = (re: RegExp, s: string) => Number(re.exec(s)?.[1] ?? 0);
      const fromXml = /<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(a)?.[1] ?? "";
      const toXml = /<xdr:to>([\s\S]*?)<\/xdr:to>/.exec(a)?.[1] ?? "";
      const col = num(/<xdr:col>(\d+)<\/xdr:col>/, fromXml);
      const row = num(/<xdr:row>(\d+)<\/xdr:row>/, fromXml);
      const dx = num(/<xdr:colOff>(-?\d+)<\/xdr:colOff>/, fromXml) / EMU;
      const dy = num(/<xdr:rowOff>(-?\d+)<\/xdr:rowOff>/, fromXml) / EMU;
      let width = 480;
      let height = 288;
      if (toXml) {
        const x2 = sheet.cols.offset(num(/<xdr:col>(\d+)<\/xdr:col>/, toXml)) + num(/<xdr:colOff>(-?\d+)<\/xdr:colOff>/, toXml) / EMU;
        const y2 = sheet.rows.offset(num(/<xdr:row>(\d+)<\/xdr:row>/, toXml)) + num(/<xdr:rowOff>(-?\d+)<\/xdr:rowOff>/, toXml) / EMU;
        width = Math.max(60, x2 - (sheet.cols.offset(col) + dx));
        height = Math.max(40, y2 - (sheet.rows.offset(row) + dy));
      } else {
        const ext = /<xdr:ext cx="(\d+)" cy="(\d+)"/.exec(a);
        if (ext) {
          width = Number(ext[1]) / EMU;
          height = Number(ext[2]) / EMU;
        }
      }
      // Başvurulardan veri aralığını kur
      const refs = [...cXml.matchAll(/<c:f>([^<]+)<\/c:f>/g)].map((m) => unescapeXml(m[1]));
      let bbox: Rect | null = null;
      let sheetName = sheet.name;
      let valsHorizontal = 0;
      let valsVertical = 0;
      for (const vm of cXml.matchAll(/<c:(?:val|yVal)>\s*<c:numRef>\s*<c:f>([^<]+)<\/c:f>/g)) {
        const q = parseQualifiedRange(unescapeXml(vm[1]).replace(/\$/g, ""));
        if (q) q.range.r1 === q.range.r2 && q.range.c1 !== q.range.c2 ? valsHorizontal++ : valsVertical++;
      }
      for (const f of refs) {
        const q = parseQualifiedRange(f.replace(/\$/g, "").replace(/^\(|\)$/g, "").split(",")[0]);
        if (!q) continue;
        if (q.sheet) sheetName = q.sheet;
        bbox = bbox ? { r1: Math.min(bbox.r1, q.range.r1), c1: Math.min(bbox.c1, q.range.c1), r2: Math.max(bbox.r2, q.range.r2), c2: Math.max(bbox.c2, q.range.c2) } : { ...q.range };
      }
      if (!bbox) continue;
      const titleText = (xml: string) => [...(/<c:title>([\s\S]*?)<\/c:title>/.exec(xml)?.[1] ?? "").matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unescapeXml(m[1])).join("");
      const head = cXml.split("<c:plotArea>")[0];
      const title = titleText(head);
      const catAx = /<c:(?:catAx|dateAx)>[\s\S]*?<\/c:(?:catAx|dateAx)>/.exec(cXml)?.[0] ?? /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(cXml)?.[0] ?? "";
      const valAxes = [...cXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map((m) => m[0]);
      const valAx = valAxes.length > 1 ? valAxes[1] : catAx.startsWith("<c:valAx") ? "" : (valAxes[0] ?? "");
      const legendPos = /<c:legendPos val="(\w)"/.exec(cXml)?.[1];
      const def: ChartDef = {
        id: wb.newId("chart"),
        type: detectType(cXml),
        range: formatQualifiedRange(sheetName, bbox),
        seriesInRows: valsHorizontal > valsVertical,
        title,
        showTitle: /<c:title>/.test(head) && !/<c:autoTitleDeleted val="1"\/>/.test(cXml),
        showLegend: /<c:legend>/.test(cXml),
        legendPos: legendPos === "b" ? "bottom" : legendPos === "t" ? "top" : legendPos === "l" ? "left" : "right",
        showDataLabels: /<c:showVal val="1"\/>|<c:showPercent val="1"\/>/.test(cXml),
        showGridlines: /<c:majorGridlines/.test(cXml),
        palette: 0,
        style: 1,
        anchor: { row, col, dx, dy },
        width,
        height,
      };
      if (/<c:plotVisOnly val="(0|false)"\/>/.test(cXml)) def.showHidden = true;
      readFormatting(cXml, def, valAx);
      const nm = /<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(a)?.[1];
      if (nm && !/^(Grafik|Chart) \d+$/.test(unescapeXml(nm))) def.name = unescapeXml(nm);
      if (!def.title && def.showTitle) def.title = "";
      // Seriler dosyadaki başvurularla birebir (aralık tahmini yapılmaz)
      const series = readSeries(cXml);
      if (series) def.series = series;
      const xt = titleText(catAx);
      const yt = titleText(valAx);
      if (xt) def.xTitle = xt;
      if (yt) def.yTitle = yt;
      sheet.charts.push(def);
    }
  }
}
