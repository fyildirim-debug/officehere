/**
 * Grafik parçası yazıcısı (DrawingML Charts, ECMA-376 Bölüm 21.2):
 * c:chartSpace + önbellekli veriler ve gömülü çalışma kitabı (xlsx).
 */
import JSZip from "jszip";
import { chartTitle, parseChartNumber, seriesColor } from "../model/chart";
import type { ChartData } from "../model/types";
import { colorXml } from "./writeDrawing";
import { esc, NS, XML_HEAD } from "./xml";

const SHEET = "Sayfa1";

function colName(i: number): string {
  let s = "";
  let n = i + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function strCache(values: string[]): string {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
}

function numCache(values: (number | null)[]): string {
  const pts = values.map((v, i) => (v === null || v === undefined || !Number.isFinite(v) ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("");
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>`;
}

const txt = (sz: number, color = '<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>') =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"><a:solidFill>${color}</a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr>`;

const GRID_LN = '<c:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></c:spPr>';

function dLblsXml(show: boolean | undefined, pos?: string): string {
  if (!show) return "";
  return `<c:dLbls>${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

/** Eksen başlığı (c:title, zengin metin). */
function axisTitleXml(t: string | undefined, rot?: number): string {
  if (!t) return "";
  return `<c:title><c:tx><c:rich><a:bodyPr${rot ? ` rot="${rot}" vert="horz"` : ""}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1330" b="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:r><a:rPr lang="tr-TR"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

/** Grafik XML'i (chartN.xml). `embedRid`: gömülü çalışma kitabı ilişkisi. */
export function chartXml(c: ChartData, embedRid?: string): string {
  const n = c.categories.length;
  const lastRow = n + 1;
  const catRef = `${SHEET}!$A$2:$A$${lastRow}`;
  const pie = c.type === "pie" || c.type === "doughnut";
  const series = pie ? c.series.slice(0, 1) : c.series;
  const ctx = { slideContent: true };
  const pal = c.palette;
  // Birleşik: son seri çizgi, diğerleri kümelenmiş sütun.
  const isLineSer = (j: number) => c.type === "line" || (c.type === "combo" && j === series.length - 1);
  const serList = series.map((s, j) => {
      const col = colName(j + 1);
      const color = seriesColor(j, s.color, pal);
      const tx = `<c:tx><c:strRef><c:f>${SHEET}!$${col}$1</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
      const cat = `<c:cat><c:strRef><c:f>${catRef}</c:f>${strCache(c.categories)}</c:strRef></c:cat>`;
      const val = `<c:val><c:numRef><c:f>${SHEET}!$${col}$2:$${col}$${lastRow}</c:f>${numCache(s.values.slice(0, n))}</c:numRef></c:val>`;
      const head = `<c:idx val="${j}"/><c:order val="${j}"/>${tx}`;
      if (pie) {
        const dpts = c.categories.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${colorXml(seriesColor(i, undefined, pal), ctx)}</a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("");
        return `<c:ser>${head}${dpts}${dLblsXml(c.dataLabels, c.type === "pie" ? "ctr" : undefined)}${cat}${val}</c:ser>`;
      }
      if (c.type === "scatter") {
        const xs = c.categories.map((t) => parseChartNumber(t));
        const xVal = xs.every((v) => v !== null) ? `<c:xVal><c:numRef><c:f>${catRef}</c:f>${numCache(xs)}</c:numRef></c:xVal>` : `<c:xVal><c:strRef><c:f>${catRef}</c:f>${strCache(c.categories)}</c:strRef></c:xVal>`;
        const yVal = `<c:yVal><c:numRef><c:f>${SHEET}!$${col}$2:$${col}$${lastRow}</c:f>${numCache(s.values.slice(0, n))}</c:numRef></c:yVal>`;
        const sp = `<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill>${colorXml(color, ctx)}</a:solidFill></c:spPr></c:marker>`;
        return `<c:ser>${head}${sp}${dLblsXml(c.dataLabels, "t")}${xVal}${yVal}<c:smooth val="0"/></c:ser>`;
      }
      if (c.type === "radar") {
        const sp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${colorXml(color, ctx)}</a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`;
        return `<c:ser>${head}${sp}${dLblsXml(c.dataLabels)}${cat}${val}</c:ser>`;
      }
      if (isLineSer(j)) {
        const sp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${colorXml(color, ctx)}</a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`;
        return `<c:ser>${head}${sp}${dLblsXml(c.dataLabels, "t")}${cat}${val}<c:smooth val="0"/></c:ser>`;
      }
      const sp = `<c:spPr><a:solidFill>${colorXml(color, ctx)}</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;
      if (c.type === "area") return `<c:ser>${head}${sp}${dLblsXml(c.dataLabels)}${cat}${val}</c:ser>`;
      return `<c:ser>${head}${sp}<c:invertIfNegative val="0"/>${dLblsXml(c.dataLabels, "outEnd")}${cat}${val}</c:ser>`;
  });
  const sers = serList.join("");

  const AX1 = 500000001;
  const AX2 = 500000002;
  const axIds = `<c:axId val="${AX1}"/><c:axId val="${AX2}"/>`;
  let plot = "";
  const grp = c.grouping ?? "clustered";
  const lineGrp = grp === "clustered" ? "standard" : grp;
  const stacked = grp !== "clustered";
  switch (c.type) {
    case "col":
    case "bar":
      plot = `<c:barChart><c:barDir val="${c.type === "bar" ? "bar" : "col"}"/><c:grouping val="${grp}"/><c:varyColors val="0"/>${sers}<c:gapWidth val="${stacked ? 150 : 219}"/><c:overlap val="${stacked ? 100 : -27}"/>${axIds}</c:barChart>`;
      break;
    case "line":
      plot = `<c:lineChart><c:grouping val="${lineGrp}"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>`;
      break;
    case "area":
      plot = `<c:areaChart><c:grouping val="${lineGrp}"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>`;
      break;
    case "scatter":
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${axIds}</c:scatterChart>`;
      break;
    case "radar":
      plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${sers}${axIds}</c:radarChart>`;
      break;
    case "combo":
      plot =
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serList.slice(0, -1).join("")}<c:gapWidth val="219"/><c:overlap val="-27"/>${axIds}</c:barChart>` +
        `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serList.slice(-1).join("")}<c:marker val="1"/>${axIds}</c:lineChart>`;
      break;
    case "pie":
      plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case "doughnut":
      plot = `<c:doughnutChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`;
      break;
  }
  if (!pie) {
    const catPos = c.type === "bar" ? "l" : "b";
    const valPos = c.type === "bar" ? "b" : "l";
    const common = `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`;
    const grids = (c.gridMajor !== false ? `<c:majorGridlines>${GRID_LN}</c:majorGridlines>` : "") + (c.gridMinor ? `<c:minorGridlines>${GRID_LN}</c:minorGridlines>` : "");
    const catTitle = axisTitleXml(c.axisTitles?.cat, c.type === "bar" ? -5400000 : undefined);
    const valTitle = axisTitleXml(c.axisTitles?.val, c.type === "bar" ? undefined : -5400000);
    const valFmt = c.grouping === "percentStacked" ? '<c:numFmt formatCode="0%" sourceLinked="1"/>' : '<c:numFmt formatCode="General" sourceLinked="1"/>';
    // Dağılımda iki değer ekseni vardır.
    const catAx =
      c.type === "scatter"
        ? `<c:valAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${catTitle}<c:numFmt formatCode="General" sourceLinked="1"/>${common}${GRID_LN}${txt(1197)}<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
        : `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/>${catTitle}<c:numFmt formatCode="General" sourceLinked="1"/>${common}${GRID_LN}${txt(1197)}<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    plot +=
      catAx +
      `<c:valAx><c:axId val="${AX2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${grids}${valTitle}${valFmt}${common}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txt(1197)}<c:crossAx val="${AX1}"/><c:crosses val="autoZero"/><c:crossBetween val="${c.type === "scatter" ? "midCat" : "between"}"/></c:valAx>`;
  }
  const title = c.showTitle
    ? `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1862" b="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:r><a:rPr lang="tr-TR"/><a:t>${esc(chartTitle(c))}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = c.legend === "none" ? "" : `<c:legend><c:legendPos val="${c.legend}"/><c:overlay val="0"/>${txt(1197)}</c:legend>`;
  const ext = embedRid ? `<c:externalData r:id="${embedRid}"><c:autoUpdate val="0"/></c:externalData>` : "";
  return (
    XML_HEAD +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:date1904 val="0"/><c:lang val="tr-TR"/><c:roundedCorners val="0"/>${c.style ? `<c:style val="${c.style}"/>` : ""}` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${c.fill ? `<a:solidFill>${colorXml(c.fill, ctx)}</a:solidFill>` : "<a:noFill/>"}<a:ln>${c.line ? `<a:solidFill>${colorXml(c.line, ctx)}</a:solidFill>` : "<a:noFill/>"}</a:ln></c:spPr>${txt(1197)}${ext}</c:chartSpace>`
  );
}

/** Grafik verisini içeren küçük bir çalışma kitabı (SpreadsheetML, satır içi dizgeler). */
export async function chartWorkbook(c: ChartData): Promise<Uint8Array> {
  const zip = new JSZip();
  const ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  zip.file(
    "[Content_Types].xml",
    `${XML_HEAD}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  );
  zip.file("_rels/.rels", `${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `${XML_HEAD}<workbook xmlns="${ss}" xmlns:r="${rel}"><sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  const series = c.type === "pie" || c.type === "doughnut" ? c.series.slice(0, 1) : c.series;
  const strCell = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  const rows: string[] = [];
  rows.push(`<row r="1">${strCell("A1", "")}${series.map((s, j) => strCell(`${colName(j + 1)}1`, s.name)).join("")}</row>`);
  c.categories.forEach((cat, i) => {
    const r = i + 2;
    const cells = series.map((s, j) => {
      const v = s.values[i];
      return v === null || v === undefined || !Number.isFinite(v) ? "" : `<c r="${colName(j + 1)}${r}"><v>${v}</v></c>`;
    });
    rows.push(`<row r="${r}">${strCell(`A${r}`, cat)}${cells.join("")}</row>`);
  });
  zip.file("xl/worksheets/sheet1.xml", `${XML_HEAD}<worksheet xmlns="${ss}"><sheetData>${rows.join("")}</sheetData></worksheet>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
