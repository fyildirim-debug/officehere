/**
 * Mini grafiklerin XLSX içinde saklanması. ECMA-376'nın tanımladığı Özel
 * XML Veri Deposu bölümü (customXml/itemN.xml + itemPropsN.xml) kullanılır;
 * içerik OfficeHere'e özgü ad alanındadır. Diğer uygulamalar bölümü yok
 * sayar, OfficeHere gidiş-dönüşte mini grafikleri geri yükler.
 */
import JSZip from "jszip";
import type { SparklineGroup } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { addr, colToName, parseCellRef, quoteSheetName } from "@officehere/formula";
import { esc, nextRid, relsPathOf, sheetPaths, unescapeXml } from "./ooxmlZip";
import type { PackageParts } from "./xlsxPost";
import { upsertExt } from "./xmlPart";

export const SPARK_NS = "urn:officehere:tablo:sparklines:1";
const REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOM_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CT_CUSTOM_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const NS_DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

const STRING_KEYS = ["type", "color", "markerColor", "negativeColor", "highColor", "lowColor", "firstColor", "lastColor", "axisColor", "minAxis", "maxAxis", "emptyAs"] as const;
const NUMBER_KEYS = ["weight", "minValue", "maxValue"] as const;
const BOOL_KEYS = ["markers", "highPoint", "lowPoint", "firstPoint", "lastPoint", "negative", "showAxis", "showHidden", "rightToLeft"] as const;

function guid(): string {
  const h = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/** Mini grafik tanımlarının XML gösterimi. */
export function sparklinesXml(wb: WorkbookModel): string {
  const sheets = wb.sheets
    .filter((s) => s.sparklines.length)
    .map((s) => {
      const groups = s.sparklines.map((g) => {
        const attrs: string[] = [];
        for (const k of STRING_KEYS) if (g[k] !== undefined) attrs.push(`${k}="${esc(String(g[k]))}"`);
        for (const k of NUMBER_KEYS) if (g[k] !== undefined) attrs.push(`${k}="${g[k]}"`);
        for (const k of BOOL_KEYS) if (g[k]) attrs.push(`${k}="1"`);
        const items = g.items.map((it) => `<sparkline ref="${addr(it.row, it.col)}" data="${esc(it.range)}"/>`).join("");
        return `<group ${attrs.join(" ")}>${items}</group>`;
      });
      return `<sheet name="${esc(s.name)}">${groups.join("")}</sheet>`;
    });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sparklines xmlns="${SPARK_NS}">${sheets.join("")}</sparklines>`;
}

/** Pakete özel XML bölümünü ekler (yazma sonrası işlemde). */
export async function writeSparklinesInto(zip: JSZip, wb: WorkbookModel): Promise<void> {
  if (!wb.sheets.some((s) => s.sparklines.length)) return;
  let n = 1;
  while (zip.file(`customXml/item${n}.xml`)) n++;
  const item = `customXml/item${n}.xml`;
  const props = `customXml/itemProps${n}.xml`;
  zip.file(item, sparklinesXml(wb));
  zip.file(props, `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${guid()}" xmlns:ds="${NS_DS}"><ds:schemaRefs><ds:schemaRef ds:uri="${SPARK_NS}"/></ds:schemaRefs></ds:datastoreItem>`);
  zip.file(relsPathOf(item), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CUSTOM_PROPS}" Target="itemProps${n}.xml"/></Relationships>`);
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  if (!/<Default Extension="xml"/.test(ct)) ct = ct.replace("</Types>", `<Override PartName="/${item}" ContentType="application/xml"/></Types>`);
  ct = ct.replace("</Types>", `<Override PartName="/${props}" ContentType="${CT_CUSTOM_PROPS}"/></Types>`);
  zip.file("[Content_Types].xml", ct);
  let rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  rels = rels.replace("</Relationships>", `<Relationship Id="${nextRid(rels)}" Type="${REL_CUSTOM_XML}" Target="../${item}"/></Relationships>`);
  zip.file("xl/_rels/workbook.xml.rels", rels);
}

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unescapeXml(m[1]) : undefined;
}

/** Özel XML'deki mini grafikleri kitaba yükler. */
export function applySparklinesXml(xml: string, wb: WorkbookModel): void {
  if (!xml.includes(SPARK_NS)) return;
  for (const sm of xml.matchAll(/<sheet\b([^>]*)>([\s\S]*?)<\/sheet>/g)) {
    const name = attr(sm[1], "name");
    const sheet = name ? wb.sheetByName(name) : undefined;
    if (!sheet) continue;
    for (const gm of sm[2].matchAll(/<group\b([^>]*)>([\s\S]*?)<\/group>/g)) {
      const tag = gm[1];
      const g: Record<string, unknown> = { id: wb.newId("spark"), items: [] };
      for (const k of STRING_KEYS) {
        const v = attr(tag, k);
        if (v !== undefined) g[k] = v;
      }
      for (const k of NUMBER_KEYS) {
        const v = attr(tag, k);
        if (v !== undefined && Number.isFinite(Number(v))) g[k] = Number(v);
      }
      for (const k of BOOL_KEYS) if (attr(tag, k) === "1") g[k] = true;
      const group = g as unknown as SparklineGroup;
      if (!["line", "column", "winLoss"].includes(group.type)) group.type = "line";
      if (!group.color) group.color = "#376092";
      for (const im of gm[2].matchAll(/<sparkline\b([^>]*)\/?>/g)) {
        const pos = parseCellRef(attr(im[1], "ref") ?? "");
        const range = attr(im[1], "data");
        if (pos && range) group.items.push({ row: pos.row, col: pos.col, range });
      }
      if (group.items.length) sheet.sparklines.push(group);
    }
  }
}

/** Ham dosya verisinden mini grafikleri okur: önce OfficeHere bölümü, yoksa sayfa uzantısındaki x14:sparklineGroups. */
export async function readSparklinesFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  let found = false;
  for (const path of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(path)) continue;
    const xml = await zip.file(path)!.async("string");
    if (xml.includes(SPARK_NS)) {
      applySparklinesXml(xml, wb);
      found = true;
    }
  }
  if (found || !zip.file("xl/workbook.xml")) return;
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const p = paths.get(sheet.name);
    const xml = p ? await zip.file(p)?.async("string") : undefined;
    if (xml && xml.includes("sparklineGroup")) sheet.sparklines.push(...parseX14Sparklines(xml, wb));
  }
}

/* --------------------- x14:sparklineGroups (sayfa uzantısı) --------------------- */

export const SPARK_EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";

const argb = (c: string | undefined, d: string) => `FF${(c ?? d).replace("#", "").toUpperCase().padStart(6, "0").slice(0, 6)}`;

/** Bir sayfanın mini grafik gruplarının x14 uzantı öğesi (diğer uygulamalar da okur). */
export function x14SparklinesExt(sheetName: string, groups: SparklineGroup[]): string | null {
  if (!groups.length) return null;
  const body = groups
    .map((g) => {
      const a: string[] = [];
      if (g.maxAxis === "custom" && g.maxValue !== undefined) a.push(`manualMax="${g.maxValue}"`);
      if (g.minAxis === "custom" && g.minValue !== undefined) a.push(`manualMin="${g.minValue}"`);
      if (g.weight !== undefined) a.push(`lineWeight="${g.weight}"`);
      if (g.type !== "line") a.push(`type="${g.type === "winLoss" ? "stacked" : "column"}"`);
      a.push(`displayEmptyCellsAs="${g.emptyAs ?? "gap"}"`);
      if (g.markers) a.push('markers="1"');
      if (g.highPoint) a.push('high="1"');
      if (g.lowPoint) a.push('low="1"');
      if (g.firstPoint) a.push('first="1"');
      if (g.lastPoint) a.push('last="1"');
      if (g.negative) a.push('negative="1"');
      if (g.showAxis) a.push('displayXAxis="1"');
      if (g.showHidden) a.push('displayHidden="1"');
      if (g.minAxis && g.minAxis !== "individual") a.push(`minAxisType="${g.minAxis}"`);
      if (g.maxAxis && g.maxAxis !== "individual") a.push(`maxAxisType="${g.maxAxis}"`);
      if (g.rightToLeft) a.push('rightToLeft="1"');
      const colors = `<x14:colorSeries rgb="${argb(g.color, "#376092")}"/><x14:colorNegative rgb="${argb(g.negativeColor, "#D00000")}"/><x14:colorAxis rgb="${argb(g.axisColor, "#000000")}"/><x14:colorMarkers rgb="${argb(g.markerColor, "#D00000")}"/><x14:colorFirst rgb="${argb(g.firstColor, "#D00000")}"/><x14:colorLast rgb="${argb(g.lastColor, "#D00000")}"/><x14:colorHigh rgb="${argb(g.highColor, "#D00000")}"/><x14:colorLow rgb="${argb(g.lowColor, "#D00000")}"/>`;
      const items = g.items
        .map((it) => {
          const f = it.range.includes("!") ? it.range : `${quoteSheetName(sheetName)}!${it.range}`;
          return `<x14:sparkline><xm:f>${esc(f.replace(/\$/g, ""))}</xm:f><xm:sqref>${colToName(it.col)}${it.row + 1}</xm:sqref></x14:sparkline>`;
        })
        .join("");
      return `<x14:sparklineGroup ${a.join(" ")}>${colors}<x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
    })
    .join("");
  return `<ext uri="${SPARK_EXT_URI}" xmlns:x14="${X14_NS}"><x14:sparklineGroups xmlns:xm="${XM_NS}">${body}</x14:sparklineGroups></ext>`;
}

/** Mini grafikleri sayfaların extLst'ine yazar. */
export async function writeX14Sparklines(wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  for (const sheet of wb.sheets) {
    if (!sheet.sparklines.length) continue;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (part) upsertExt(part, SPARK_EXT_URI, x14SparklinesExt(sheet.name, sheet.sparklines));
  }
}

function hexOf(el: string | undefined): string | undefined {
  const m = el ? /\brgb="([0-9A-Fa-f]{8})"/.exec(el) : null;
  return m ? `#${m[1].slice(2).toUpperCase()}` : undefined;
}

/** Sayfa XML'indeki x14:sparklineGroup öğeleri. */
export function parseX14Sparklines(sheetXml: string, wb: WorkbookModel): SparklineGroup[] {
  const out: SparklineGroup[] = [];
  for (const gm of sheetXml.matchAll(/<(?:\w+:)?sparklineGroup\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?sparklineGroup>/g)) {
    const at = (n: string) => attr(gm[1], n);
    const on = (n: string) => at(n) === "1" || at(n) === "true";
    const body = gm[2];
    const color = (tag: string) => hexOf(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>`).exec(body)?.[0]);
    const t = at("type");
    const g: SparklineGroup = { id: wb.newId("spark"), type: t === "column" ? "column" : t === "stacked" ? "winLoss" : "line", items: [], color: color("colorSeries") ?? "#376092" };
    const num = (n: string) => (at(n) !== undefined && Number.isFinite(Number(at(n))) ? Number(at(n)) : undefined);
    if (num("lineWeight") !== undefined) g.weight = num("lineWeight");
    if (on("markers")) g.markers = true;
    if (on("high")) g.highPoint = true;
    if (on("low")) g.lowPoint = true;
    if (on("first")) g.firstPoint = true;
    if (on("last")) g.lastPoint = true;
    if (on("negative")) g.negative = true;
    if (on("displayXAxis")) g.showAxis = true;
    if (on("displayHidden")) g.showHidden = true;
    if (on("rightToLeft")) g.rightToLeft = true;
    const empty = at("displayEmptyCellsAs");
    if (empty === "zero" || empty === "span") g.emptyAs = empty;
    const minT = at("minAxisType");
    const maxT = at("maxAxisType");
    if (minT === "group" || minT === "custom") g.minAxis = minT;
    if (maxT === "group" || maxT === "custom") g.maxAxis = maxT;
    if (num("manualMin") !== undefined) g.minValue = num("manualMin");
    if (num("manualMax") !== undefined) g.maxValue = num("manualMax");
    const pick: [keyof SparklineGroup, string][] = [["negativeColor", "colorNegative"], ["axisColor", "colorAxis"], ["markerColor", "colorMarkers"], ["firstColor", "colorFirst"], ["lastColor", "colorLast"], ["highColor", "colorHigh"], ["lowColor", "colorLow"]];
    for (const [k, tag] of pick) {
      const c = color(tag);
      if (c) (g as unknown as Record<string, unknown>)[k] = c;
    }
    for (const sm of body.matchAll(/<(?:\w+:)?sparkline>([\s\S]*?)<\/(?:\w+:)?sparkline>/g)) {
      const f = unescapeXml(/<(?:\w+:)?f>([^<]*)</.exec(sm[1])?.[1] ?? "").trim();
      const pos = parseCellRef((/<(?:\w+:)?sqref>([^<]*)</.exec(sm[1])?.[1] ?? "").trim());
      if (f && pos) g.items.push({ row: pos.row, col: pos.col, range: f });
    }
    if (g.items.length) out.push(g);
  }
  return out;
}

