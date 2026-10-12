/**
 * Mini grafiklerin XLSX içinde saklanması. ECMA-376'nın tanımladığı Özel
 * XML Veri Deposu bölümü (customXml/itemN.xml + itemPropsN.xml) kullanılır;
 * içerik OfficeHere'e özgü ad alanındadır. Diğer uygulamalar bölümü yok
 * sayar, OfficeHere gidiş-dönüşte mini grafikleri geri yükler.
 */
import JSZip from "jszip";
import type { SparklineGroup } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { addr, parseCellRef } from "@officehere/formula";
import { esc, nextRid, relsPathOf, unescapeXml } from "./ooxmlZip";

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

/** Ham dosya verisinden mini grafikleri okur. */
export async function readSparklinesFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  for (const path of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(path)) continue;
    const xml = await zip.file(path)!.async("string");
    if (xml.includes(SPARK_NS)) applySparklinesXml(xml, wb);
  }
}
