/**
 * Dilimleyicilerin XLSX'te saklanması: ECMA-376 Özel XML Veri Deposu bölümü
 * (customXml/itemN.xml), OfficeHere ad alanında. Tablo ve özet tablo ADIYLA
 * başvurulur. Diğer uygulamalar bölümü yok sayar; filtre durumu (tablo
 * otomatik filtresi, özet tablo gizli öğeleri) dosyada ayrıca korunur.
 */
import JSZip from "jszip";
import type { SlicerDef } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { esc, nextRid, relsPathOf, unescapeXml } from "./ooxmlZip";

export const SLICER_NS = "urn:officehere:tablo:slicers:1";
const REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOM_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CT_CUSTOM_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const NS_DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

type Stored = Omit<SlicerDef, "id" | "source"> & { source: { kind: "table"; table: string; column: string } | { kind: "pivot"; pivot: string; field: number } };

export function slicersXml(wb: WorkbookModel): string | null {
  const tableName = (id: string) => wb.sheets.flatMap((s) => s.tables).find((t) => t.id === id)?.name;
  const pivotName = (id: string) => wb.sheets.flatMap((s) => s.pivots).find((p) => p.id === id)?.name;
  const parts: string[] = [];
  for (const s of wb.sheets) {
    const items: string[] = [];
    for (const sl of s.slicers) {
      const src = sl.source.kind === "table" ? { kind: "table" as const, table: tableName(sl.source.tableId), column: sl.source.column } : { kind: "pivot" as const, pivot: pivotName(sl.source.pivotId), field: sl.source.field };
      if ((src.kind === "table" && !src.table) || (src.kind === "pivot" && !src.pivot)) continue;
      const { id: _id, source: _src, ...rest } = sl;
      void _id;
      void _src;
      items.push(`<slicer>${esc(JSON.stringify({ ...rest, source: src }))}</slicer>`);
    }
    if (items.length) parts.push(`<sheet name="${esc(s.name)}">${items.join("")}</sheet>`);
  }
  return parts.length ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${SLICER_NS}">${parts.join("")}</slicers>` : null;
}

export async function writeSlicersInto(zip: JSZip, wb: WorkbookModel): Promise<void> {
  const xml = slicersXml(wb);
  if (!xml) return;
  let n = 1;
  while (zip.file(`customXml/item${n}.xml`)) n++;
  const item = `customXml/item${n}.xml`;
  const props = `customXml/itemProps${n}.xml`;
  zip.file(item, xml);
  const guid = `{${crypto.randomUUID().toUpperCase()}}`;
  zip.file(props, `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${guid}" xmlns:ds="${NS_DS}"><ds:schemaRefs><ds:schemaRef ds:uri="${SLICER_NS}"/></ds:schemaRefs></ds:datastoreItem>`);
  zip.file(relsPathOf(item), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CUSTOM_PROPS}" Target="itemProps${n}.xml"/></Relationships>`);
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  if (!/<Default Extension="xml"/.test(ct)) ct = ct.replace("</Types>", `<Override PartName="/${item}" ContentType="application/xml"/></Types>`);
  ct = ct.replace("</Types>", `<Override PartName="/${props}" ContentType="${CT_CUSTOM_PROPS}"/></Types>`);
  zip.file("[Content_Types].xml", ct);
  let rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  rels = rels.replace("</Relationships>", `<Relationship Id="${nextRid(rels)}" Type="${REL_CUSTOM_XML}" Target="../${item}"/></Relationships>`);
  zip.file("xl/_rels/workbook.xml.rels", rels);
}

/** Özel XML'deki dilimleyicileri kitaba yükler (tablo/özet tablo adları kimliğe çevrilir). */
export function applySlicersXml(xml: string, wb: WorkbookModel): void {
  if (!xml.includes(SLICER_NS)) return;
  const up = (s: string) => s.toLocaleUpperCase("tr-TR");
  for (const sm of xml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"[^>]*>([\s\S]*?)<\/sheet>/g)) {
    const sheet = wb.sheetByName(unescapeXml(sm[1]));
    if (!sheet) continue;
    for (const im of sm[2].matchAll(/<slicer>([\s\S]*?)<\/slicer>/g)) {
      try {
        const st = JSON.parse(unescapeXml(im[1])) as Stored;
        let source: SlicerDef["source"] | null = null;
        if (st.source.kind === "table") {
          const name = st.source.table;
          const t = wb.sheets.flatMap((s) => s.tables).find((x) => up(x.name) === up(name));
          if (t) source = { kind: "table", tableId: t.id, column: st.source.column };
        } else {
          const name = st.source.pivot;
          const p = wb.sheets.flatMap((s) => s.pivots).find((x) => up(x.name) === up(name));
          if (p) source = { kind: "pivot", pivotId: p.id, field: st.source.field };
        }
        if (!source) continue;
        sheet.slicers.push({ ...st, id: wb.newId("slicer"), source, selected: Array.isArray(st.selected) ? st.selected : [] });
      } catch {
        /* bozuk kayıt atlanır */
      }
    }
  }
}

export async function readSlicersFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  for (const path of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(path)) continue;
    const xml = await zip.file(path)!.async("string");
    if (xml.includes(SLICER_NS)) applySlicersXml(xml, wb);
  }
}
