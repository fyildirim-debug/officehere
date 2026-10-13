/**
 * Dinamik dizi formülleri: dosyada dizi formülü (t="array") olarak yazılır ve
 * hücreye "cm" (hücre üst verisi) eklenir; xl/metadata.xml'deki XLDAPR türü
 * formülün taşan dinamik dizi olduğunu bildirir ([MS-XLSX] dynamicArrayProperties).
 * Üst veri yazılmazsa masaüstü uygulamalar formülü eski tip {=...} dizi formülü
 * olarak gösterir ve sonuç boyutu değişince yeniden taşmaz.
 */
import type JSZip from "jszip";
import { EN_SYNTAX, tokenize } from "@officehere/formula";
import type { WorkbookModel } from "../model/workbook";
import type { PackageParts } from "./xlsxPost";
import { nextRid } from "./ooxmlZip";

/** Sonucu dizi olabilen (dinamik dizi hesaplamasına göre yazılması gereken) işlevler. */
const DYNAMIC_FUNCS = new Set([
  "SORT", "SORTBY", "FILTER", "UNIQUE", "SEQUENCE", "RANDARRAY", "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS",
  "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND", "TEXTSPLIT", "MAP", "BYROW", "BYCOL", "SCAN", "REDUCE", "MAKEARRAY",
  "LET", "LAMBDA", "XLOOKUP", "XMATCH", "REGEXEXTRACT", "GROUPBY", "PIVOTBY",
]);

/** Formül dinamik dizi işlevi içeriyor mu (sonucu tek hücre olsa da dinamik dizi olarak yazılır)? */
export function usesDynamicFunctions(formula: string): boolean {
  for (const t of tokenize(formula.startsWith("=") ? formula.slice(1) : formula, EN_SYNTAX)) if (t.type === "func" && DYNAMIC_FUNCS.has(String(t.value).toUpperCase())) return true;
  return false;
}

const METADATA_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
  "</metadata>";
const CT_METADATA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const REL_METADATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";

/** Dizi formüllü hücrelere cm="1" ekler; en az bir tane varsa üst veri parçasını pakete koyar. */
export async function writeDynamicArrayMeta(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  let any = false;
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    const data = part?.get("sheetData");
    if (!part || !data || !data.includes('t="array"')) continue;
    const next = data.replace(/<c\b([^>]*?)(\/?)>(?=<f\b[^>]*\bt="array")/g, (m0, attrs: string, slash: string) => {
      if (/\bcm="/.test(attrs)) return m0;
      any = true;
      return `<c${attrs} cm="1"${slash}>`;
    });
    part.set("sheetData", next);
  }
  if (!any || zip.file("xl/metadata.xml")) return;
  zip.file("xl/metadata.xml", METADATA_XML);
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  if (!ct.includes('PartName="/xl/metadata.xml"')) ct = ct.replace("</Types>", `<Override PartName="/xl/metadata.xml" ContentType="${CT_METADATA}"/></Types>`);
  zip.file("[Content_Types].xml", ct);
  let rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  if (!rels.includes(REL_METADATA)) rels = rels.replace("</Relationships>", `<Relationship Id="${nextRid(rels)}" Type="${REL_METADATA}" Target="metadata.xml"/></Relationships>`);
  zip.file("xl/_rels/workbook.xml.rels", rels);
}
