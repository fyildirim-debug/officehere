/**
 * exceljs çıktısının son işlenmesi. Çalışma sayfası ve çalışma kitabı
 * XML'ine eklenen tüm öğeler (çizim, sayfa sonları, uzantılar, özet tablo
 * önbellekleri, tanımlı adlar, koruma…) XmlPart üzerinden yerleştirilir ve
 * parçalar yazılırken şemadaki sıraya dizilir (exceljs'in kendi yazdığı
 * öğeler de: ör. legacyDrawing'i extLst'ten sonra yazar).
 */
import JSZip from "jszip";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { sheetPaths } from "./ooxmlZip";
import { WORKBOOK_ORDER, WORKSHEET_ORDER, XmlPart, attrOf, childElements, getExt, rebuildElement, upsertExt } from "./xmlPart";
import { writeChartsInto } from "./xlsxCharts";
import { writePivotsInto } from "./xlsxPivot";
import { writeSparklinesInto, writeX14Sparklines } from "./xlsxSpark";
import { writePageBreaksInto } from "./xlsxBreaks";
import { writeDefinedNamesInto } from "./xlsxNames";
import { writeSheetExtras, writeWorkbookExtras } from "./xlsxExtras";
import { writeDataTablesInto } from "./xlsxWhatIf";
import { writeSlicersInto } from "./xlsxSlicers";
import { writeDynamicArrayMeta } from "./xlsxDynamic";
import { writeCheckboxesInto } from "./xlsxCheckbox";
import { patchThemeXml } from "../model/theme";

/** Paketteki XML parçalarının önbelleği: her parça bir kez ayrıştırılır, sonda sırasıyla yazılır. */
export class PackageParts {
  private cache = new Map<string, XmlPart>();
  constructor(readonly zip: JSZip) {}

  async part(path: string, order: readonly string[] = WORKSHEET_ORDER): Promise<XmlPart | null> {
    const hit = this.cache.get(path);
    if (hit) return hit;
    const f = this.zip.file(path);
    if (!f) return null;
    const p = new XmlPart(await f.async("string"), order);
    this.cache.set(path, p);
    return p;
  }

  workbook(): Promise<XmlPart | null> {
    return this.part("xl/workbook.xml", WORKBOOK_ORDER);
  }

  flush(): void {
    for (const [path, p] of this.cache) this.zip.file(path, p.toString());
  }
}

/** İçeriği olmayan ama gizli/yüksekliği özel/anahat düzeyli satırlar sheetData'ya eklenir. */
function fixRows(part: XmlPart, sheet: SheetModel): void {
  const rows = new Set<number>([...sheet.rows.customItems().map(([r]) => r), ...sheet.rows.hiddenItems(), ...sheet.rows.levelItems().map(([r]) => r)]);
  if (!rows.size) return;
  let xml = part.get("sheetData") ?? "<sheetData></sheetData>";
  // İçeriğe göre (otomatik) yükseklikteki satırlar customHeight taşımaz
  xml = xml.replace(/<row r="(\d+)"([^>]*?) customHeight="1"/g, (m0, r: string, rest: string) => (sheet.rows.isManual(Number(r) - 1) ? m0 : `<row r="${r}"${rest}`));
  const existing = new Set<number>([...xml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]) - 1));
  const missing = [...rows].filter((r) => !existing.has(r)).sort((a, b) => a - b);
  if (missing.length) {
    if (/^<sheetData\s*\/>$/.test(xml)) xml = "<sheetData></sheetData>";
    // Mevcut satır başlangıçları bir kez taranır; eksikler sırayla araya yerleştirilir
    const starts = [...xml.matchAll(/<row r="(\d+)"/g)].map((m) => ({ r: Number(m[1]) - 1, at: m.index! }));
    const endAt = xml.lastIndexOf("</sheetData>");
    const pieces: string[] = [];
    let last = 0;
    let mi = 0;
    const tagFor = (r: number) => `<row r="${r + 1}"${sheet.rows.isCustom(r) ? ` ht="${Math.round(((sheet.rows.rawSize(r) * 72) / 96) * 100) / 100}"${sheet.rows.isManual(r) ? ' customHeight="1"' : ""}` : ""}${sheet.rows.isHidden(r) ? ' hidden="1"' : ""}${sheet.rows.level(r) ? ` outlineLevel="${sheet.rows.level(r)}"` : ""}/>`;
    for (const s of starts) {
      let add = "";
      while (mi < missing.length && missing[mi] < s.r) add += tagFor(missing[mi++]);
      if (add) {
        pieces.push(xml.slice(last, s.at), add);
        last = s.at;
      }
    }
    let tail = "";
    while (mi < missing.length) tail += tagFor(missing[mi++]);
    pieces.push(xml.slice(last, endAt), tail, xml.slice(endAt));
    xml = pieces.join("");
  }
  part.set("sheetData", xml);
}

const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
export const CF_EXT_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
const RULE_EXT_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

export function guid(): string {
  const h = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/**
 * Veri çubukları: kuraldaki x14:id ile sayfa düzeyindeki x14:conditionalFormatting
 * eşleştirilir (2010 sonrası çubuk görünümü: kenarlık, negatif değer rengi, eksen).
 * exceljs boş <x14:id/> yazar ve sayfa uzantısını hiç yazmaz.
 */
export function fixDataBars(part: XmlPart): void {
  const entries: string[] = [];
  part.map("conditionalFormatting", (cf) => {
    if (!/type="dataBar"/.test(cf)) return cf;
    const sqref = attrOf(cf, "sqref") ?? "";
    const rules = childElements(cf).map((c) => {
      if (c.name !== "cfRule" || attrOf(c.xml, "type") !== "dataBar") return c.xml;
      const id = guid();
      const color = /<(?:\w+:)?color\b[^>]*\brgb="([0-9A-Fa-f]{8})"/.exec(c.xml)?.[1] ?? "FF638EC6";
      const inner = childElements(c.xml).filter((k) => k.name !== "extLst").map((k) => k.xml);
      inner.push(`<extLst><ext uri="${RULE_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>${id}</x14:id></ext></extLst>`);
      entries.push(
        `<x14:conditionalFormatting xmlns:xm="${XM_NS}"><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="${color}"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`,
      );
      return rebuildElement(c.xml, inner);
    });
    return rebuildElement(cf, rules);
  });
  if (!entries.length) return;
  // exceljs'in yazdığı (ör. 3 Üçgen simge kümesi) koşullu biçim uzantısıyla birleştirilir
  const ext = getExt(part, CF_EXT_URI);
  let list = ext ? childElements(ext).find((c) => c.name === "conditionalFormattings")?.xml : undefined;
  const items = list ? [...childElements(list).map((c) => c.xml), ...entries] : entries;
  list = list ? rebuildElement(list, items) : `<x14:conditionalFormattings>${items.join("")}</x14:conditionalFormattings>`;
  const others = ext ? childElements(ext).filter((c) => c.name !== "conditionalFormattings").map((c) => c.xml) : [];
  const extXml = ext ? rebuildElement(ext, [list, ...others]) : `<ext uri="${CF_EXT_URI}" xmlns:x14="${X14_NS}">${list}</ext>`;
  upsertExt(part, CF_EXT_URI, extXml);
}

/** exceljs çıktısını son işler: satırlar, çizimler, uzantılar, sayfa sonları, adlar ve şema sırası. */
export async function postProcess(bytes: Uint8Array, wb: WorkbookModel): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const paths = await sheetPaths(zip);
  const parts = new PackageParts(zip);
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (!part) continue;
    fixRows(part, sheet);
    fixDataBars(part);
  }
  if (wb.sheets.some((s) => s.charts.length)) await writeChartsInto(zip, wb, paths, parts);
  if (wb.sheets.some((s) => s.pivots.length)) await writePivotsInto(zip, wb, paths, parts);
  await writeSparklinesInto(zip, wb);
  await writeSlicersInto(zip, wb);
  await writeX14Sparklines(wb, paths, parts);
  await writePageBreaksInto(wb, paths, parts);
  await writeSheetExtras(wb, paths, parts);
  await writeDataTablesInto(wb, paths, parts);
  await writeDynamicArrayMeta(zip, wb, paths, parts);
  await writeCheckboxesInto(zip, wb, paths, parts);
  await writeDefinedNamesInto(wb, parts);
  await writeWorkbookExtras(wb, parts);
  parts.flush();
  // Çalışma kitabı teması (renk ve yazı tipi şemaları)
  const themeFile = zip.file("xl/theme/theme1.xml");
  if (themeFile) zip.file("xl/theme/theme1.xml", patchThemeXml(await themeFile.async("string"), wb.theme));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
