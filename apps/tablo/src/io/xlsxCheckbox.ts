/**
 * Hücre içi onay kutuları dosyada bir hücre biçimi özelliği olarak durur ([MS-XLSX]
 * özellik torbaları): xl/featurePropertyBag/featurePropertyBag.xml'de "Checkbox"
 * torbası, styles.xml'de onay kutulu hücre biçimlerinin (xf) extLst'inde
 * xfpb:xfComplement. Hücrenin değeri DOĞRU/YANLIŞ olarak yazılır.
 */
import JSZip from "jszip";
import { nameToCol } from "@officehere/formula";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { nextRid, resolveTarget, sheetPaths } from "./ooxmlZip";
import type { PackageParts } from "./xlsxPost";
import { attrOf, childElements, rebuildElement } from "./xmlPart";

const FPB_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag";
const XF_EXT_URI = "{C7286773-470A-42A8-94C5-96B5CB345126}";
const REL_FPB = "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag";
const CT_FPB = "application/vnd.ms-excel.featurepropertybag+xml";
const FPB_PATH = "xl/featurePropertyBag/featurePropertyBag.xml";
const FPB_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<FeaturePropertyBags xmlns="${FPB_NS}">` +
  '<bag type="Checkbox"/>' +
  '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>' +
  '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>' +
  '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>' +
  "</FeaturePropertyBags>";
const XF_EXT = `<extLst><ext uri="${XF_EXT_URI}" xmlns:xfpb="${FPB_NS}"><xfpb:xfComplement i="0"/></ext></extLst>`;

function hasCheckbox(wb: WorkbookModel, id: number | undefined): boolean {
  return id !== undefined && !!wb.styles.get(id).checkbox;
}

/** "B12" → [satır, sütun] (0 tabanlı). */
function parseA1(ref: string): [number, number] | null {
  const m = /^([A-Z]{1,3})(\d+)$/.exec(ref);
  return m ? [Number(m[2]) - 1, nameToCol(m[1])] : null;
}

/** Biçimin (xf) onay kutulu kopyası: extLst en sona eklenir. */
function withCheckbox(xf: string): string {
  if (/\/>\s*$/.test(xf)) return xf.replace(/\s*\/>\s*$/, `>${XF_EXT}</xf>`);
  return xf.replace(/<\/(\w+:)?xf>\s*$/, `${XF_EXT}</$1xf>`);
}

/** Onay kutulu hücre/sütun/satırların biçimlerini işaretli kopyalarına yönlendirir ve özellik torbasını yazar. */
export async function writeCheckboxesInto(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  const used = (s: SheetModel) => [...s.styles.values(), ...s.colStyles.values(), ...s.rowStyles.values()].some((id) => hasCheckbox(wb, id));
  if (!wb.sheets.some(used)) return;
  const stylesFile = zip.file("xl/styles.xml");
  if (!stylesFile) return;
  const stylesXml = await stylesFile.async("string");
  const cellXfs = /<(\w+:)?cellXfs\b[^>]*>[\s\S]*?<\/(\w+:)?cellXfs>/.exec(stylesXml)?.[0];
  if (!cellXfs) return;
  const xfs = childElements(cellXfs).filter((c) => c.name === "xf").map((c) => c.xml);
  const added: string[] = [];
  const variant = new Map<number, number>();
  const boxed = (s: number) => {
    let v = variant.get(s);
    if (v === undefined) {
      v = xfs.length + added.length;
      added.push(withCheckbox(xfs[s] ?? xfs[0]));
      variant.set(s, v);
    }
    return v;
  };
  const setS = (tag: string, attr: string, value: number) => (new RegExp(`\\b${attr}="\\d+"`).test(tag) ? tag.replace(new RegExp(`\\b${attr}="\\d+"`), `${attr}="${value}"`) : tag.replace(/^(<\w+(?::\w+)?)/, `$1 ${attr}="${value}"`));
  for (const sheet of wb.sheets) {
    if (!used(sheet)) continue;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (!part) continue;
    const data = part.get("sheetData");
    if (data) {
      part.set(
        "sheetData",
        data
          .replace(/<c\b[^>]*?\/?>/g, (tag) => {
            const pos = parseA1(attrOf(tag, "r") ?? "");
            if (!pos || !wb.styleAt(sheet, pos[0], pos[1]).checkbox) return tag;
            return setS(tag, "s", boxed(Number(attrOf(tag, "s") ?? 0)));
          })
          .replace(/<row\b[^>]*?\/?>/g, (tag) => {
            const r = Number(attrOf(tag, "r") ?? 0) - 1;
            if (!hasCheckbox(wb, sheet.rowStyles.get(r)) || attrOf(tag, "s") === undefined) return tag;
            return setS(tag, "s", boxed(Number(attrOf(tag, "s"))));
          }),
      );
    }
    const cols = part.get("cols");
    if (cols) {
      part.set(
        "cols",
        cols.replace(/<col\b[^>]*?\/?>/g, (tag) => {
          const min = Number(attrOf(tag, "min") ?? 0) - 1;
          if (!hasCheckbox(wb, sheet.colStyles.get(min)) || attrOf(tag, "style") === undefined) return tag;
          return setS(tag, "style", boxed(Number(attrOf(tag, "style"))));
        }),
      );
    }
  }
  if (!added.length) return;
  const all = [...xfs, ...added];
  const next = rebuildElement(cellXfs, all).replace(/(<(?:\w+:)?cellXfs\b[^>]*?)\s+count="\d+"/, "$1").replace(/^<((?:\w+:)?cellXfs)\b/, `<$1 count="${all.length}"`);
  zip.file("xl/styles.xml", stylesXml.replace(cellXfs, next));
  zip.file(FPB_PATH, FPB_XML);
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  if (!ct.includes(`PartName="/${FPB_PATH}"`)) ct = ct.replace("</Types>", `<Override PartName="/${FPB_PATH}" ContentType="${CT_FPB}"/></Types>`);
  zip.file("[Content_Types].xml", ct);
  let rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  if (!rels.includes(REL_FPB)) rels = rels.replace("</Relationships>", `<Relationship Id="${nextRid(rels)}" Type="${REL_FPB}" Target="featurePropertyBag/featurePropertyBag.xml"/></Relationships>`);
  zip.file("xl/_rels/workbook.xml.rels", rels);
}

/** Dosyadaki onay kutulu biçimleri (xfComplement → Checkbox torbası) hücre stillerine işler. */
export async function readCheckboxesFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  const styles = await zip.file("xl/styles.xml")?.async("string");
  if (!styles || !/xfComplement/.test(styles)) return;
  // Özellik torbasında Checkbox olmalı (başka hücre denetimleri onay kutusu sayılmaz)
  const wbRels = (await zip.file("xl/_rels/workbook.xml.rels")?.async("string")) ?? "";
  const rel = [...wbRels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((t) => /\/FeaturePropertyBag$/i.test(attrOf(t, "Type") ?? ""));
  const target = rel ? attrOf(rel, "Target") : undefined;
  const fpb = target ? await zip.file(resolveTarget("xl/workbook.xml", target))?.async("string") : undefined;
  if (!fpb || !/type="Checkbox"/.test(fpb)) return;
  const cellXfs = /<(\w+:)?cellXfs\b[^>]*>[\s\S]*?<\/(\w+:)?cellXfs>/.exec(styles)?.[0] ?? "";
  const boxed = new Set<number>();
  childElements(cellXfs)
    .filter((c) => c.name === "xf")
    .forEach((c, i) => {
      if (/xfComplement\b/.test(c.xml)) boxed.add(i);
    });
  if (!boxed.size) return;
  const paths = await sheetPaths(zip);
  const mark = (id: number | undefined) => wb.styles.patch(id, { checkbox: true });
  for (const sheet of wb.sheets) {
    const p = paths.get(sheet.name);
    const xml = p ? await zip.file(p)?.async("string") : undefined;
    if (!xml) continue;
    for (const m of xml.matchAll(/<(?:\w+:)?c\b[^>]*?\/?>/g)) {
      const s = attrOf(m[0], "s");
      if (s === undefined || !boxed.has(Number(s))) continue;
      const pos = parseA1(attrOf(m[0], "r") ?? "");
      if (!pos) continue;
      wb.setStyleRaw(sheet, pos[0], pos[1], mark(wb.styleIdAt(sheet, pos[0], pos[1])));
    }
    for (const m of xml.matchAll(/<(?:\w+:)?col\b[^>]*?\/?>/g)) {
      const s = attrOf(m[0], "style");
      if (s === undefined || !boxed.has(Number(s))) continue;
      for (let c = Number(attrOf(m[0], "min")) - 1; c <= Number(attrOf(m[0], "max")) - 1 && c < 16384; c++) sheet.colStyles.set(c, mark(sheet.colStyles.get(c)));
    }
    for (const m of xml.matchAll(/<(?:\w+:)?row\b[^>]*?\/?>/g)) {
      const s = attrOf(m[0], "s");
      if (s === undefined || !boxed.has(Number(s)) || attrOf(m[0], "customFormat") !== "1") continue;
      const r = Number(attrOf(m[0], "r")) - 1;
      sheet.rowStyles.set(r, mark(sheet.rowStyles.get(r)));
    }
  }
}

