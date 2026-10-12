/** Tanımlı adların workbook.xml'e yazılması (sabit ve formül adları dahil; exceljs yalnız aralık yazar). */
import type JSZip from "jszip";
import type { WorkbookModel } from "../model/workbook";
import { esc } from "./ooxmlZip";
import { addXlfnPrefixes } from "./xlsxFormula";

export async function writeDefinedNamesInto(zip: JSZip, wb: WorkbookModel): Promise<void> {
  const f = zip.file("xl/workbook.xml");
  if (!f) return;
  let xml = await f.async("string");
  const names = wb.engine.getNames();
  // exceljs'in yazdığı yerleşik adlar (_xlnm.*) korunur, kullanıcı adları yeniden yazılır
  const builtins = [...(/<definedNames>([\s\S]*?)<\/definedNames>/.exec(xml)?.[1] ?? "").matchAll(/<definedName\b[^>]*>[^<]*<\/definedName>/g)].map((m) => m[0]).filter((t) => /name="_xlnm\./.test(t));
  const user = names.map((n) => {
    const idx = n.sheet !== undefined ? wb.sheets.findIndex((s) => s.id === n.sheet) : -1;
    return `<definedName name="${esc(n.name)}"${idx >= 0 ? ` localSheetId="${idx}"` : ""}>${esc(addXlfnPrefixes(n.formula.replace(/^=/, "")))}</definedName>`;
  });
  const all = [...builtins, ...user];
  xml = xml.replace(/<definedNames>[\s\S]*?<\/definedNames>|<definedNames\/>/, "");
  if (all.length) {
    const block = `<definedNames>${all.join("")}</definedNames>`;
    // ECMA-376 sırası: sheets → functionGroups → externalReferences → definedNames → calcPr
    xml = /<\/sheets>/.test(xml) ? xml.replace("</sheets>", `</sheets>${block}`) : xml;
    // externalReferences sheets'ten sonra gelirse definedNames onun arkasına taşınır
    xml = xml.replace(/(<definedNames>[\s\S]*?<\/definedNames>)(<externalReferences>[\s\S]*?<\/externalReferences>)/, "$2$1");
  }
  zip.file("xl/workbook.xml", xml);
}
