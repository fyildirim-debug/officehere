/** Tanımlı adların workbook.xml'e yazılması (sabit ve formül adları dahil; exceljs yalnız aralık yazar). */
import type { WorkbookModel } from "../model/workbook";
import { esc } from "./ooxmlZip";
import { addXlfnPrefixes } from "./xlsxFormula";
import type { PackageParts } from "./xlsxPost";

export async function writeDefinedNamesInto(wb: WorkbookModel, parts: PackageParts): Promise<void> {
  const part = await parts.workbook();
  if (!part) return;
  const names = wb.engine.getNames();
  // exceljs'in yazdığı yerleşik adlar (_xlnm.*) korunur, kullanıcı adları yeniden yazılır
  const builtins = [...(part.get("definedNames") ?? "").matchAll(/<definedName\b[^>]*>[^<]*<\/definedName>/g)].map((m) => m[0]).filter((t) => /name="_xlnm\./.test(t));
  const user = names.map((n) => {
    const idx = n.sheet !== undefined ? wb.sheets.findIndex((s) => s.id === n.sheet) : -1;
    return `<definedName name="${esc(n.name)}"${idx >= 0 ? ` localSheetId="${idx}"` : ""}>${esc(addXlfnPrefixes(n.formula.replace(/^=/, "")))}</definedName>`;
  });
  const all = [...builtins, ...user];
  // Yeri (sheets → externalReferences → definedNames → calcPr) şema sırası belirler
  part.set("definedNames", all.length ? `<definedNames>${all.join("")}</definedNames>` : null);
}
