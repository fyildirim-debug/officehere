/**
 * Sayfa ve çalışma kitabı XML'ine doğrudan yazılan ek öğeler (exceljs'in
 * yazmadıkları ya da eksik yazdıkları): veri doğrulamaları (çakışmasız, sınırsız
 * aralık), tablo sütunlarının hesaplanmış sütun ve toplam satırı bilgisi,
 * senaryolar (<scenarios>), çalışma kitabı yapı koruması (<workbookProtection>).
 * Yerleri XmlPart'ın şema sırasıyla belirlenir.
 */
import { colToName, shiftFormula } from "@officehere/formula";
import type { Rect, TableDef, ValidationRule } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { subtractRects } from "../model/rect";
import { esc, relsPathOf, resolveTarget } from "./ooxmlZip";
import { addXlfnPrefixes } from "./xlsxFormula";
import type { PackageParts } from "./xlsxPost";
import { writeScenariosInto, writeWorkbookProtection } from "./xlsxWhatIf";
import { autoFilterXml, patchTableAutoFilter } from "./xlsxFilters";

/* ------------------------------ Veri doğrulama ------------------------------ */

function dvFormula(v: string | undefined, type: ValidationRule["type"]): string | undefined {
  if (v === undefined || v === "") return undefined;
  if (v.startsWith("=")) return addXlfnPrefixes(v.slice(1));
  if (type === "list") return `"${v.split(";").map((x) => x.trim()).join(",")}"`;
  const n = Number(v.replace(",", "."));
  return Number.isFinite(n) && /^-?\d+([.,]\d+)?$/.test(v) ? String(n) : v;
}

/** sqref için açık A1 aralığı (tam sütun/satır "F:F" yerine "F1:F1048576"). */
function sqrefOf(r: Rect): string {
  const a = `${colToName(r.c1)}${r.r1 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${colToName(r.c2)}${r.r2 + 1}`;
}

/** Tek doğrulama kuralının XML'i. */
export function dataValidationXml(rule: ValidationRule, ranges: Rect[]): string {
  const a: string[] = [];
  if (rule.type !== "any") a.push(`type="${rule.type}"`);
  if (rule.errorStyle && rule.errorStyle !== "stop") a.push(`errorStyle="${rule.errorStyle}"`);
  if (rule.operator && rule.type !== "list" && rule.type !== "custom" && rule.type !== "any" && rule.operator !== "between") a.push(`operator="${rule.operator}"`);
  if (rule.allowBlank !== false) a.push('allowBlank="1"');
  // OOXML'de showDropDown="1" açılır listeyi GİZLER
  if (rule.type === "list" && rule.showDropdown === false) a.push('showDropDown="1"');
  if (rule.prompt || rule.promptTitle) a.push('showInputMessage="1"');
  if (rule.type !== "any") a.push('showErrorMessage="1"');
  if (rule.errorTitle) a.push(`errorTitle="${esc(rule.errorTitle)}"`);
  if (rule.error) a.push(`error="${esc(rule.error)}"`);
  if (rule.promptTitle) a.push(`promptTitle="${esc(rule.promptTitle)}"`);
  if (rule.prompt) a.push(`prompt="${esc(rule.prompt)}"`);
  a.push(`sqref="${ranges.map(sqrefOf).join(" ")}"`);
  const f1 = rule.type === "any" ? undefined : dvFormula(rule.formula1, rule.type);
  const f2 = rule.type === "any" || rule.type === "list" ? undefined : dvFormula(rule.formula2, rule.type);
  return `<dataValidation ${a.join(" ")}>${f1 !== undefined ? `<formula1>${esc(f1)}</formula1>` : ""}${f2 !== undefined ? `<formula2>${esc(f2)}</formula2>` : ""}</dataValidation>`;
}

/** Sayfanın doğrulamaları: sonra eklenen kural çakışan hücrelerde öncekinin yerine geçer (aralıklar çakışmaz). */
export function dataValidationsXml(sheet: SheetModel): string | null {
  const list = sheet.validations.map((v) => ({ rule: v.rule, ranges: v.ranges.map((r) => ({ ...r })) }));
  for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++) list[i].ranges = subtractRects(list[i].ranges, list[j].ranges);
  const items = list.filter((v) => v.ranges.length).map((v) => dataValidationXml(v.rule, v.ranges));
  return items.length ? `<dataValidations count="${items.length}">${items.join("")}</dataValidations>` : null;
}

/* ------------------------------ Tablo sütunları ------------------------------ */

const TOTAL_FN: Record<number, string> = { 101: "average", 102: "countNums", 103: "count", 104: "max", 105: "min", 107: "stdDev", 109: "sum", 110: "var" };

/** Sütun hesaplanmış sütun mu: veri satırlarının çoğu aynı göreli formülü taşıyorsa ilk satırdaki biçimi. */
export function calculatedColumnFormula(wb: WorkbookModel, sheet: SheetModel, t: TableDef, col: number): string | null {
  const d1 = t.ref.r1 + (t.headerRow ? 1 : 0);
  const d2 = Math.min(t.ref.r2 - (t.totalRow ? 1 : 0), d1 + 5000);
  if (d2 < d1) return null;
  const counts = new Map<string, number>();
  for (let r = d1; r <= d2; r++) {
    const f = wb.getFormula(sheet, r, col);
    if (!f) continue;
    const base = r === d1 ? f : shiftFormula(f, d1 - r, 0);
    counts.set(base, (counts.get(base) ?? 0) + 1);
  }
  let best: [string, number] | null = null;
  for (const e of counts) if (!best || e[1] > best[1]) best = e;
  if (!best || best[0].includes("#REF!") || best[1] < Math.max(1, Math.ceil((d2 - d1 + 1) * 0.6))) return null;
  return best[0];
}

function tableColumnsXml(wb: WorkbookModel, sheet: SheetModel, t: TableDef): string {
  const cols = t.columns.map((name, i) => {
    const c = t.ref.c1 + i;
    const a = [`id="${i + 1}"`, `name="${esc(name)}"`];
    let inner = "";
    const calc = calculatedColumnFormula(wb, sheet, t, c);
    if (calc) inner += `<calculatedColumnFormula>${esc(addXlfnPrefixes(calc.replace(/^=/, "")))}</calculatedColumnFormula>`;
    if (t.totalRow) {
      const f = wb.getFormula(sheet, t.ref.r2, c);
      const code = f ? Number(/^=SUBTOTAL\((\d+),[^,()]*\)$/i.exec(f)?.[1]) : NaN;
      if (f && TOTAL_FN[code]) a.push(`totalsRowFunction="${TOTAL_FN[code]}"`);
      else if (f) {
        a.push('totalsRowFunction="custom"');
        inner += `<totalsRowFormula>${esc(addXlfnPrefixes(f.replace(/^=/, "")))}</totalsRowFormula>`;
      } else {
        const v = wb.getValue(sheet, t.ref.r2, c);
        if (typeof v === "string" && v) a.push(`totalsRowLabel="${esc(v)}"`);
      }
    }
    return inner ? `<tableColumn ${a.join(" ")}>${inner}</tableColumn>` : `<tableColumn ${a.join(" ")}/>`;
  });
  return `<tableColumns count="${cols.length}">${cols.join("")}</tableColumns>`;
}

async function writeTableColumns(wb: WorkbookModel, sheet: SheetModel, sheetPath: string, parts: PackageParts): Promise<void> {
  if (!sheet.tables.length) return;
  const rels = await parts.zip.file(relsPathOf(sheetPath))?.async("string");
  if (!rels) return;
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    if (!/Type="[^"]*\/table"/.test(m[0])) continue;
    const target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (!target) continue;
    const path = resolveTarget(sheetPath, target);
    const f = parts.zip.file(path);
    if (!f) continue;
    const xml = await f.async("string");
    const display = /\bdisplayName="([^"]*)"/.exec(xml)?.[1] ?? "";
    const t = sheet.tables.find((x) => x.name.replace(/[^\p{L}\p{N}_]/gu, "_") === display);
    if (!t) continue;
    parts.zip.file(path, patchTableAutoFilter(xml.replace(/<tableColumns\b[\s\S]*?<\/tableColumns>/, tableColumnsXml(wb, sheet, t)), t.filter));
  }
}

export async function writeSheetExtras(wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (!part || !path) continue;
    part.set("dataValidations", dataValidationsXml(sheet));
    // Otomatik filtre ölçütleri (exceljs yalnız aralığı yazar)
    if (sheet.autoFilter && Object.keys(sheet.autoFilter.columns).length) part.set("autoFilter", autoFilterXml(sheet.autoFilter));
    await writeTableColumns(wb, sheet, path, parts);
  }
  await writeScenariosInto(wb, paths, parts);
}

export async function writeWorkbookExtras(wb: WorkbookModel, parts: PackageParts): Promise<void> {
  const part = await parts.workbook();
  if (part) writeWorkbookProtection(wb, part);
}
