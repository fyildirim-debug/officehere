/**
 * El ile sayfa sonları (<rowBreaks>/<colBreaks>): exceljs sütun sonlarını hiç
 * yazmaz, satır sonlarını da okumaz; ikisi de çalışma sayfası XML'ine
 * doğrudan yazılır ve oradan okunur. brk id'si yeni sayfanın ilk satırının /
 * sütununun sıfır tabanlı dizinidir (son, o satırın üstünde / sütunun solunda).
 */
import JSZip from "jszip";
import type { WorkbookModel } from "../model/workbook";
import { sheetPaths } from "./ooxmlZip";
import { WORKSHEET_ORDER, XmlPart } from "./xmlPart";
import type { PackageParts } from "./xlsxPost";

function breaksXml(tag: "rowBreaks" | "colBreaks", list: number[], max: number): string | null {
  if (!list.length) return null;
  return `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((i) => `<brk id="${i}" max="${max}" man="1"/>`).join("")}</${tag}>`;
}

/** Sayfa sonlarını parçaya yerleştirir (varsa eskisinin yerine; yeri şema sırası belirler). */
export function setBreaks(part: XmlPart, rows: number[], cols: number[]): void {
  part.set("rowBreaks", breaksXml("rowBreaks", rows, 16383));
  part.set("colBreaks", breaksXml("colBreaks", cols, 1048575));
}

/** Sayfa sonlarını çalışma sayfası XML'ine ekler (varsa eskisinin yerine). */
export function injectBreaks(xml: string, rows: number[], cols: number[]): string {
  const part = new XmlPart(xml, WORKSHEET_ORDER);
  setBreaks(part, rows, cols);
  return part.toString();
}

export async function writePageBreaksInto(wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  for (const sheet of wb.sheets) {
    const rows = sheet.page.rowBreaks ?? [];
    const cols = sheet.page.colBreaks ?? [];
    if (!rows.length && !cols.length) continue;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (part) setBreaks(part, rows, cols);
  }
}

/** Sayfa XML'indeki el ile sayfa sonları. */
export function parseBreaks(xml: string, tag: "rowBreaks" | "colBreaks"): number[] {
  const m = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${tag}>`).exec(xml);
  if (!m) return [];
  const ids = [...m[1].matchAll(/<(?:\w+:)?brk\b([^>]*)\/?>/g)]
    .filter((b) => !/\bman="(0|false)"/.test(b[1]))
    .map((b) => Number(/\bid="(\d+)"/.exec(b[1])?.[1]))
    .filter((n) => Number.isFinite(n) && n > 0);
  return [...new Set(ids)].sort((a, b) => a - b);
}

export async function readPageBreaksFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const f = path ? zip.file(path) : null;
    if (!f) continue;
    const xml = await f.async("string");
    if (!/Breaks\b/.test(xml)) continue;
    const rows = parseBreaks(xml, "rowBreaks");
    const cols = parseBreaks(xml, "colBreaks");
    if (rows.length) sheet.page.rowBreaks = rows;
    if (cols.length) sheet.page.colBreaks = cols;
  }
}
