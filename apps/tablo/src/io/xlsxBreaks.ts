/**
 * El ile sayfa sonları (<rowBreaks>/<colBreaks>): exceljs sütun sonlarını hiç
 * yazmaz, satır sonlarını da okumaz; ikisi de çalışma sayfası XML'ine
 * doğrudan yazılır ve oradan okunur. brk id'si yeni sayfanın ilk satırının /
 * sütununun sıfır tabanlı dizinidir (son, o satırın üstünde / sütunun solunda).
 */
import JSZip from "jszip";
import type { WorkbookModel } from "../model/workbook";
import { sheetPaths } from "./ooxmlZip";

/** CT_Worksheet sırasında sayfa sonlarından sonra gelen öğeler. */
const AFTER = ["<customProperties", "<cellWatches", "<ignoredErrors", "<smartTags", "<drawing ", "<drawing>", "<legacyDrawing", "<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"];

function breaksXml(tag: "rowBreaks" | "colBreaks", list: number[], max: number): string {
  if (!list.length) return "";
  return `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((i) => `<brk id="${i}" max="${max}" man="1"/>`).join("")}</${tag}>`;
}

/** Sayfa sonlarını çalışma sayfası XML'ine ekler (varsa eskisinin yerine). */
export function injectBreaks(xml: string, rows: number[], cols: number[]): string {
  let out = xml.replace(/<(rowBreaks|colBreaks)\b[^>]*\/>|<(rowBreaks|colBreaks)\b[\s\S]*?<\/\2>/g, "");
  const add = breaksXml("rowBreaks", rows, 16383) + breaksXml("colBreaks", cols, 1048575);
  if (!add) return out;
  let pos = -1;
  for (const t of AFTER) {
    const i = out.indexOf(t);
    if (i >= 0 && (pos < 0 || i < pos)) pos = i;
  }
  if (pos < 0) return out;
  out = out.slice(0, pos) + add + out.slice(pos);
  return out;
}

export async function writePageBreaksInto(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>): Promise<void> {
  for (const sheet of wb.sheets) {
    const rows = sheet.page.rowBreaks ?? [];
    const cols = sheet.page.colBreaks ?? [];
    if (!rows.length && !cols.length) continue;
    const path = paths.get(sheet.name);
    const f = path ? zip.file(path) : null;
    if (!path || !f) continue;
    zip.file(path, injectBreaks(await f.async("string"), rows, cols));
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
