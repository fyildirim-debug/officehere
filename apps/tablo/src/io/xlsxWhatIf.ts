/**
 * Durum Çözümlemesi ve koruma öğelerinin XLSX'e yazılması/okunması:
 * - Senaryolar: sayfadaki <scenarios><scenario><inputCells r val/></scenario></scenarios>
 * - Çalışma kitabı yapı koruması: workbook.xml <workbookProtection>
 * - Veri tabloları: sonuç aralığının sol üst hücresindeki <f t="dataTable" ...>
 */
import { addr, parseCellRef, parseRangeAddress } from "@officehere/formula";
import type { DataTableDef, Scenario, WorkbookProtection } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { esc, sheetPaths, unescapeXml } from "./ooxmlZip";
import type { PackageParts } from "./xlsxPost";
import type { XmlPart } from "./xmlPart";
import { loadZip } from "./zipCache";

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unescapeXml(m[1]) : undefined;
}

/* ------------------------------ Senaryolar ------------------------------ */

function valText(v: Scenario["cells"][number]["value"]): string {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "1" : "0";
  return String(v);
}

export function scenariosXml(list: Scenario[]): string | null {
  if (!list.length) return null;
  const items = list.map((s) => {
    const a = [`name="${esc(s.name)}"`];
    if (s.locked === false) a.push('locked="0"');
    if (s.hidden) a.push('hidden="1"');
    a.push(`count="${s.cells.length}"`);
    if (s.user) a.push(`user="${esc(s.user)}"`);
    if (s.comment) a.push(`comment="${esc(s.comment)}"`);
    const cells = s.cells.map((c) => `<inputCells r="${addr(c.row, c.col)}" val="${esc(valText(c.value))}"/>`).join("");
    return `<scenario ${a.join(" ")}>${cells}</scenario>`;
  });
  return `<scenarios current="0" show="0">${items.join("")}</scenarios>`;
}

export function parseScenarios(sheetXml: string): Scenario[] {
  const block = /<(?:\w+:)?scenarios\b[^>]*>([\s\S]*?)<\/(?:\w+:)?scenarios>/.exec(sheetXml)?.[1];
  if (!block) return [];
  const out: Scenario[] = [];
  for (const m of block.matchAll(/<(?:\w+:)?scenario\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?scenario>/g)) {
    const s: Scenario = { name: attr(m[1], "name") ?? `Senaryo ${out.length + 1}`, cells: [] };
    if (attr(m[1], "locked") === "0" || attr(m[1], "locked") === "false") s.locked = false;
    if (attr(m[1], "hidden") === "1" || attr(m[1], "hidden") === "true") s.hidden = true;
    const user = attr(m[1], "user");
    const comment = attr(m[1], "comment");
    if (user) s.user = user;
    if (comment) s.comment = comment;
    for (const c of m[2].matchAll(/<(?:\w+:)?inputCells\b([^>]*)\/?>/g)) {
      const pos = parseCellRef(attr(c[1], "r") ?? "");
      if (!pos) continue;
      const raw = attr(c[1], "val") ?? "";
      const n = Number(raw);
      s.cells.push({ row: pos.row, col: pos.col, value: raw === "" ? null : Number.isFinite(n) && /^-?[\d.]+(E[+-]?\d+)?$/i.test(raw) ? n : raw });
    }
    out.push(s);
  }
  return out;
}

export async function writeScenariosInto(wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  for (const sheet of wb.sheets) {
    if (!sheet.scenarios.length) continue;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    if (part) part.set("scenarios", scenariosXml(sheet.scenarios));
  }
}

/* ------------------------------ Kitap koruması ------------------------------ */

export function workbookProtectionXml(p: WorkbookProtection | null): string | null {
  if (!p || (!p.structure && !p.windows)) return null;
  const a: string[] = [];
  if (p.hash) a.push(`workbookAlgorithmName="${esc(p.hash.algorithm)}" workbookHashValue="${p.hash.value}" workbookSaltValue="${p.hash.salt}" workbookSpinCount="${p.hash.spin}"`);
  else if (p.legacy) a.push(`workbookPassword="${p.legacy}"`);
  if (p.structure) a.push('lockStructure="1"');
  if (p.windows) a.push('lockWindows="1"');
  return `<workbookProtection ${a.join(" ")}/>`;
}

export function writeWorkbookProtection(wb: WorkbookModel, part: XmlPart): void {
  part.set("workbookProtection", workbookProtectionXml(wb.protection));
}

export function parseWorkbookProtection(wbXml: string): WorkbookProtection | null {
  const tag = /<(?:\w+:)?workbookProtection\b[^>]*>/.exec(wbXml)?.[0];
  if (!tag) return null;
  const on = (n: string) => attr(tag, n) === "1" || attr(tag, n) === "true";
  const p: WorkbookProtection = { structure: on("lockStructure") };
  if (on("lockWindows")) p.windows = true;
  const value = attr(tag, "workbookHashValue");
  if (value) p.hash = { algorithm: attr(tag, "workbookAlgorithmName") ?? "SHA-512", value, salt: attr(tag, "workbookSaltValue") ?? "", spin: Number(attr(tag, "workbookSpinCount") ?? 100000) };
  const legacy = attr(tag, "workbookPassword");
  if (legacy) p.legacy = legacy;
  return p.structure || p.windows ? p : null;
}

/* ------------------------------ Veri tabloları ------------------------------ */

/** Sonuç aralığının sol üst hücresine eklenecek <f t="dataTable"> öğesi. */
export function dataTableFormulaXml(dt: DataTableDef): string {
  const a = [`t="dataTable"`, `ref="${addr(dt.ref.r1, dt.ref.c1)}:${addr(dt.ref.r2, dt.ref.c2)}"`];
  if (dt.rowInput && dt.colInput) a.push('dt2D="1"', 'dtr="1"', `r1="${addr(dt.rowInput.row, dt.rowInput.col)}"`, `r2="${addr(dt.colInput.row, dt.colInput.col)}"`);
  else if (dt.rowInput) a.push('dt2D="0"', 'dtr="1"', `r1="${addr(dt.rowInput.row, dt.rowInput.col)}"`);
  else if (dt.colInput) a.push('dt2D="0"', 'dtr="0"', `r1="${addr(dt.colInput.row, dt.colInput.col)}"`);
  return `<f ${a.join(" ")}/>`;
}

/** sheetData içinde veri tablosunun sol üst hücresine <f t="dataTable"> yerleştirir (değerler önbellek olarak kalır). */
export function injectDataTables(sheetData: string, list: DataTableDef[]): string {
  let xml = sheetData;
  for (const dt of list) {
    const ref = addr(dt.ref.r1, dt.ref.c1);
    const f = dataTableFormulaXml(dt);
    const re = new RegExp(`<c r="${ref}"([^>]*?)(/>|>([\\s\\S]*?)</c>)`);
    const m = re.exec(xml);
    if (m) {
      const inner = (m[3] ?? "").replace(/<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/g, "");
      xml = xml.slice(0, m.index) + `<c r="${ref}"${m[1]}>${f}${inner}</c>` + xml.slice(m.index + m[0].length);
    }
  }
  return xml;
}

export function parseDataTables(sheetXml: string, newId: () => string): DataTableDef[] {
  const out: DataTableDef[] = [];
  for (const m of sheetXml.matchAll(/<(?:\w+:)?f\b([^>]*\bt="dataTable"[^>]*)\/?>/g)) {
    const ref = parseRangeAddress(attr(m[1], "ref") ?? "");
    if (!ref) continue;
    const r1 = parseCellRef(attr(m[1], "r1") ?? "");
    const r2 = parseCellRef(attr(m[1], "r2") ?? "");
    const twoD = attr(m[1], "dt2D") === "1" || attr(m[1], "dt2D") === "true";
    const rowDir = attr(m[1], "dtr") === "1" || attr(m[1], "dtr") === "true";
    const dt: DataTableDef = { id: newId(), ref };
    if (twoD) {
      if (r1) dt.rowInput = { row: r1.row, col: r1.col };
      if (r2) dt.colInput = { row: r2.row, col: r2.col };
    } else if (r1) {
      if (rowDir) dt.rowInput = { row: r1.row, col: r1.col };
      else dt.colInput = { row: r1.row, col: r1.col };
    }
    if (dt.rowInput || dt.colInput) out.push(dt);
  }
  return out;
}

export async function writeDataTablesInto(wb: WorkbookModel, paths: Map<string, string>, parts: PackageParts): Promise<void> {
  for (const sheet of wb.sheets) {
    if (!sheet.dataTables.length) continue;
    const path = paths.get(sheet.name);
    const part = path ? await parts.part(path) : null;
    const sd = part?.get("sheetData");
    if (part && sd) part.set("sheetData", injectDataTables(sd, sheet.dataTables));
  }
}

/** Dosyadan senaryolar, veri tabloları ve kitap koruması. */
export async function readWhatIfFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await loadZip(data);
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!wbXml) return;
  wb.protection = parseWorkbookProtection(wbXml);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const p = paths.get(sheet.name);
    const xml = p ? await zip.file(p)?.async("string") : undefined;
    if (!xml) continue;
    if (xml.includes("scenario")) sheet.scenarios = parseScenarios(xml);
    if (xml.includes("dataTable")) sheet.dataTables = parseDataTables(xml, () => wb.newId("dt"));
  }
}
