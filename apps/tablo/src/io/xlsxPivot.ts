/**
 * PivotTable'ların XLSX'e yazılması ve geri okunması (ECMA-376 Bölüm 1,
 * §18.10 PivotTable). Rapor hücreleri zaten normal hücre değerleri olarak
 * yazılır; burada pivotCacheDefinition / pivotCacheRecords / pivotTable
 * parçaları eklenir. Önbellek kayıtları boş bırakılır ve refreshOnLoad
 * ile açılışta kaynaktan yeniden hesaplanır. Tam tanım (filtreler, düzen)
 * bilinmeyen uygulamaların yok saydığı bir extLst ögesinde de saklanır.
 */
import JSZip from "jszip";
import { formatRange, isError, parseRangeAddress, type Scalar } from "@officehere/formula";
import type { PivotDef, PivotFn } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { fieldItems, itemKey, newPivotDef, VALUES_FIELD, type PivotSource } from "../pivot/compute";
import { planPivot, readSource, setPivotMeta } from "../pivot/pivotOps";
import { NS_R, esc, nextRid, relsPathOf, resolveTarget, sheetPaths as sheetPathsOf, unescapeXml } from "./ooxmlZip";

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_CACHE_DEF = `${REL}/pivotCacheDefinition`;
const REL_CACHE_REC = `${REL}/pivotCacheRecords`;
const REL_PIVOT = `${REL}/pivotTable`;
const CT_CACHE_DEF = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const CT_CACHE_REC = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
const CT_PIVOT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
/** OfficeHere'nin tam PivotTable tanımı için uzantı kimliği. */
const EXT_URI = "{6C2A9E51-3B7D-4F0A-9C1E-0FF1CE4E0A01}";

const FN_XML: Record<PivotFn, string | null> = { sum: null, count: "count", average: "average", max: "max", min: "min", product: "product", countNums: "countNums" };

function sharedItemsXml(src: PivotSource, field: number, withItems: boolean): string {
  let hasStr = false;
  let hasNum = false;
  let hasBlank = false;
  let hasBool = false;
  let allInt = true;
  let min = Infinity;
  let max = -Infinity;
  for (const r of src.records) {
    const v = r[field] ?? null;
    if (v === null || v === "") hasBlank = true;
    else if (typeof v === "number") {
      hasNum = true;
      if (!Number.isInteger(v)) allInt = false;
      min = Math.min(min, v);
      max = Math.max(max, v);
    } else if (typeof v === "boolean") hasBool = true;
    else hasStr = true;
  }
  const attrs: string[] = [];
  if (!hasStr && !hasBool) {
    attrs.push('containsSemiMixedTypes="0"');
    attrs.push('containsString="0"');
  }
  if (hasNum) {
    attrs.push('containsNumber="1"');
    if (allInt) attrs.push('containsInteger="1"');
    attrs.push(`minValue="${min}"`, `maxValue="${max}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (hasStr && hasNum) attrs.push('containsMixedTypes="1"');
  if (!withItems) return `<sharedItems${attrs.length ? " " + attrs.join(" ") : ""}/>`;
  const items = fieldItems(src, field);
  attrs.push(`count="${items.length}"`);
  const xs = items
    .map(({ value: v }) => {
      if (v === null || v === "") return "<m/>";
      if (typeof v === "number") return `<n v="${v}"/>`;
      if (typeof v === "boolean") return `<b v="${v ? 1 : 0}"/>`;
      if (isError(v)) return `<e v="${esc(v.code)}"/>`;
      return `<s v="${esc(String(v))}"/>`;
    })
    .join("");
  return `<sharedItems ${attrs.join(" ")}>${xs}</sharedItems>`;
}

function cacheDefinitionXml(def: PivotDef, src: PivotSource, sheetName: string, recRid: string): string {
  const axis = new Set([...def.rows, ...def.cols, ...def.filters.map((f) => f.field)].filter((f) => f >= 0));
  const fields = src.headers
    .map((h, i) => `<cacheField name="${esc(h)}" numFmtId="0">${sharedItemsXml(src, i, axis.has(i))}</cacheField>`)
    .join("");
  const ref = formatRange({ r1: def.source.r1, c1: def.source.c1, r2: def.source.r1 + src.records.length, c2: def.source.c1 + src.headers.length - 1 });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" r:id="${recRid}" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="${ref}" sheet="${esc(sheetName)}"/></cacheSource><cacheFields count="${src.headers.length}">${fields}</cacheFields></pivotCacheDefinition>`;
}

const CACHE_RECORDS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_R}" count="0"/>`;

function pivotTableXml(def: PivotDef, src: PivotSource, cacheId: number, sourceSheetName: string): string {
  const out = def.output!;
  const tableTop = def.anchor.row;
  const rowF = def.rows.filter((f) => f >= 0);
  const colF = def.cols.filter((f) => f >= 0);
  const multi = def.values.length > 1;
  const compact = def.layout === "compact";
  const tabular = def.layout === "tabular";
  const headerRows = colF.length ? colF.length + 1 + (multi && !def.valuesOnRows ? 1 : 0) : 1;
  const labelCols = rowF.length === 0 ? (colF.length || (multi && def.valuesOnRows) ? 1 : 0) : compact ? 1 : rowF.length + (multi && def.valuesOnRows ? 1 : 0);
  const loc = formatRange({ r1: tableTop, c1: out.c1, r2: Math.max(tableTop, out.r2), c2: out.c2 });
  const pageCount = def.filters.length;
  const pivotFields = src.headers
    .map((_, i) => {
      const a: string[] = [];
      const isRow = rowF.includes(i);
      const isCol = colF.includes(i);
      const page = def.filters.find((f) => f.field === i);
      if (isRow) a.push('axis="axisRow"');
      else if (isCol) a.push('axis="axisCol"');
      else if (page) a.push('axis="axisPage"');
      if (def.values.some((v) => v.field === i)) a.push('dataField="1"');
      if (!compact) a.push('compact="0"');
      if (tabular) a.push('outline="0"');
      if (def.subtotals === "bottom" && (isRow || isCol)) a.push('subtotalTop="0"');
      if (def.subtotals === "none" && (isRow || isCol)) a.push('defaultSubtotal="0"');
      if (page && page.selected && page.selected.length > 1) a.push('multipleItemSelectionAllowed="1"');
      a.push('showAll="0"');
      if (!isRow && !isCol && !page) return `<pivotField ${a.join(" ")}/>`;
      const items = fieldItems(src, i);
      const hidden = new Set(def.hidden?.[i] ?? []);
      const pageSel = page?.selected ? new Set(page.selected) : null;
      const order = items.map((it, idx) => ({ it, idx }));
      if (def.sort?.[i]) {
        order.reverse();
        a.push('sortType="descending"');
      } else a.push('sortType="ascending"');
      const xs = order.map(({ it, idx }) => `<item x="${idx}"${hidden.has(it.key) || (pageSel && pageSel.size > 1 && !pageSel.has(it.key)) ? ' h="1"' : ""}/>`).join("");
      const def0 = def.subtotals === "none" && (isRow || isCol) ? "" : '<item t="default"/>';
      const n = items.length + (def0 ? 1 : 0);
      return `<pivotField ${a.join(" ")}><items count="${n}">${xs}${def0}</items></pivotField>`;
    })
    .join("");
  const sigma = '<field x="-2"/>';
  const rowFieldsXml = rowF.length || (multi && def.valuesOnRows) ? `<rowFields count="${rowF.length + (multi && def.valuesOnRows ? 1 : 0)}">${rowF.map((f) => `<field x="${f}"/>`).join("")}${multi && def.valuesOnRows ? sigma : ""}</rowFields>` : "";
  const colFieldsXml = colF.length || (multi && !def.valuesOnRows) ? `<colFields count="${colF.length + (multi && !def.valuesOnRows ? 1 : 0)}">${colF.map((f) => `<field x="${f}"/>`).join("")}${multi && !def.valuesOnRows ? sigma : ""}</colFields>` : "";
  const pageXml = pageCount
    ? `<pageFields count="${pageCount}">${def.filters
        .map((f) => {
          let item = "";
          if (f.selected && f.selected.length === 1) {
            const idx = fieldItems(src, f.field).findIndex((x) => x.key === f.selected![0]);
            if (idx >= 0) item = ` item="${idx}"`;
          }
          return `<pageField fld="${f.field}"${item} hier="-1"/>`;
        })
        .join("")}</pageFields>`
    : "";
  const dataXml = def.values.length
    ? `<dataFields count="${def.values.length}">${def.values
        .map((v) => {
          const fn = FN_XML[v.fn];
          return `<dataField name="${esc(v.name ?? `${captionOf(v.fn)} ${src.headers[v.field]}`)}" fld="${v.field}"${fn ? ` subtotal="${fn}"` : ""} baseField="0" baseItem="0"/>`;
        })
        .join("")}</dataFields>`
    : "";
  const ext = JSON.stringify({ ...def, output: undefined, buttons: undefined, sourceSheetName });
  const attrs = [
    `name="${esc(def.name)}"`,
    `cacheId="${cacheId}"`,
    'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"',
    'dataCaption="Değerler"',
    'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0"',
    multi && def.valuesOnRows ? 'dataOnRows="1"' : "",
    compact ? 'outline="1" outlineData="1"' : tabular ? 'compact="0" compactData="0" gridDropZones="1"' : 'compact="0" compactData="0" outline="1" outlineData="1"',
    'multipleFieldFilters="0"',
    `rowGrandTotals="${def.grandCols ? 1 : 0}"`,
    `colGrandTotals="${def.grandRows ? 1 : 0}"`,
    def.rowHeaders ? "" : 'showHeaders="0"',
    'rowHeaderCaption="Satır Etiketleri" colHeaderCaption="Sütun Etiketleri" grandTotalCaption="Genel Toplam"',
  ].filter(Boolean);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="${NS_MAIN}" ${attrs.join(" ")}><location ref="${loc}" firstHeaderRow="1" firstDataRow="${headerRows}" firstDataCol="${labelCols}"${pageCount ? ` rowPageCount="${pageCount}" colPageCount="1"` : ""}/><pivotFields count="${src.headers.length}">${pivotFields}</pivotFields>${rowFieldsXml}${colFieldsXml}${pageXml}${dataXml}<pivotTableStyleInfo name="${esc(def.style)}" showRowHeaders="${def.rowHeaders ? 1 : 0}" showColHeaders="${def.colHeaders ? 1 : 0}" showRowStripes="${def.bandedRows ? 1 : 0}" showColStripes="${def.bandedCols ? 1 : 0}" showLastColumn="1"/><extLst><ext uri="${EXT_URI}" xmlns:oh="urn:officehere:tablo"><oh:pivotDef>${esc(ext)}</oh:pivotDef></ext></extLst></pivotTableDefinition>`;
}

function captionOf(fn: PivotFn): string {
  return { sum: "Toplam", count: "Say", average: "Ortalama", max: "En Büyük", min: "En Küçük", product: "Çarpım", countNums: "Sayıları Say" }[fn];
}

/** PivotTable parçalarını pakete ekler (postProcess içinden). */
export async function writePivotsInto(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>): Promise<void> {
  const all = wb.sheets.flatMap((s) => s.pivots.filter((p) => p.output && p.rows.length + p.cols.length + p.values.length > 0).map((p) => ({ s, p })));
  if (!all.length) return;
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  let wbXml = await zip.file("xl/workbook.xml")!.async("string");
  let wbRels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const caches: string[] = [];
  let n = 0;
  for (const { s, p } of all) {
    const src = readSource(wb, p);
    const srcSheet = wb.sheetById(p.sourceSheet);
    const sheetPath = paths.get(s.name);
    if (typeof src === "string" || !srcSheet || !sheetPath) continue;
    n++;
    const cacheId = n;
    const defPath = `xl/pivotCache/pivotCacheDefinition${n}.xml`;
    const recPath = `xl/pivotCache/pivotCacheRecords${n}.xml`;
    const tblPath = `xl/pivotTables/pivotTable${n}.xml`;
    zip.file(defPath, cacheDefinitionXml(p, src, srcSheet.name, "rId1"));
    zip.file(relsPathOf(defPath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CACHE_REC}" Target="pivotCacheRecords${n}.xml"/></Relationships>`);
    zip.file(recPath, CACHE_RECORDS_XML);
    zip.file(tblPath, pivotTableXml(p, src, cacheId, srcSheet.name));
    zip.file(relsPathOf(tblPath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CACHE_DEF}" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`);
    ct = ct.replace("</Types>", `<Override PartName="/${defPath}" ContentType="${CT_CACHE_DEF}"/><Override PartName="/${recPath}" ContentType="${CT_CACHE_REC}"/><Override PartName="/${tblPath}" ContentType="${CT_PIVOT}"/></Types>`);
    const wrid = nextRid(wbRels);
    wbRels = wbRels.replace("</Relationships>", `<Relationship Id="${wrid}" Type="${REL_CACHE_DEF}" Target="pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`);
    caches.push(`<pivotCache cacheId="${cacheId}" r:id="${wrid}"/>`);
    // Sayfa → PivotTable ilişkisi
    const sheetRelsPath = relsPathOf(sheetPath);
    let sheetRels = zip.file(sheetRelsPath) ? await zip.file(sheetRelsPath)!.async("string") : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const srid = nextRid(sheetRels);
    sheetRels = sheetRels.replace("</Relationships>", `<Relationship Id="${srid}" Type="${REL_PIVOT}" Target="../pivotTables/pivotTable${n}.xml"/></Relationships>`);
    zip.file(sheetRelsPath, sheetRels);
  }
  if (!caches.length) return;
  const tag = `<pivotCaches>${caches.join("")}</pivotCaches>`;
  const after = ["<smartTagPr", "<smartTagTypes", "<webPublishing", "<fileRecoveryPr", "<webPublishObjects", "<extLst", "</workbook>"];
  const pos = after.map((b) => wbXml.indexOf(b)).filter((i) => i >= 0).sort((a, b) => a - b)[0];
  wbXml = wbXml.slice(0, pos) + tag + wbXml.slice(pos);
  if (!/xmlns:r=/.test(wbXml.slice(0, wbXml.indexOf(">", wbXml.indexOf("<workbook"))))) wbXml = wbXml.replace("<workbook", `<workbook xmlns:r="${NS_R}"`);
  zip.file("xl/workbook.xml", wbXml);
  zip.file("xl/_rels/workbook.xml.rels", wbRels);
  zip.file("[Content_Types].xml", ct);
}

/* ------------------------------ Okuma ------------------------------ */

function attr(xml: string, name: string): string | undefined {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(xml);
  return m ? unescapeXml(m[1]) : undefined;
}

/** Paketteki PivotTable tanımlarını canlı PivotTable olarak yükler. */
export async function readPivots(zip: JSZip, wb: WorkbookModel, paths: Map<string, string>): Promise<void> {
  for (const sheet of wb.sheets) {
    const sheetPath = paths.get(sheet.name);
    if (!sheetPath) continue;
    const relsFile = zip.file(relsPathOf(sheetPath));
    if (!relsFile) continue;
    const rels = await relsFile.async("string");
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      if (!m[0].includes(REL_PIVOT)) continue;
      const tPath = resolveTarget(sheetPath, attr(m[0], "Target") ?? "");
      const tFile = zip.file(tPath);
      if (!tFile) continue;
      try {
        const xml = await tFile.async("string");
        const def = (await parseOwn(xml, wb, sheet)) ?? (await parseStandard(zip, tPath, xml, wb, sheet));
        if (!def) continue;
        const src = readSource(wb, def);
        if (typeof src !== "string") setPivotMeta(def, planPivot(def, src));
        sheet.pivots.push(def);
      } catch {
        /* okunamayan PivotTable yok sayılır; hücre değerleri yerinde kalır */
      }
    }
  }
}

async function parseOwn(xml: string, wb: WorkbookModel, sheet: SheetModel): Promise<PivotDef | null> {
  const m = /<oh:pivotDef>([\s\S]*?)<\/oh:pivotDef>/.exec(xml);
  if (!m) return null;
  const raw = JSON.parse(unescapeXml(m[1])) as PivotDef & { sourceSheetName?: string };
  const src = raw.sourceSheetName ? wb.sheetByName(raw.sourceSheetName) : sheet;
  if (!src) return null;
  const loc = parseRangeAddress(attr(/<location\b[^>]*>/.exec(xml)?.[0] ?? "", "ref") ?? "");
  const def: PivotDef = { ...raw, id: wb.newId("pivot"), sourceSheet: src.id };
  delete (def as { sourceSheetName?: string }).sourceSheetName;
  def.output = outputFromLocation(loc, def);
  return def;
}

function outputFromLocation(loc: { r1: number; c1: number; r2: number; c2: number } | null, def: PivotDef) {
  if (!loc) return undefined;
  const top = Math.max(0, loc.r1 - (def.filters.length ? def.filters.length + 1 : 0));
  return { r1: top, c1: loc.c1, r2: loc.r2, c2: Math.max(loc.c2, loc.c1 + 1) };
}

/** Başka uygulamaların yazdığı PivotTable: alanlar, işlevler, düzen. */
async function parseStandard(zip: JSZip, tPath: string, xml: string, wb: WorkbookModel, sheet: SheetModel): Promise<PivotDef | null> {
  const tRels = zip.file(relsPathOf(tPath));
  if (!tRels) return null;
  const rel = [...(await tRels.async("string")).matchAll(/<Relationship\b[^>]*>/g)].find((x) => x[0].includes(REL_CACHE_DEF));
  if (!rel) return null;
  const cPath = resolveTarget(tPath, attr(rel[0], "Target") ?? "");
  const cXml = await zip.file(cPath)?.async("string");
  if (!cXml) return null;
  const ws = /<worksheetSource\b[^>]*>/.exec(cXml)?.[0];
  if (!ws) return null;
  let srcSheet: SheetModel | undefined = sheet;
  let srcRect = null;
  const tableName = attr(ws, "name");
  if (tableName) {
    const t = wb.sheets.flatMap((s) => s.tables.map((x) => ({ s, x }))).find((o) => o.x.name === tableName);
    if (t) {
      srcSheet = t.s;
      srcRect = t.x.ref;
    }
  } else {
    const sn = attr(ws, "sheet");
    if (sn) srcSheet = wb.sheetByName(sn);
    srcRect = parseRangeAddress(attr(ws, "ref") ?? "");
  }
  if (!srcSheet || !srcRect) return null;
  const loc = parseRangeAddress(attr(/<location\b[^>]*>/.exec(xml)?.[0] ?? "", "ref") ?? "");
  if (!loc) return null;
  const def = newPivotDef(wb.newId("pivot"), attr(xml, "name") ?? "PivotTable1", srcSheet.id, srcRect, { row: loc.r1, col: loc.c1 });
  const fieldsOf = (tag: string) => {
    const fm = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml);
    return fm ? [...fm[1].matchAll(/<field x="(-?\d+)"/g)].map((x) => Number(x[1])) : [];
  };
  def.rows = fieldsOf("rowFields").map((f) => (f === -2 ? VALUES_FIELD : f)).filter((f) => f !== VALUES_FIELD);
  def.cols = fieldsOf("colFields").map((f) => (f === -2 ? VALUES_FIELD : f)).filter((f) => f !== VALUES_FIELD);
  def.valuesOnRows = attr(/<pivotTableDefinition\b[^>]*>/.exec(xml)?.[0] ?? "", "dataOnRows") === "1";
  const fnOf: Record<string, PivotFn> = { count: "count", average: "average", max: "max", min: "min", product: "product", countNums: "countNums", sum: "sum" };
  def.values = [...xml.matchAll(/<dataField\b[^>]*>/g)].map((d) => ({ field: Number(attr(d[0], "fld") ?? 0), fn: fnOf[attr(d[0], "subtotal") ?? "sum"] ?? "sum", name: attr(d[0], "name") }));
  def.filters = [...xml.matchAll(/<pageField\b[^>]*>/g)].map((pf) => ({ field: Number(attr(pf[0], "fld") ?? 0), selected: null }));
  const head = /<pivotTableDefinition\b[^>]*>/.exec(xml)?.[0] ?? "";
  if (attr(head, "rowGrandTotals") === "0") def.grandCols = false;
  if (attr(head, "colGrandTotals") === "0") def.grandRows = false;
  if (attr(head, "compact") === "0") def.layout = attr(head, "outline") === "1" ? "outline" : "tabular";
  const style = /<pivotTableStyleInfo\b[^>]*>/.exec(xml)?.[0];
  if (style) {
    def.style = attr(style, "name") ?? def.style;
    def.bandedRows = attr(style, "showRowStripes") === "1";
    def.bandedCols = attr(style, "showColStripes") === "1";
  }
  // Gizli öğeler: önbellek öğelerine göre anahtar
  const cacheFields = [...cXml.matchAll(/<cacheField\b[\s\S]*?(?:<\/cacheField>|\/>)/g)].map((c) => [...c[0].matchAll(/<(s|n|b|m|e)(?: v="([^"]*)")?\s*\/>/g)].map((it) => itemKeyOf(it[1], it[2])));
  const pfs = [...xml.matchAll(/<pivotField\b[\s\S]*?(?:<\/pivotField>|\/>)/g)];
  pfs.forEach((pf, i) => {
    const hidden = [...pf[0].matchAll(/<item x="(\d+)"[^>]*h="1"/g)].map((h) => cacheFields[i]?.[Number(h[1])]).filter((k): k is string => !!k);
    if (hidden.length) def.hidden = { ...(def.hidden ?? {}), [i]: hidden };
  });
  if (/subtotalTop="0"/.test(xml)) def.subtotals = "bottom";
  def.output = outputFromLocation(loc, def);
  return def;
}

function itemKeyOf(kind: string, v: string | undefined): string {
  const val: Scalar = kind === "n" ? Number(v) : kind === "b" ? v === "1" : kind === "m" ? null : unescapeXml(v ?? "");
  return itemKey(val);
}

/** Ham dosya verisinden PivotTable'ları okur (readXlsx sonrasında). */
export async function readPivotsFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  if (!Object.keys(zip.files).some((f) => f.startsWith("xl/pivotTables/"))) return;
  await readPivots(zip, wb, await sheetPathsOf(zip));
}
