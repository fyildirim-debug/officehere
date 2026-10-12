/**
 * XLSX okuma/yazma (exceljs, MIT). Değerler, formüller (önbellek sonuçlarıyla),
 * stiller, sayı biçimleri, birleştirmeler, boyutlar, gizli satır/sütunlar,
 * bölmeleri dondurma, otomatik filtre, notlar, köprüler, koşullu biçimlendirme,
 * veri doğrulama, adlandırılmış aralıklar, tablolar, resimler ve grafikler.
 */
import { addr, cellKey, colToName, tokenize, EN_SYNTAX, errorFromCode, formatRange, isError, parseRangeAddress, resolveFunctionName, stripFunctionPrefix, FUNCTIONS, type Scalar } from "@officehere/formula";
import { MAX_COLS, MAX_ROWS } from "../model/rect";
import { widthCharsToPx, widthPxToChars, ptToPx, pxToPt } from "../model/axis";
import type { CondRule, PageSetup, Rect, ValidationRule } from "../model/types";
import { WorkbookModel, type SheetModel } from "../model/workbook";
import { ColorResolver, fromXStyle, toXStyle } from "./xlsxStyles";
import { readCharts, postProcess } from "./xlsxCharts";

/* eslint-disable @typescript-eslint/no-explicit-any */

async function loadExcelJS(): Promise<any> {
  const mod: any = await import("exceljs");
  return mod.default ?? mod;
}

const PAPER: Record<number, PageSetup["paper"]> = { 1: "Letter", 5: "Legal", 8: "A3", 9: "A4", 11: "A5" };
const PAPER_ID: Record<string, number> = { Letter: 1, Legal: 5, A3: 8, A4: 9, A5: 11 };

function utcSerial(d: Date): number {
  return 25569 + d.getTime() / 86400000;
}

/** Dosyaya yazarken yeni fonksiyonlara "_xlfn." öneki eklenir (metin sabitleri atlanır). */
export function addXlfnPrefixes(formula: string): string {
  return tokenize(formula, EN_SYNTAX)
    .map((t) => {
      if (t.type !== "func") return t.text;
      const def = FUNCTIONS.get(String(t.value).toUpperCase());
      return def?.xlfn ? `_xlfn.${t.text}` : t.text;
    })
    .join("");
}

function rectsFromSqref(s: string): Rect[] {
  return s
    .split(/\s+/)
    .map((p) => parseRangeAddress(p))
    .filter((r): r is Rect => !!r);
}

/* ============================== OKUMA ============================== */

export async function readXlsx(data: ArrayBuffer): Promise<WorkbookModel> {
  const ExcelJS = await loadExcelJS();
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(data);
  const wb = new WorkbookModel(true);
  const cr = new ColorResolver();
  cr.loadTheme(x._themes?.theme1);
  wb.props = { title: x.title || undefined, subject: x.subject || undefined, author: x.creator || undefined, company: x.company || undefined, created: x.created, modified: x.modified };

  const worksheets: any[] = [...x.worksheets].sort((a: any, b: any) => (a.orderNo ?? 0) - (b.orderNo ?? 0));
  for (const ws of worksheets) {
    const sheet = wb.addSheet(ws.name);
    readSheet(wb, sheet, ws, cr, x);
  }
  // Adlandırılmış aralıklar
  for (const dn of x.definedNames?.model ?? []) {
    if (!dn.name || dn.name.startsWith("_xlnm")) continue;
    if (!dn.ranges?.length) continue;
    try {
      wb.engine.defineName(dn.name, "=" + dn.ranges[0]);
    } catch {
      /* yok say */
    }
  }
  if (!wb.sheets.length) wb.addSheet("Sayfa1");
  const firstVisible = wb.sheets.findIndex((s) => !s.hidden);
  wb.activeIndex = Math.max(0, x.views?.[0]?.activeTab ?? firstVisible);
  if (wb.sheets[wb.activeIndex]?.hidden) wb.activeIndex = Math.max(0, firstVisible);
  // Grafikler (exceljs okumaz; DrawingML doğrudan ayrıştırılır)
  try {
    await readCharts(data, wb);
  } catch {
    /* grafikler okunamazsa yok say */
  }
  wb.engine.recalc(true);
  return wb;
}

function readSheet(wb: WorkbookModel, sheet: SheetModel, ws: any, cr: ColorResolver, x: any): void {
  if (ws.state && ws.state !== "visible") sheet.hidden = true;
  const tab = cr.resolve(ws.properties?.tabColor);
  if (tab) sheet.tabColor = tab;
  if (ws.properties?.defaultColWidth) sheet.cols.setDefault(widthCharsToPx(ws.properties.defaultColWidth));
  if (ws.properties?.defaultRowHeight && Math.abs(ws.properties.defaultRowHeight - 15) > 0.1) sheet.rows.setDefault(ptToPx(ws.properties.defaultRowHeight));
  // Sütunlar
  const colCount = Math.min(Math.max(ws.columnCount ?? 0, ws._columns?.length ?? 0, ws.columns?.length ?? 0), MAX_COLS);
  for (let c = 1; c <= colCount; c++) {
    const col = ws.getColumn(c);
    if (col.width !== undefined && col.width !== null && (col as any)._worksheet) {
      const w = widthCharsToPx(col.width);
      if (Math.abs(w - sheet.cols.defaultSize) > 0.5) sheet.cols.set(c - 1, w);
    }
    if (col.hidden) sheet.cols.setHidden(c - 1, true);
    const st = col.style && Object.keys(col.style).length ? fromXStyle(col.style, cr) : null;
    if (st && Object.keys(st).length) sheet.colStyles.set(c - 1, wb.styles.intern(st));
  }
  const notesCells: { row: number; col: number; text: string; author?: string }[] = [];
  const validations = new Map<string, { rule: ValidationRule; cells: [number, number][] }>();
  ws.eachRow({ includeEmpty: true }, (row: any, rn: number) => {
    const r = rn - 1;
    if (r >= MAX_ROWS) return;
    if (row.height !== undefined && row.height !== null && (row.customHeight || Math.abs(row.height - 15) > 0.1)) sheet.rows.set(r, ptToPx(row.height));
    if (row.hidden) sheet.rows.setHidden(r, true);
    row.eachCell({ includeEmpty: true }, (cell: any, cn: number) => {
      const c = cn - 1;
      readCell(wb, sheet, cell, r, c, cr);
      if (cell.note) {
        const n = cell.note;
        const text = typeof n === "string" ? n : (n.texts ?? []).map((t: any) => t.text).join("");
        if (text) notesCells.push({ row: r, col: c, text: text.replace(/^[^:\n]{1,40}:\n/, "") });
      }
    });
  });
  for (const n of notesCells) sheet.notes.set(cellKey(n.row, n.col), { text: n.text });
  // Veri doğrulama (hücre değeri olmasa da)
  for (const [key, dv] of Object.entries<any>(ws.dataValidations?.model ?? {})) {
    if (!dv?.type) continue;
    const rr = parseRangeAddress(key.replace(/^range:/, "").replace(/\$/g, ""));
    if (!rr) continue;
    const rule = readValidation(dv);
    const k = JSON.stringify(rule);
    let g = validations.get(k);
    if (!g) validations.set(k, (g = { rule, cells: [] }));
    for (let i = rr.r1; i <= Math.min(rr.r2, rr.r1 + 5000); i++) for (let j = rr.c1; j <= Math.min(rr.c2, rr.c1 + 200); j++) g.cells.push([i, j]);
  }
  for (const g of validations.values()) sheet.validations.push({ id: wb.newId("dv"), rule: g.rule, ranges: compressCells(g.cells) });
  // Birleştirmeler
  const merges: string[] = ws.model?.merges ?? [];
  for (const m of merges) {
    const r = parseRangeAddress(m);
    if (r) sheet.merges.push(r);
  }
  // Görünüm
  const v = ws.views?.[0];
  if (v) {
    if (v.state === "frozen") sheet.frozen = { rows: v.ySplit ?? 0, cols: v.xSplit ?? 0 };
    if (v.showGridLines === false) sheet.showGridlines = false;
    if (v.showRowColHeaders === false) sheet.showHeadings = false;
    if (v.zoomScale) sheet.zoom = v.zoomScale;
    if (v.activeCell) {
      const a = parseRangeAddress(v.activeCell);
      if (a) sheet.selection = { ranges: [{ r1: a.r1, c1: a.c1, r2: a.r1, c2: a.c1 }], active: { row: a.r1, col: a.c1 }, anchor: { row: a.r1, col: a.c1 } };
    }
    if (v.view === "pageLayout") sheet.viewMode = "pageLayout";
    if (v.view === "pageBreakPreview") sheet.viewMode = "pageBreak";
  }
  // Otomatik filtre
  if (ws.autoFilter) {
    const af = typeof ws.autoFilter === "string" ? ws.autoFilter : `${ws.autoFilter.from?.toString?.() ?? ""}:${ws.autoFilter.to?.toString?.() ?? ""}`;
    const r = parseRangeAddress(af.replace(/\$/g, ""));
    if (r) sheet.autoFilter = { range: r, columns: {} };
  }
  // Koşullu biçimlendirme
  for (const cf of ws.conditionalFormattings ?? []) {
    const ranges = rectsFromSqref(cf.ref ?? "");
    for (const rule of cf.rules ?? []) {
      const cr2 = readCondRule(rule, cr);
      if (cr2) sheet.condFormats.push({ ...cr2, id: wb.newId("cf"), ranges });
    }
  }
  // Sayfa yapısı
  const ps = ws.pageSetup ?? {};
  if (ps.orientation === "landscape") sheet.page.orientation = "landscape";
  if (ps.paperSize && PAPER[ps.paperSize]) sheet.page.paper = PAPER[ps.paperSize];
  if (ps.margins) sheet.page.margins = { ...sheet.page.margins, ...ps.margins };
  if (ps.printArea) {
    const r = parseRangeAddress(String(ps.printArea).split(",")[0].replace(/\$/g, "").replace(/^.*!/, ""));
    if (r) sheet.page.printArea = r;
  }
  if (ps.showGridLines) sheet.page.printGridlines = true;
  if (ps.fitToPage) sheet.page.fitToWidth = true;
  const hf = ws.headerFooter ?? {};
  if (hf.oddHeader) sheet.page.header = stripHF(hf.oddHeader);
  if (hf.oddFooter) sheet.page.footer = stripHF(hf.oddFooter);
  if (ws.sheetProtection?.sheet) sheet.protected = true;
  // Tablolar
  try {
    for (const entry of ws.getTables?.() ?? []) {
      const t = Array.isArray(entry) ? entry[0] : entry;
      const m = t?.table ?? t?.model ?? t;
      if (!m?.ref && !m?.tableRef) continue;
      const ref = parseRangeAddress(m.tableRef ?? m.ref);
      if (!ref) continue;
      let rr = ref;
      if (rr.r1 === rr.r2 && m.columns?.length) {
        const rows = (m.rows?.length ?? 0) + (m.headerRow !== false ? 1 : 0) + (m.totalsRow ? 1 : 0);
        rr = { r1: ref.r1, c1: ref.c1, r2: ref.r1 + Math.max(0, rows - 1), c2: ref.c1 + m.columns.length - 1 };
      }
      const st = m.style ?? {};
      sheet.tables.push({
        id: wb.newId("tbl"),
        name: m.name ?? m.displayName ?? `Tablo${sheet.tables.length + 1}`,
        ref: rr,
        style: st.theme ?? "TableStyleMedium2",
        headerRow: m.headerRow !== false,
        totalRow: !!m.totalsRow,
        bandedRows: st.showRowStripes !== false,
        bandedCols: !!st.showColumnStripes,
        firstColumn: !!st.showFirstColumn,
        lastColumn: !!st.showLastColumn,
        filterButton: (m.columns ?? []).some((c: any) => c.filterButton !== false),
        columns: (m.columns ?? []).map((c: any) => c.name),
        filter: { range: m.totalsRow ? { ...rr, r2: rr.r2 - 1 } : rr, columns: {} },
      });
    }
  } catch {
    /* tablo okunamadı */
  }
  // Resimler
  try {
    for (const img of ws.getImages?.() ?? []) {
      const media = x.getImage(Number(img.imageId));
      if (!media) continue;
      const ext = media.extension === "jpg" ? "jpeg" : media.extension ?? "png";
      let b64 = media.base64 as string | undefined;
      if (!b64 && media.buffer) b64 = bytesToBase64(new Uint8Array(media.buffer));
      if (!b64) continue;
      const src = b64.startsWith("data:") ? b64 : `data:image/${ext};base64,${b64.replace(/^.*base64,/, "")}`;
      const tl = img.range.tl;
      const col = Math.floor(tl.nativeCol ?? tl.col ?? 0);
      const row = Math.floor(tl.nativeRow ?? tl.row ?? 0);
      const dx = (tl.nativeColOff ?? 0) / 9525;
      const dy = (tl.nativeRowOff ?? 0) / 9525;
      let width = 200;
      let height = 150;
      if (img.range.ext) {
        width = img.range.ext.width;
        height = img.range.ext.height;
      } else if (img.range.br) {
        const br = img.range.br;
        const x2 = sheet.cols.offset(Math.floor(br.nativeCol ?? br.col)) + (br.nativeColOff ?? 0) / 9525;
        const y2 = sheet.rows.offset(Math.floor(br.nativeRow ?? br.row)) + (br.nativeRowOff ?? 0) / 9525;
        width = Math.max(10, x2 - (sheet.cols.offset(col) + dx));
        height = Math.max(10, y2 - (sheet.rows.offset(row) + dy));
      }
      sheet.images.push({ id: wb.newId("img"), name: `Resim ${sheet.images.length + 1}`, src, anchor: { row, col, dx, dy }, width, height });
    }
  } catch {
    /* resim okunamadı */
  }
}

/** Üst/alt bilgi kodları: dosya (&P, &N, &A, &F, &D) ↔ arayüz (&[Sayfa] …). */
const HF_CODES: [string, string][] = [
  ["&P", "&[Sayfa]"],
  ["&N", "&[Sayfalar]"],
  ["&A", "&[Sekme]"],
  ["&F", "&[Dosya]"],
  ["&D", "&[Tarih]"],
];

export function stripHF(s: string): string {
  let t = s.replace(/&[LCR]/g, "").replace(/&"[^"]*"/g, "").replace(/&\d+/g, "");
  for (const [file, ui] of HF_CODES) t = t.split(file).join(ui);
  return t.replace(/&(?!\[)[A-Z]/g, "").replace(/&&/g, "&").trim();
}

export function hfToFile(s: string): string {
  let t = s.replace(/&(?!\[)/g, "&&");
  for (const [file, ui] of HF_CODES) t = t.split(ui).join(file);
  return "&C" + t;
}

function bytesToBase64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return typeof btoa === "function" ? btoa(s) : Buffer.from(bytes).toString("base64");
}

function readCell(wb: WorkbookModel, sheet: SheetModel, cell: any, r: number, c: number, cr: ColorResolver): void {
  const v = cell.value;
  const styleObj = cell.style;
  let input: Scalar | { formula: string } = null;
  let link: string | undefined;
  if (cell.type === 1) {
    // Birleştirilmiş alt hücre: yalnız stil
  } else if (v !== null && v !== undefined) {
    if (typeof v === "number" || typeof v === "boolean") input = v;
    else if (typeof v === "string") input = v;
    else if (v instanceof Date) input = utcSerial(v);
    else if (typeof v === "object") {
      if ("formula" in v || "sharedFormula" in v) {
        const f = cell.formula as string | undefined;
        if (f) input = { formula: "=" + normalizeFileFormula(f) };
        else if (v.result !== undefined) input = toScalar(v.result);
      } else if ("richText" in v) input = (v.richText as any[]).map((t) => t.text).join("");
      else if ("hyperlink" in v) {
        input = typeof v.text === "string" ? v.text : (v.text?.richText ?? []).map((t: any) => t.text).join("") || v.hyperlink;
        link = v.hyperlink;
      } else if ("error" in v) input = errorFromCode(String(v.error)) ?? null;
    }
  }
  if (input !== null) wb.setInputRaw(sheet, r, c, input);
  if (link) sheet.links.set(cellKey(r, c), link.startsWith("#") ? link : link);
  if (styleObj && Object.keys(styleObj).length) {
    const st = fromXStyle(styleObj, cr);
    if (Object.keys(st).length) {
      const id = wb.styles.intern(st);
      const inherited = sheet.colStyles.get(c);
      if (id !== inherited) sheet.styles.set(cellKey(r, c), id);
    }
  }
}

function toScalar(v: any): Scalar {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") return v;
  if (v instanceof Date) return utcSerial(v);
  if (typeof v === "object" && "error" in v) return errorFromCode(String(v.error)) ?? null;
  return String(v);
}

/** Dosyadaki formül: önekleri ve bilinmeyen biçimleri temizler. */
export function normalizeFileFormula(f: string): string {
  return f.replace(/_xlfn\.|_xlws\.|_xludf\./g, "").replace(/_xlpm\./g, "");
}

function readValidation(dv: any): ValidationRule {
  const strip = (f: any) => (f === undefined || f === null ? undefined : f instanceof Date ? String(utcSerial(f)) : String(f));
  let f1 = strip(dv.formulae?.[0]);
  const f2 = strip(dv.formulae?.[1]);
  if (dv.type === "list" && f1) {
    if (f1.startsWith('"')) f1 = f1.slice(1, -1).split(",").join(";");
    else f1 = "=" + f1.replace(/^=/, "");
  } else if (f1 && !/^-?[\d.]+$/.test(f1) && dv.type !== "date") f1 = "=" + f1.replace(/^=/, "");
  return {
    type: dv.type,
    operator: dv.operator,
    formula1: f1,
    formula2: f2,
    allowBlank: dv.allowBlank,
    showDropdown: dv.showDropDown !== true,
    errorTitle: dv.errorTitle,
    error: dv.error,
    errorStyle: dv.errorStyle,
    promptTitle: dv.promptTitle,
    prompt: dv.prompt,
  };
}

/** Tek hücre listesini dikdörtgenlere sıkıştırır (sütun bazında ardışık satırlar). */
export function compressCells(cells: [number, number][]): Rect[] {
  const byCol = new Map<number, number[]>();
  for (const [r, c] of cells) {
    const l = byCol.get(c) ?? [];
    l.push(r);
    byCol.set(c, l);
  }
  const runs: Rect[] = [];
  for (const [c, rows] of byCol) {
    rows.sort((a, b) => a - b);
    let start = rows[0];
    let prev = rows[0];
    for (let i = 1; i <= rows.length; i++) {
      if (i < rows.length && rows[i] === prev + 1) {
        prev = rows[i];
        continue;
      }
      runs.push({ r1: start, c1: c, r2: prev, c2: c });
      start = rows[i];
      prev = rows[i];
    }
  }
  // Aynı satır aralığındaki komşu sütunları birleştir
  runs.sort((a, b) => a.r1 - b.r1 || a.r2 - b.r2 || a.c1 - b.c1);
  const out: Rect[] = [];
  for (const r of runs) {
    const last = out[out.length - 1];
    if (last && last.r1 === r.r1 && last.r2 === r.r2 && last.c2 + 1 === r.c1) last.c2 = r.c2;
    else out.push({ ...r });
  }
  return out;
}

function readCondRule(rule: any, cr: ColorResolver): Omit<CondRule, "id" | "ranges"> | null {
  const style = rule.style ? fromXStyle(rule.style, cr) : undefined;
  const formulas = (rule.formulae ?? []).map((f: any) => "=" + normalizeFileFormula(String(f)));
  switch (rule.type) {
    case "cellIs":
      return { type: "cellIs", operator: rule.operator, values: formulas, style };
    case "containsText": {
      const op = rule.operator;
      if (op === "containsBlanks") return { type: "containsBlanks", style };
      if (op === "notContainsBlanks") return { type: "notContainsBlanks", style };
      if (op === "containsErrors") return { type: "containsErrors", style };
      return { type: op === "notContains" ? "notContainsText" : op === "beginsWith" ? "beginsWith" : op === "endsWith" ? "endsWith" : "containsText", text: rule.text, style };
    }
    case "top10":
      return { type: "top10", rank: rule.rank, percent: !!rule.percent, bottom: !!rule.bottom, style };
    case "aboveAverage":
      return { type: "aboveAverage", below: rule.aboveAverage === false, style };
    case "duplicateValues":
    case "uniqueValues":
      return { type: rule.type, style };
    case "expression":
      return { type: "expression", values: formulas, style };
    case "timePeriod":
      return { type: "timePeriod", timePeriod: rule.timePeriod, style };
    case "dataBar":
      return { type: "dataBar", color: cr.resolve(rule.color) ?? "#638EC6" };
    case "colorScale":
      return { type: "colorScale", colors: (rule.color ?? []).map((c: any) => cr.resolve(c) ?? "#FFFFFF") };
  }
  return null;
}

/* ============================== YAZMA ============================== */

export async function writeXlsx(wb: WorkbookModel): Promise<Uint8Array<ArrayBuffer>> {
  const ExcelJS = await loadExcelJS();
  const x = new ExcelJS.Workbook();
  x.creator = wb.props.author ?? "OfficeHere Tablo";
  x.lastModifiedBy = wb.props.author ?? "OfficeHere Tablo";
  x.created = wb.props.created ?? new Date();
  x.modified = new Date();
  if (wb.props.title) x.title = wb.props.title;
  if (wb.props.subject) x.subject = wb.props.subject;
  if (wb.props.company) x.company = wb.props.company;
  x.calcProperties = { fullCalcOnLoad: true };
  x.views = [{ x: 0, y: 0, width: 28800, height: 15000, firstSheet: 0, activeTab: wb.activeIndex, visibility: "visible" }];
  wb.engine.recalc();
  for (const sheet of wb.sheets) writeSheet(wb, sheet, x);
  // Adlandırılmış aralıklar
  for (const n of wb.engine.getNames()) {
    const body = n.formula.slice(1);
    if (/^('[^']+'|[\p{L}_][\p{L}\p{N}_.]*)!\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/u.test(body)) {
      try {
        x.definedNames.add(body, n.name);
      } catch {
        /* yok say */
      }
    }
  }
  const buf = await x.xlsx.writeBuffer();
  let bytes: Uint8Array = new Uint8Array(buf as ArrayBuffer);
  bytes = await postProcess(bytes, wb);
  return new Uint8Array(bytes);
}

function writeSheet(wb: WorkbookModel, sheet: SheetModel, x: any): void {
  const views: any[] = [];
  const base: any = { showGridLines: sheet.showGridlines, showRowColHeaders: sheet.showHeadings, zoomScale: sheet.zoom, activeCell: addr(sheet.selection.active.row, sheet.selection.active.col) };
  if (sheet.viewMode === "pageLayout") base.view = "pageLayout";
  if (sheet.viewMode === "pageBreak") base.view = "pageBreakPreview";
  if (sheet.frozen.rows || sheet.frozen.cols) views.push({ ...base, state: "frozen", xSplit: sheet.frozen.cols, ySplit: sheet.frozen.rows, topLeftCell: addr(sheet.frozen.rows, sheet.frozen.cols) });
  else views.push({ ...base, state: "normal" });
  const ws = x.addWorksheet(sheet.name, {
    properties: { tabColor: sheet.tabColor ? { argb: "FF" + sheet.tabColor.slice(1) } : undefined, defaultRowHeight: pxToPt(sheet.rows.defaultSize) },
    views,
    state: sheet.hidden ? "hidden" : "visible",
  });
  if (Math.abs(sheet.cols.defaultSize - 64) > 0.5) ws.properties.defaultColWidth = widthPxToChars(sheet.cols.defaultSize);
  // Sütunlar
  const colSet = new Set<number>([...sheet.cols.customItems().map(([c]) => c), ...sheet.cols.hiddenItems(), ...sheet.colStyles.keys()]);
  for (const c of [...colSet].sort((a, b) => a - b)) {
    if (c >= MAX_COLS) continue;
    const col = ws.getColumn(c + 1);
    if (sheet.cols.isCustom(c)) col.width = widthPxToChars(sheet.cols.rawSize(c));
    if (sheet.cols.isHidden(c)) col.hidden = true;
    const cs = sheet.colStyles.get(c);
    if (cs) col.style = toXStyle(wb.styles.get(cs));
  }
  // Satırlar
  const rowSet = new Set<number>([...sheet.rows.customItems().map(([r]) => r), ...sheet.rows.hiddenItems(), ...sheet.rowStyles.keys()]);
  for (const r of rowSet) {
    const row = ws.getRow(r + 1);
    if (sheet.rows.isCustom(r)) row.height = pxToPt(sheet.rows.rawSize(r));
    if (sheet.rows.isHidden(r)) row.hidden = true;
    const rs = sheet.rowStyles.get(r);
    if (rs) row.style = toXStyle(wb.styles.get(rs));
  }
  // Tablolar (hücrelerden önce: addTable hücrelere yazar)
  for (const t of sheet.tables) {
    try {
      const rows: any[][] = [];
      const start = t.headerRow ? t.ref.r1 + 1 : t.ref.r1;
      const end = t.totalRow ? t.ref.r2 - 1 : t.ref.r2;
      for (let r = start; r <= end; r++) {
        const row: any[] = [];
        for (let c = t.ref.c1; c <= t.ref.c2; c++) row.push(wb.getValue(sheet, r, c));
        rows.push(row);
      }
      if (!rows.length) rows.push(t.columns.map(() => null));
      ws.addTable({
        name: t.name.replace(/[^\p{L}\p{N}_]/gu, "_"),
        displayName: t.name.replace(/[^\p{L}\p{N}_]/gu, "_"),
        ref: addr(t.ref.r1, t.ref.c1),
        headerRow: t.headerRow,
        totalsRow: t.totalRow,
        style: { theme: t.style, showRowStripes: t.bandedRows, showColumnStripes: t.bandedCols, showFirstColumn: t.firstColumn, showLastColumn: t.lastColumn },
        columns: t.columns.map((name, i) => ({ name, filterButton: t.filterButton, ...(t.totalRow ? (i === 0 ? { totalsRowLabel: "Toplam" } : i === t.columns.length - 1 ? { totalsRowFunction: "sum" } : {}) : {}) })),
        rows,
      });
    } catch {
      /* tablo yazılamadı: hücreler yine yazılır */
    }
  }
  // Hücreler
  const done = new Set<number>();
  wb.engine.forEachCell(sheet.id, (r, c, rec) => {
    const cell = ws.getCell(r + 1, c + 1);
    const k = cellKey(r, c);
    done.add(k);
    if (rec.formula) {
      const result = cellResult(rec.value);
      const f = addXlfnPrefixes(rec.formula.text);
      const spill = wb.engine.getSpill(sheet.id, r, c);
      if (spill && spill.anchorRow === r && spill.anchorCol === c && (spill.rows > 1 || spill.cols > 1)) {
        cell.value = { formula: f, result, shareType: "array", ref: formatRange({ r1: r, c1: c, r2: r + spill.rows - 1, c2: c + spill.cols - 1 }) };
        for (let i = 0; i < spill.rows; i++)
          for (let j = 0; j < spill.cols; j++) {
            if (!i && !j) continue;
            const v = wb.getValue(sheet, r + i, c + j);
            if (v !== null) ws.getCell(r + i + 1, c + j + 1).value = cellResult(v);
          }
      } else cell.value = result === undefined ? { formula: f } : { formula: f, result };
    } else {
      const v = rec.value;
      const link = sheet.links.get(k);
      if (link && typeof v === "string") cell.value = { text: v, hyperlink: link };
      else cell.value = cellResult(v) ?? null;
    }
  });
  // Stiller
  for (const [k, id] of sheet.styles) {
    const r = Math.floor(k / 16384);
    const c = k % 16384;
    const st = wb.styles.get(id);
    ws.getCell(r + 1, c + 1).style = toXStyle(st);
  }
  // Köprüler (değersiz hücreler)
  for (const [k, link] of sheet.links) {
    if (done.has(k)) continue;
    const r = Math.floor(k / 16384);
    const c = k % 16384;
    ws.getCell(r + 1, c + 1).value = { text: link, hyperlink: link };
  }
  // Notlar
  for (const [k, n] of sheet.notes) {
    const r = Math.floor(k / 16384);
    const c = k % 16384;
    ws.getCell(r + 1, c + 1).note = n.text;
  }
  // Birleştirmeler
  for (const m of sheet.merges) {
    try {
      ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    } catch {
      /* çakışan birleştirme */
    }
  }
  if (sheet.autoFilter) ws.autoFilter = formatRange(sheet.autoFilter.range);
  // Veri doğrulama
  for (const v of sheet.validations) {
    const dv = writeValidation(v.rule);
    let count = 0;
    for (const r of v.ranges) {
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 2000); row++)
        for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 100); c++) {
          if (++count > 20000) break;
          ws.getCell(row + 1, c + 1).dataValidation = dv;
        }
    }
  }
  // Koşullu biçimlendirme
  let priority = 1;
  for (const rule of sheet.condFormats) {
    const xr = writeCondRule(rule, priority++);
    if (!xr) continue;
    ws.addConditionalFormatting({ ref: rule.ranges.map((r) => formatRange(r)).join(" "), rules: [xr] });
  }
  // Sayfa yapısı
  const p = sheet.page;
  ws.pageSetup = {
    paperSize: PAPER_ID[p.paper] ?? 9,
    orientation: p.orientation,
    margins: p.margins,
    showGridLines: p.printGridlines,
    ...(p.printArea ? { printArea: formatRange(p.printArea) } : {}),
    ...(p.fitToWidth ? { fitToPage: true, fitToWidth: 1, fitToHeight: 0 } : {}),
  };
  if (p.header || p.footer) ws.headerFooter = { oddHeader: p.header ? hfToFile(p.header) : undefined, oddFooter: p.footer ? hfToFile(p.footer) : undefined };
  if (sheet.protected) ws.sheetProtection = { sheet: true, objects: true, scenarios: true };
  // Resimler
  for (const img of sheet.images) {
    const m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/.exec(img.src);
    if (!m) continue;
    const id = x.addImage({ base64: img.src, extension: m[1] === "jpg" ? "jpeg" : m[1] });
    ws.addImage(id, {
      tl: { col: img.anchor.col + img.anchor.dx / Math.max(1, sheet.cols.size(img.anchor.col)), row: img.anchor.row + img.anchor.dy / Math.max(1, sheet.rows.size(img.anchor.row)) },
      ext: { width: img.width, height: img.height },
      editAs: "oneCell",
    });
  }
}

function cellResult(v: Scalar): any {
  if (v === null) return undefined;
  if (isError(v)) return { error: v.code };
  return v;
}

function writeValidation(r: ValidationRule): any {
  const out: any = { type: r.type === "any" ? "any" : r.type, allowBlank: r.allowBlank !== false, showErrorMessage: true, showInputMessage: !!r.prompt };
  if (r.operator && r.type !== "list" && r.type !== "custom") out.operator = r.operator;
  const f = (s?: string) => (s === undefined ? undefined : s.startsWith("=") ? s.slice(1) : s);
  if (r.type === "list") {
    const f1 = r.formula1 ?? "";
    out.formulae = [f1.startsWith("=") ? f1.slice(1) : `"${f1.split(";").map((x) => x.trim()).join(",")}"`];
    if (r.showDropdown === false) out.showDropDown = true;
  } else {
    out.formulae = [f(r.formula1), f(r.formula2)].filter((x) => x !== undefined).map((x) => (x !== undefined && /^-?\d+([.,]\d+)?$/.test(x) ? Number(x.replace(",", ".")) : x));
  }
  if (r.error) out.error = r.error;
  if (r.errorTitle) out.errorTitle = r.errorTitle;
  if (r.errorStyle) out.errorStyle = r.errorStyle;
  if (r.prompt) out.prompt = r.prompt;
  if (r.promptTitle) out.promptTitle = r.promptTitle;
  return out;
}

function writeCondRule(r: CondRule, priority: number): any {
  const style = r.style ? toXStyle(r.style) : undefined;
  if (style) {
    // Koşullu biçimde yazı tipi adı/boyutu yazılmaz
    if (style.font) {
      delete style.font.name;
      delete style.font.size;
      delete style.font.family;
      delete style.font.scheme;
      if (!Object.keys(style.font).length) delete style.font;
    }
    if (style.fill) style.fill = { type: "pattern", pattern: "solid", bgColor: style.fill.fgColor };
  }
  const fx = (v: string) => addXlfnPrefixes(v.startsWith("=") ? v.slice(1) : v);
  switch (r.type) {
    case "cellIs":
      return { type: "cellIs", operator: r.operator, formulae: (r.values ?? []).map(fx), priority, style };
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith":
      return { type: "containsText", operator: r.type === "containsText" ? "containsText" : r.type === "notContainsText" ? "notContains" : r.type, text: r.text, priority, style };
    case "containsBlanks":
    case "notContainsBlanks":
    case "containsErrors":
      return { type: "containsText", operator: r.type, priority, style };
    case "duplicateValues":
    case "uniqueValues":
      return { type: r.type, priority, style };
    case "top10":
      return { type: "top10", rank: r.rank ?? 10, percent: !!r.percent, bottom: !!r.bottom, priority, style };
    case "aboveAverage":
      return { type: "aboveAverage", aboveAverage: !r.below, priority, style };
    case "expression":
      return { type: "expression", formulae: (r.values ?? []).map(fx), priority, style };
    case "timePeriod":
      return { type: "timePeriod", timePeriod: r.timePeriod, priority, style };
    case "dataBar":
      return { type: "dataBar", priority, gradient: true, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF" + (r.color ?? "#638EC6").slice(1) } };
    case "colorScale": {
      const cols = r.colors ?? ["#F8696B", "#FFEB84", "#63BE7B"];
      const cfvo = cols.length === 2 ? [{ type: "min" }, { type: "max" }] : [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }];
      return { type: "colorScale", priority, cfvo, color: cols.map((c) => ({ argb: "FF" + c.slice(1) })) };
    }
  }
  return null;
}

export { colToName, resolveFunctionName, stripFunctionPrefix };
