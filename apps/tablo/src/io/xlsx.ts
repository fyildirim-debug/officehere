/**
 * XLSX okuma/yazma (exceljs, MIT). Değerler, formüller (önbellek sonuçlarıyla),
 * stiller, sayı biçimleri, birleştirmeler, boyutlar, gizli satır/sütunlar,
 * bölmeleri dondurma, otomatik filtre, notlar, köprüler, koşullu biçimlendirme,
 * veri doğrulama, adlandırılmış aralıklar, tablolar, resimler ve grafikler.
 */
import { readPageBreaksFromData } from "./xlsxBreaks";
import { addr, cellKey, colToName, errorFromCode, formatRange, isError, nameToCol, parseRangeAddress, resolveFunctionName, stripFunctionPrefix, type Scalar } from "@officehere/formula";
import { MAX_COLS, MAX_ROWS } from "../model/rect";
import { widthCharsToPx, widthPxToChars, ptToPx, pxToPt } from "../model/axis";
import type { PageSetup, Rect, ValidationRule } from "../model/types";
import { WorkbookModel, type SheetModel } from "../model/workbook";
import { ColorResolver, fromXStyle, rawNumberFormats, toXStyle } from "./xlsxStyles";
import { readCharts } from "./xlsxCharts";
import { postProcess } from "./xlsxPost";
import { compressCells, readCondRule, readValidation, writeCondRule } from "./xlsxRules";

import { usesDynamicFunctions } from "./xlsxDynamic";
import { readCheckboxesFromData } from "./xlsxCheckbox";
import { addXlfnPrefixes, normalizeFileFormula, utcSerial } from "./xlsxFormula";

export { compressCells, addXlfnPrefixes, normalizeFileFormula };
import { readPivotsFromData } from "./xlsxPivot";
import { OPTIONAL_PARTS, readCommentsFromData, readDefinedNamesFromData, stripOptionalParts } from "./xlsxPackage";
import { prepareXlsx } from "./xlsxGuard";
import { BIG_MERGE_CELLS, CELL_BUDGET, clampMerge, disjointMerges, MergeLookup, usedExtent } from "./bookUtil";
import { readTablesFromData } from "./xlsxTables";
import { loadZip, releaseZip, seedZip } from "./zipCache";
import { readSparklinesFromData } from "./xlsxSpark";
import { readWhatIfFromData } from "./xlsxWhatIf";
import { readFiltersFromData } from "./xlsxFilters";
import { readSlicersFromData } from "./xlsxSlicers";
import { parseThemeXml } from "../model/theme";
import { activateTheme } from "../model/themeRuntime";

/* eslint-disable @typescript-eslint/no-explicit-any */

async function loadExcelJS(): Promise<any> {
  const mod: any = await import("exceljs");
  return mod.default ?? mod;
}

const PAPER: Record<number, PageSetup["paper"]> = { 1: "Letter", 5: "Legal", 8: "A3", 9: "A4", 11: "A5" };
const PAPER_ID: Record<string, number> = { Letter: 1, Legal: 5, A3: 8, A4: 9, A5: 11 };


function rectsFromSqref(s: string): Rect[] {
  return s
    .split(/\s+/)
    .map((p) => parseRangeAddress(p))
    .filter((r): r is Rect => !!r);
}

/* ============================== OKUMA ============================== */

/** exceljs ile yükler; başarısız olursa isteğe bağlı parçaları sırayla çıkararak yeniden dener. */
async function loadTolerant(ExcelJS: any, data: ArrayBuffer, warnings: string[]): Promise<any> {
  let firstError: unknown;
  try {
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(data);
    return x;
  } catch (e) {
    firstError = e;
  }
  const removed: string[] = [];
  for (const part of OPTIONAL_PARTS) {
    removed.push(part.kind);
    try {
      const stripped = await stripOptionalParts(data, removed);
      const x = new ExcelJS.Workbook();
      await x.xlsx.load(stripped);
      for (const k of removed) warnings.push(OPTIONAL_PARTS.find((p) => p.kind === k)!.label);
      return x;
    } catch {
      /* bir sonraki parça türünü de çıkarıp dene */
    }
  }
  throw firstError;
}

/** Okuma raporu: yapısal onarım yapıldıysa kullanıcıya gösterilecek ileti. */
export interface XlsxReadReport {
  notice?: string;
}

export async function readXlsx(input: ArrayBuffer, report?: XlsxReadReport): Promise<WorkbookModel> {
  const ExcelJS = await loadExcelJS();
  // Paket denetimi ve onarımı (sınırlar, eksik yapı, sayfa XML'i); ölümcül sorunlarda anlaşılır hata
  const pkg = await prepareXlsx(input);
  const data = pkg.data;
  seedZip(data, pkg.zip);
  try {
    return await readPrepared(ExcelJS, pkg, report);
  } finally {
    releaseZip(data);
  }
}

async function readPrepared(ExcelJS: any, pkg: Awaited<ReturnType<typeof prepareXlsx>>, report?: XlsxReadReport): Promise<WorkbookModel> {
  const data = pkg.data;
  const warnings: string[] = [...pkg.warnings];
  const notices: string[] = pkg.notice ? [pkg.notice] : [];
  const x = await loadTolerant(ExcelJS, pkg.excel, warnings);
  const wb = new WorkbookModel(true);
  wb.loadWarnings = warnings;
  const cr = new ColorResolver();
  cr.loadTheme(x._themes?.theme1);
  // Tema: hücre stilleri okunmadan önce gövde yazı tipi bilinmeli
  const pt = parseThemeXml(x._themes?.theme1);
  if (pt?.colors) {
    const clean = (n: string | undefined, fb: string) => (n && !/office/i.test(n) ? n : fb);
    wb.theme = {
      name: clean(pt.name, "Dosya Teması"),
      colors: pt.colors,
      colorsName: clean(pt.colorsName, "Dosya Renkleri"),
      majorFont: pt.majorFont ?? wb.theme.majorFont,
      minorFont: pt.minorFont ?? wb.theme.minorFont,
      fontsName: clean(pt.fontsName, "Dosya Yazı Tipleri"),
    };
  }
  activateTheme(wb.theme);
  wb.props = { title: x.title || undefined, subject: x.subject || undefined, author: x.creator || undefined, company: x.company || undefined, created: x.created, modified: x.modified };

  const worksheets: any[] = [...x.worksheets].sort((a: any, b: any) => (a.orderNo ?? 0) - (b.orderNo ?? 0));
  let rawFormats: Map<string, string> | undefined;
  try {
    rawFormats = rawNumberFormats(await (await loadZip(data)).file("xl/styles.xml")?.async("string"));
  } catch {
    /* özgün biçim kodları isteğe bağlı */
  }
  for (const ws of worksheets) {
    const sheet = wb.addSheet(ws.name);
    // Birleştirmeler exceljs'e verilmez (xlsxGuard: kendi taramamız, geçerli ve çakışmasız)
    const merges = pkg.merges.get(ws.name) ?? [];
    try {
      readSheet(wb, sheet, ws, cr, x, warnings, rawFormats, merges);
    } catch {
      warnings.push(`"${ws.name}" sayfasının bir bölümü`);
    }
    // Dev birleştirmeler kullanılan alana kırpılır, yine dev ise atılır (birbirleriyle çakışmadıkları
    // için kırpılmışları da çakışmaz)
    let used: { rows: number; cols: number } | undefined;
    for (const m of merges) {
      if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= BIG_MERGE_CELLS) {
        sheet.merges.push(m);
        continue;
      }
      used ??= usedExtent(wb, sheet);
      const c = clampMerge(m, used);
      if (c) {
        sheet.merges.push(c);
        if (c.r2 !== m.r2 || c.c2 !== m.c2) warnings.push("çok büyük birleştirilmiş alanlar (kullanılan alana kırpıldı)");
      } else warnings.push("çok büyük birleştirilmiş alanlar");
    }
  }
  // exceljs'in sessizce atladığı sayfa kalmasın (ör. ayrıştırılamayan sayfa parçası)
  const loaded = new Set(worksheets.map((ws) => ws.name));
  for (const name of pkg.sheetNames) if (!loaded.has(name)) warnings.push(`"${name}" sayfası (okunamadı)`);
  // Tablolar (exceljs'in tablo okuyucusu güvenilmez; paket doğrudan okunur)
  try {
    if (await readTablesFromData(data, wb)) warnings.push("tabloların bir bölümü");
  } catch {
    warnings.push("tablolar");
  }
  // Adlandırılmış aralıklar (sabit/formül adları da): paket doğrudan okunur, olmazsa exceljs modeli
  let names: { name: string; formula: string; localSheetId?: number }[] = [];
  try {
    names = await readDefinedNamesFromData(data);
  } catch {
    names = (x.definedNames?.model ?? []).filter((dn: any) => dn.name && !dn.name.startsWith("_xlnm") && dn.ranges?.length).map((dn: any) => ({ name: dn.name, formula: dn.ranges[0] }));
  }
  for (const dn of names) {
    try {
      const scope = dn.localSheetId !== undefined ? wb.sheets[dn.localSheetId]?.id : undefined;
      wb.engine.defineName(dn.name, "=" + normalizeFileFormula(dn.formula), scope);
    } catch {
      /* geçersiz ad */
    }
  }
  if (!wb.sheets.length) wb.addSheet("Sayfa1");
  const firstVisible = wb.sheets.findIndex((s) => !s.hidden);
  wb.activeIndex = Math.max(0, x.views?.[0]?.activeTab ?? firstVisible);
  if (wb.sheets[wb.activeIndex]?.hidden) wb.activeIndex = Math.max(0, firstVisible);
  // Notlar: exceljs bazı üreticilerin düz <t> metinlerini okumaz; paket doğrudan okunur
  try {
    const notes = await readCommentsFromData(data);
    for (const sheet of wb.sheets) {
      for (const n of notes.get(sheet.name) ?? []) {
        const a = parseRangeAddress(n.ref);
        if (!a || !n.text) continue;
        const k = cellKey(a.r1, a.c1);
        if (sheet.notes.get(k)?.text) continue;
        sheet.notes.set(k, { text: n.text.replace(/^[^:\n]{1,40}:\n/, ""), author: n.author });
      }
    }
  } catch {
    /* exceljs'in okuduğu notlar kalır */
  }
  // Grafikler (exceljs okumaz; DrawingML doğrudan ayrıştırılır)
  try {
    await readCharts(data, wb);
  } catch {
    warnings.push("grafikler");
  }
  wb.syncTables();
  wb.engine.recalc(true);
  // PivotTable tanımları (kaynak değerleri hesaplandıktan sonra)
  try {
    await readPivotsFromData(data, wb);
  } catch {
    if (!warnings.includes("özet tablolar")) warnings.push("özet tablolar");
  }
  try {
    await readSparklinesFromData(data, wb);
  } catch {
    warnings.push("mini grafikler");
  }
  try {
    await readPageBreaksFromData(data, wb);
  } catch {
    warnings.push("sayfa sonları");
  }
  try {
    await readWhatIfFromData(data, wb);
  } catch {
    warnings.push("senaryolar ve veri tabloları");
  }
  try {
    await readFiltersFromData(data, wb);
    await readSlicersFromData(data, wb);
  } catch {
    warnings.push("filtre ölçütleri ve dilimleyiciler");
  }
  try {
    await readCheckboxesFromData(data, wb);
  } catch {
    warnings.push("onay kutuları");
  }
  // Sessiz kayıp denetimi: sayfalarda veri vardı ama hiçbir hücre okunamadıysa uyarılır
  if (pkg.hadCells && wb.sheets.every((s) => countCells(wb, s) === 0)) {
    warnings.push("sayfa verileri");
    notices.push("Dosyadaki hücre verileri okunamadı; dosya bozuk olabilir. Kaydederseniz özgün içerik kaybolabilir.");
  }
  wb.loadWarnings = [...new Set(warnings)];
  if (report && notices.length) report.notice = notices.join(" ");
  return wb;
}

function countCells(wb: WorkbookModel, s: SheetModel): number {
  let n = 0;
  wb.engine.forEachCell(s.id, () => {
    n++;
  });
  return n;
}



function readSheet(wb: WorkbookModel, sheet: SheetModel, ws: any, cr: ColorResolver, x: any, warnings: string[], rawFormats?: Map<string, string>, merges: Rect[] = []): void {
  if (ws.state && ws.state !== "visible") sheet.hidden = true;
  const tab = cr.resolve(ws.properties?.tabColor);
  if (tab) sheet.tabColor = tab;
  if (ws.properties?.defaultColWidth) sheet.cols.setDefault(widthCharsToPx(ws.properties.defaultColWidth));
  if (ws.properties?.defaultRowHeight && Math.abs(ws.properties.defaultRowHeight - 15) > 0.1) sheet.rows.setDefault(ptToPx(ws.properties.defaultRowHeight));
  // Sütunlar
  // exceljs'in columnCount'u tüm satırları dolaşır; tanımlı sütunlar yeterli (sınırla)
  const colCount = Math.min(ws._columns?.length ?? ws.columns?.length ?? 0, MAX_COLS);
  for (let c = 1; c <= colCount; c++) {
    const col = ws.getColumn(c);
    if (col.width !== undefined && col.width !== null && (col as any)._worksheet) {
      // Excel en çok 255 karakter genişlik kabul eder (bozuk dosyada 2^31 olabilir)
      const w = widthCharsToPx(Math.min(255, Math.max(0, Number(col.width) || 0)));
      if (Math.abs(w - sheet.cols.defaultSize) > 0.5) sheet.cols.set(c - 1, w);
    }
    if (col.hidden) sheet.cols.setHidden(c - 1, true);
    if (col.outlineLevel) sheet.cols.setLevel(c - 1, col.outlineLevel);
    const st = col.style && Object.keys(col.style).length ? fromXStyle(col.style, cr, rawFormats) : null;
    if (st && Object.keys(st).length) sheet.colStyles.set(c - 1, wb.styles.intern(st));
  }
  const notesCells: { row: number; col: number; text: string; author?: string }[] = [];
  const arrayAreas: (Rect & { row: number; col: number })[] = [];
  const validations = new Map<string, { rule: ValidationRule; cells: [number, number][]; rects: Rect[] }>();
  // Birleştirilmiş alanın gizli hücreleri yalnız stil taşır (exceljs birleştirirken değerlerini atardı);
  // dev birleştirmelerde değerler korunur (kırpılıp kullanılan alanın tamamını kaplayabilirler)
  const small = merges.filter((m) => (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= BIG_MERGE_CELLS);
  const lookup = small.length ? new MergeLookup(small) : null;
  const hiddenInMerge = (r: number, c: number) => {
    const m = lookup?.at(r, c);
    return !!m && (m.r1 !== r || m.c1 !== c);
  };
  // Satırlar seyrek dizide durur: yalnız var olanlar dolaşılır (exceljs'in eachRow'u boşlukları da
  // nesneye çevirir; bozuk dosyada satır 2^31 olunca takılırdı). Hücreler sütun sınırıyla dolaşılır.
  const rowList: any[] = ws._rows ?? [];
  const eachRow = (fn: (row: any, rn: number) => void) => {
    if (!ws._rows) return ws.eachRow({ includeEmpty: true }, fn);
    for (const key of Object.keys(rowList)) {
      const row = rowList[Number(key)];
      if (row) fn(row, row.number ?? Number(key) + 1);
    }
  };
  eachRow((row: any, rn: number) => {
    const r = rn - 1;
    if (!(r >= 0) || r >= MAX_ROWS) return;
    if (row.height !== undefined && row.height !== null && (row.customHeight || Math.abs(row.height - 15) > 0.1)) sheet.rows.set(r, ptToPx(Math.min(409.5, Math.max(0, Number(row.height) || 0))), !!row.customHeight);
    if (row.hidden) sheet.rows.setHidden(r, true);
    if (row.outlineLevel) sheet.rows.setLevel(r, row.outlineLevel);
    if (row.collapsed) sheet.rows.setCollapsed(r, true);
    const nCells = Math.min(row._cells?.length ?? row.cellCount ?? 0, MAX_COLS);
    for (let cn = 1; cn <= nCells; cn++) {
      const cell = row.getCell(cn);
      const c = cn - 1;
      const hidden = hiddenInMerge(r, c);
      readCell(wb, sheet, cell, r, c, cr, rawFormats, hidden);
      // Dizi formülünün (dinamik dizi) kapladığı alan: ana hücre dışındakiler yalnız önbellek değeridir
      const v = hidden ? null : cell.value;
      if (v && typeof v === "object" && v.shareType === "array" && typeof v.ref === "string") {
        const ar = parseRangeAddress(v.ref.replace(/\$/g, ""));
        if (ar && (ar.r1 !== ar.r2 || ar.c1 !== ar.c2)) arrayAreas.push({ ...ar, row: r, col: c });
      }
      if (cell.note) {
        const n = cell.note;
        const text = typeof n === "string" ? n : (n.texts ?? []).map((t: any) => t.text).join("");
        if (text) notesCells.push({ row: r, col: c, text: text.replace(/^[^:\n]{1,40}:\n/, "") });
      }
    }
  });
  for (const n of notesCells) sheet.notes.set(cellKey(n.row, n.col), { text: n.text });
  // Taşma alanındaki önbellek değerleri sabit olarak kalırsa formülün taşmasını engeller (#TAŞMA!)
  for (const a of arrayAreas)
    for (let i = a.r1; i <= a.r2; i++)
      for (let j = a.c1; j <= a.c2; j++) if ((i !== a.row || j !== a.col) && !wb.getFormula(sheet, i, j)) wb.setInputRaw(sheet, i, j, null);
  // Veri doğrulama (hücre değeri olmasa da)
  for (const [key, dv] of Object.entries<any>(ws.dataValidations?.model ?? {})) {
    if (!dv?.type) continue;
    const rr = parseRangeAddress(key.replace(/^range:/, "").replace(/\$/g, ""));
    if (!rr) continue;
    const rule = readValidation(dv);
    const k = JSON.stringify(rule);
    let g = validations.get(k);
    if (!g) validations.set(k, (g = { rule, cells: [], rects: [] }));
    // Tek hücreler birleştirilir; aralıklar olduğu gibi (sınırsız boyut) korunur
    if (rr.r1 === rr.r2 && rr.c1 === rr.c2) g.cells.push([rr.r1, rr.c1]);
    else g.rects.push(rr);
  }
  for (const g of validations.values()) sheet.validations.push({ id: wb.newId("dv"), rule: g.rule, ranges: [...g.rects, ...(g.cells.length ? compressCells(g.cells) : [])] });
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
  if (ps.showRowColHeaders) sheet.page.printHeadings = true;
  if (ps.blackAndWhite) sheet.page.blackAndWhite = true;
  if (ps.draft) sheet.page.draft = true;
  if (ps.pageOrder === "overThenDown") sheet.page.pageOrder = "overThenDown";
  const tr = /^\$?(\d+):\$?(\d+)$/.exec(String(ps.printTitlesRow ?? "").replace(/^.*!/, ""));
  if (tr) sheet.page.titleRows = { r1: Number(tr[1]) - 1, r2: Number(tr[2]) - 1 };
  const tc = /^\$?([A-Z]+):\$?([A-Z]+)$/i.exec(String(ps.printTitlesColumn ?? "").replace(/^.*!/, ""));
  if (tc) sheet.page.titleCols = { c1: nameToCol(tc[1].toUpperCase()), c2: nameToCol(tc[2].toUpperCase()) };
  if (ps.fitToPage) {
    const fw = ps.fitToWidth ?? 1;
    const fh = ps.fitToHeight ?? 1;
    sheet.page.fitMode = fw && fh ? "sheet" : fh && !fw ? "rows" : "cols";
  }
  const hf = ws.headerFooter ?? {};
  if (hf.oddHeader) sheet.page.header = stripHF(hf.oddHeader);
  if (hf.oddFooter) sheet.page.footer = stripHF(hf.oddFooter);
  if (ws.sheetProtection?.sheet) sheet.protected = true;
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
    warnings.push("resimler");
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

function readCell(wb: WorkbookModel, sheet: SheetModel, cell: any, r: number, c: number, cr: ColorResolver, rawFormats?: Map<string, string>, hiddenInMerge = false): void {
  const v = cell.value;
  const styleObj = cell.style;
  let input: Scalar | { formula: string } = null;
  let link: string | undefined;
  if (hiddenInMerge || cell.type === 1) {
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
  if (styleObj) {
    const st = Object.keys(styleObj).length ? fromXStyle(styleObj, cr, rawFormats) : {};
    const id = Object.keys(st).length ? wb.styles.intern(st) : 0;
    // Biçimsiz hücre stilli sütunda da kendi (boş) stilini taşır; yoksa sütun stilini devralırdı
    if (id !== (sheet.colStyles.get(c) ?? 0)) sheet.styles.set(cellKey(r, c), id);
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
  // Stil nesneleri stil numarası başına bir kez üretilir: exceljs aynı nesneyi WeakMap'le tanır ve
  // hücre başına yazı tipi/kenarlık XML'i üretip karşılaştırmaz (200 bin satırda yazma süresinin çoğuydu)
  const xStyles = new Map<number, any>();
  const xStyle = (id: number) => {
    let v = xStyles.get(id);
    if (!v) xStyles.set(id, (v = toXStyle(wb.styles.get(id))));
    return v;
  };
  for (const sheet of wb.sheets) writeSheet(wb, sheet, x, xStyle);
  // Tanımlı adlar postProcess içinde workbook.xml'e doğrudan yazılır (sabit/formül adları dahil)
  // exceljs sıkıştırmadan yazar; paket postProcess'te tek kez sıkıştırılır (iki kez DEFLATE yerine)
  const buf = await x.xlsx.writeBuffer({ zip: { compression: "STORE" } });
  let bytes: Uint8Array = new Uint8Array(buf as ArrayBuffer);
  bytes = await postProcess(bytes, wb);
  return new Uint8Array(bytes);
}

function writeSheet(wb: WorkbookModel, sheet: SheetModel, x: any, xStyle: (id: number) => any): void {
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
  const colSet = new Set<number>([...sheet.cols.customItems().map(([c]) => c), ...sheet.cols.hiddenItems(), ...sheet.colStyles.keys(), ...sheet.cols.levelItems().map(([c]) => c)]);
  for (const c of [...colSet].sort((a, b) => a - b)) {
    if (c >= MAX_COLS) continue;
    const col = ws.getColumn(c + 1);
    // Genişliği özel olmayan (yalnız stilli/gizli) sütuna da varsayılan genişlik yazılır: exceljs boş bırakınca 9 karakter yazıyordu
    col.width = widthPxToChars(sheet.cols.isCustom(c) ? sheet.cols.rawSize(c) : sheet.cols.defaultSize);
    if (sheet.cols.isHidden(c)) col.hidden = true;
    if (sheet.cols.level(c)) col.outlineLevel = sheet.cols.level(c);
    const cs = sheet.colStyles.get(c);
    if (cs) col.style = xStyle(cs);
  }
  if (sheet.rows.maxLevel()) ws.properties.outlineLevelRow = sheet.rows.maxLevel();
  if (sheet.cols.maxLevel()) ws.properties.outlineLevelCol = sheet.cols.maxLevel();
  // Satırlar
  const rowSet = new Set<number>([...sheet.rows.customItems().map(([r]) => r), ...sheet.rows.hiddenItems(), ...sheet.rowStyles.keys(), ...sheet.rows.levelItems().map(([r]) => r), ...sheet.rows.collapsedItems()]);
  for (const r of rowSet) {
    const row = ws.getRow(r + 1);
    if (sheet.rows.isCustom(r)) row.height = pxToPt(sheet.rows.rawSize(r));
    if (sheet.rows.isHidden(r)) row.hidden = true;
    if (sheet.rows.level(r)) row.outlineLevel = sheet.rows.level(r);
    const rs = sheet.rowStyles.get(r);
    if (rs) row.style = xStyle(rs);
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
  // Hücreler. Kendi stili olmayan hücreye satır/sütun stilinin (yoksa boş) paylaşılan nesnesi verilir:
  // exceljs her hücre için birleştirilmiş yeni nesne üretip stilini baştan XML'e çeviriyordu.
  const done = new Set<number>();
  const EMPTY_X = {};
  wb.engine.forEachCell(sheet.id, (r, c, rec) => {
    const cell = ws.getCell(r + 1, c + 1);
    const k = cellKey(r, c);
    done.add(k);
    if (!sheet.styles.has(k)) {
      const id = wb.styleIdAt(sheet, r, c);
      cell.style = id ? xStyle(id) : EMPTY_X;
    }
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
      } else if (usesDynamicFunctions(rec.formula.text)) {
        // Tek hücrelik dinamik dizi formülü (ör. ÇAPRAZARA): örtük kesişim (@) eklenmesin
        cell.value = { formula: f, result, shareType: "array", ref: addr(r, c) };
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
    ws.getCell(r + 1, c + 1).style = xStyle(id);
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
  // Birleştirmeler postProcess'te (writeMerges) doğrudan yazılır: exceljs'in mergeCells'i her
  // birleştirmede öncekilerin hepsini gezer (O(n²)). Gizli hücreler, exceljs'in birleştirmesindeki gibi,
  // değer taşımaz ve ana hücrenin stilini alır (kenarlıklar kenardaki hücrelerden çizilir). Ana hücre
  // biçimsizse yalnız var olan gizli hücreler boşaltılır (boş, biçimsiz hücre zaten yazılmaz). Dev
  // birleştirmelerde ve toplam bütçe aşılınca hücre üretilmez.
  let budget = CELL_BUDGET;
  for (const m of disjointMerges(sheet.merges).merges) {
    const n = (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1);
    if (n > BIG_MERGE_CELLS || n > budget) continue;
    budget -= n;
    const master = ws.findCell(m.r1 + 1, m.c1 + 1) ?? (wb.styleIdAt(sheet, m.r1, m.c1) ? ws.getCell(m.r1 + 1, m.c1 + 1) : undefined);
    const style = master?.style;
    const plain = !style || !Object.keys(style).length;
    for (let r = m.r1; r <= m.r2; r++)
      for (let c = m.c1; c <= m.c2; c++) {
        if (r === m.r1 && c === m.c1) continue;
        const cell = plain ? ws.findCell(r + 1, c + 1) : ws.getCell(r + 1, c + 1);
        if (!cell) continue;
        cell.value = null;
        cell.style = plain ? EMPTY_X : style;
      }
  }
  if (sheet.autoFilter) ws.autoFilter = formatRange(sheet.autoFilter.range);
  // Veri doğrulama: postProcess'te (xlsxExtras) çakışmasız ve sınırsız aralıkla yazılır
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
    ...(p.printHeadings ? { showRowColHeaders: true } : {}),
    ...(p.blackAndWhite ? { blackAndWhite: true } : {}),
    ...(p.draft ? { draft: true } : {}),
    ...(p.pageOrder === "overThenDown" ? { pageOrder: "overThenDown" } : {}),
    ...(p.titleRows ? { printTitlesRow: `${p.titleRows.r1 + 1}:${p.titleRows.r2 + 1}` } : {}),
    ...(p.titleCols ? { printTitlesColumn: `${colToName(p.titleCols.c1)}:${colToName(p.titleCols.c2)}` } : {}),
    ...(p.printArea ? { printArea: formatRange(p.printArea) } : {}),
    ...(p.fitMode === "sheet" ? { fitToPage: true, fitToWidth: 1, fitToHeight: 1 } : p.fitMode === "rows" ? { fitToPage: true, fitToWidth: 0, fitToHeight: 1 } : p.fitMode === "cols" || p.fitToWidth ? { fitToPage: true, fitToWidth: 1, fitToHeight: 0 } : {}),
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

export { colToName, resolveFunctionName, stripFunctionPrefix };
