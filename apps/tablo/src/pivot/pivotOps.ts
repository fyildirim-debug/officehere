/**
 * PivotTable işlemleri: kaynak okuma, rapor hücrelerini sayfaya yazma
 * (biçimleriyle), oluşturma, alan değişikliği, yenileme, taşıma, silme.
 * Her işlem tek geri alma adımıdır (tanım + çıktı hücreleri birlikte).
 */
import { keyCol, keyRow, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { MAX_COLS, MAX_ROWS, contains, intersects, rect } from "../model/rect";
import type { CellStyle } from "../model/styles";
import type { PivotDef, Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { displayText } from "../model/display";
import { computePivot, newPivotDef, type PivotCell, type PivotLayout, type PivotSource } from "./compute";
import { fitColumnWidth } from "../grid/autofit";

/* ------------------------------ Stiller ------------------------------ */

export interface PivotStylePreset {
  id: string;
  label: string;
  headerFill: string;
  headerColor?: string;
  totalFill: string;
  border: string;
  band: string;
}

export const PIVOT_STYLES: PivotStylePreset[] = [
  { id: "PivotStyleLight16", label: "Açık Mavi, Pivot Stil Açık 16", headerFill: "#DDEBF7", totalFill: "#DDEBF7", border: "#9BC2E6", band: "#EEF5FB" },
  { id: "PivotStyleLight15", label: "Açık Gri, Pivot Stil Açık 15", headerFill: "#EDEDED", totalFill: "#EDEDED", border: "#A5A5A5", band: "#F6F6F6" },
  { id: "PivotStyleLight21", label: "Açık Yeşil, Pivot Stil Açık 21", headerFill: "#E2EFDA", totalFill: "#E2EFDA", border: "#A9D08E", band: "#F0F7EC" },
  { id: "PivotStyleLight20", label: "Açık Turuncu, Pivot Stil Açık 20", headerFill: "#FCE4D6", totalFill: "#FCE4D6", border: "#F4B084", band: "#FDF1EA" },
  { id: "PivotStyleMedium2", label: "Mavi, Pivot Stil Orta 2", headerFill: "#4472C4", headerColor: "#FFFFFF", totalFill: "#D9E1F2", border: "#8EA9DB", band: "#E9EEF8" },
  { id: "PivotStyleMedium7", label: "Yeşil, Pivot Stil Orta 7", headerFill: "#70AD47", headerColor: "#FFFFFF", totalFill: "#E2EFDA", border: "#A9D08E", band: "#EFF6EA" },
  { id: "PivotStyleMedium3", label: "Turuncu, Pivot Stil Orta 3", headerFill: "#ED7D31", headerColor: "#FFFFFF", totalFill: "#FCE4D6", border: "#F4B084", band: "#FDF0E8" },
  { id: "PivotStyleDark1", label: "Koyu Gri, Pivot Stil Koyu 1", headerFill: "#404040", headerColor: "#FFFFFF", totalFill: "#D9D9D9", border: "#7F7F7F", band: "#F2F2F2" },
];

export function stylePreset(id: string): PivotStylePreset {
  return PIVOT_STYLES.find((s) => s.id === id) ?? PIVOT_STYLES[0];
}

function cellStyleFor(c: PivotCell, def: PivotDef, lay: PivotLayout): CellStyle {
  const p = stylePreset(def.style);
  const st: CellStyle = {};
  const headerEnd = lay.tableRow + lay.headerRows - 1;
  const inHeader = c.row >= lay.tableRow && c.row <= headerEnd;
  switch (c.kind) {
    case "filterLabel":
      st.bold = true;
      break;
    case "caption":
    case "colLabelsHeader":
    case "rowLabelsHeader":
    case "fieldHeader":
    case "colItem":
    case "grandColHeader":
      st.bold = true;
      st.fill = p.headerFill;
      if (p.headerColor) st.color = p.headerColor;
      if (c.row === headerEnd) st.bBottom = { style: "thin", color: p.border };
      break;
    case "grandLabel":
    case "grandValue":
      st.bold = true;
      st.fill = p.totalFill;
      st.bTop = { style: "thin", color: p.border };
      break;
    case "subtotalLabel":
    case "subtotalValue":
    case "rowParent":
      if (c.bold !== false) st.bold = true;
      break;
  }
  if (c.kind === "grandValue" && inHeader) delete st.bTop;
  if (c.indent) st.indent = c.indent;
  if (c.numFmt) st.numFmt = c.numFmt;
  if (c.kind === "value" && def.bandedRows && c.band && c.band % 2 === 0) st.fill = p.band;
  if (c.kind === "value" && def.bandedCols && (c.col - lay.labelCols) % 2 === 1) st.fill = p.band;
  return st;
}

/* ------------------------------ Kaynak ------------------------------ */

export function readSource(wb: WorkbookModel, def: PivotDef): PivotSource | string {
  const sheet = wb.sheetById(def.sourceSheet);
  if (!sheet) return "Özet tablo kaynak verisinin bulunduğu sayfa artık yok.";
  const r = def.source;
  if (r.r2 <= r.r1 && r.r1 === r.r2) return "Kaynak aralık en az bir başlık satırı ve bir veri satırı içermelidir.";
  const ext = wb.dataExtent(sheet);
  const r2 = Math.min(r.r2, Math.max(r.r1, ext.rows - 1));
  const c2 = Math.min(r.c2, Math.max(r.c1, ext.cols - 1));
  const headers: string[] = [];
  for (let c = r.c1; c <= c2; c++) {
    const h = displayText(wb, sheet, r.r1, c).trim();
    if (!h) return "Özet tablo oluşturulamadı: kaynak aralığın ilk satırındaki her sütunun bir başlığı olmalıdır. Boş başlıklı sütunlara ad verip yeniden deneyin.";
    headers.push(h);
  }
  const records: Scalar[][] = [];
  for (let row = r.r1 + 1; row <= r2; row++) {
    const rec: Scalar[] = [];
    for (let c = r.c1; c <= c2; c++) rec.push(wb.getValue(sheet, row, c));
    records.push(rec);
  }
  const numFmts = headers.map((_, i) => {
    const f = wb.styleAt(sheet, Math.min(r.r1 + 1, r2), r.c1 + i).numFmt;
    return f && f !== "General" ? f : undefined;
  });
  return { headers, records, numFmts };
}

/* ------------------------------ Yerleşim ------------------------------ */

export interface PivotRender {
  layout: PivotLayout;
  /** Çıktının sayfadaki sol üst hücresi (filtreler dahil). */
  origin: { row: number; col: number };
  rect: Rect;
  /** Tablo üst satırı (filtrelerin altında). */
  tableTop: number;
}

export function planPivot(def: PivotDef, src: PivotSource): PivotRender {
  const layout = computePivot(def, src);
  // Filtreler bağlantı hücresinin üstünde; yer yoksa tablo aşağı kayar
  const tableTop = Math.max(def.anchor.row, layout.tableRow);
  const origin = { row: tableTop - layout.tableRow, col: def.anchor.col };
  const rows = layout.empty ? Math.max(layout.rows, layout.tableRow + 1) : layout.rows;
  const rectOut = rect(origin.row, origin.col, Math.min(MAX_ROWS - 1, origin.row + rows - 1), Math.min(MAX_COLS - 1, origin.col + Math.max(1, layout.cols) - 1));
  return { layout, origin, rect: rectOut, tableTop };
}

/** Boş PivotTable'ın yer tutucu alanı (çizimde gösterilir). */
export function placeholderRect(def: PivotDef): Rect {
  return rect(def.anchor.row, def.anchor.col, def.anchor.row + 17, def.anchor.col + 2);
}

function writeOutput(wb: WorkbookModel, sheet: SheetModel, def: PivotDef, plan: PivotRender): void {
  for (const c of plan.layout.cells) {
    const row = plan.origin.row + c.row;
    const col = plan.origin.col + c.col;
    if (row >= MAX_ROWS || col >= MAX_COLS) continue;
    wb.setInputRaw(sheet, row, col, c.value);
    const st = cellStyleFor(c, def, plan.layout);
    wb.setStyleRaw(sheet, row, col, Object.keys(st).length ? wb.styles.intern(st) : undefined);
  }
  // Başlık ve toplam satırlarının boş hücreleri de biçim alır
  const lay = plan.layout;
  if (!lay.empty) {
    const p = stylePreset(def.style);
    const headerEnd = lay.tableRow + lay.headerRows - 1;
    const filled = new Set(lay.cells.map((c) => c.row * 100000 + c.col));
    const grandRow = def.grandRows && def.rows.length ? lay.rows - 1 : -1;
    for (let r = lay.tableRow; r < lay.rows; r++) {
      const isHeader = r <= headerEnd;
      const isGrand = r === grandRow;
      if (!isHeader && !isGrand) continue;
      for (let c = 0; c < lay.cols; c++) {
        if (filled.has(r * 100000 + c)) continue;
        const st: CellStyle = isHeader ? { fill: p.headerFill, ...(r === headerEnd ? { bBottom: { style: "thin", color: p.border } } : {}) } : { fill: p.totalFill, bTop: { style: "thin", color: p.border } };
        wb.setStyleRaw(sheet, plan.origin.row + r, plan.origin.col + c, wb.styles.intern(st));
      }
    }
  }
  setPivotMeta(def, plan);
}

/** Çıktı aralığını ve açılır düğme konumlarını tanıma işler (hücrelere dokunmaz). */
export function setPivotMeta(def: PivotDef, plan: PivotRender): void {
  const lay = plan.layout;
  def.output = plan.rect;
  const btn: { row: number; col: number; field: number }[] = [];
  const o = plan.origin;
  def.filters.forEach((f, i) => btn.push({ row: o.row + i, col: o.col + 1, field: f.field }));
  if (!lay.empty) {
    const hr = lay.tableRow + lay.headerRows - 1;
    const rowF = def.rows.filter((f) => f >= 0);
    const colF = def.cols.filter((f) => f >= 0);
    if (def.layout === "compact") {
      if (rowF.length && def.rowHeaders) btn.push({ row: o.row + hr, col: o.col, field: rowF[0] });
    } else rowF.forEach((f, i) => btn.push({ row: o.row + hr, col: o.col + i, field: f }));
    if (colF.length && def.colHeaders) btn.push({ row: o.row + lay.tableRow, col: o.col + lay.labelCols, field: colF[0] });
  }
  def.buttons = btn;
}

function clearOutput(wb: WorkbookModel, sheet: SheetModel, r: Rect | undefined): void {
  if (!r) return;
  for (const k of wb.cellKeysIn(sheet, r)) {
    const row = keyRow(k);
    const col = keyCol(k);
    wb.setInputRaw(sheet, row, col, null);
    wb.setStyleRaw(sheet, row, col, undefined);
  }
}

/* ------------------------------ Sorgular ------------------------------ */

export function pivotAt(sheet: SheetModel, row: number, col: number): PivotDef | undefined {
  return sheet.pivots.find((p) => {
    const r = p.output && !(p.values.length === 0 && p.rows.length === 0 && p.cols.length === 0) ? p.output : placeholderRect(p);
    return contains(r, row, col) || (p.output ? contains(p.output, row, col) : false);
  });
}

export function pivotOutputs(sheet: SheetModel): Rect[] {
  return sheet.pivots.map((p) => p.output).filter((r): r is Rect => !!r);
}

export function uniquePivotName(wb: WorkbookModel): string {
  const names = new Set(wb.sheets.flatMap((s) => s.pivots.map((p) => p.name.toLocaleUpperCase("tr-TR"))));
  let n = 1;
  while (names.has(`PIVOTTABLE${n}`)) n++;
  return `ÖzetTablo${n}`;
}

/** Hedef alanda PivotTable dışı veri var mı (üzerine yazma onayı için)? */
export function overwritesData(wb: WorkbookModel, sheet: SheetModel, r: Rect, def: PivotDef): boolean {
  const own = def.output;
  for (const k of wb.cellKeysIn(sheet, r)) {
    const row = keyRow(k);
    const col = keyCol(k);
    if (own && contains(own, row, col)) continue;
    if (wb.engine.hasInput(sheet.id, row, col)) return true;
  }
  return false;
}

/* ------------------------------ İşlemler ------------------------------ */

/** Tanımı güncelleyip raporu yeniden çizer (tek geri alma adımı). */
export function applyPivot(ctl: Controller, sheet: SheetModel, id: string, patch: (d: PivotDef) => void, label = "Özet Tablo"): boolean {
  const cur = sheet.pivots.find((p) => p.id === id);
  if (!cur) return false;
  const next: PivotDef = structuredClone(cur);
  patch(next);
  const src = readSource(ctl.wb, next);
  if (typeof src === "string") {
    ctl.alert(src);
    return false;
  }
  const plan = planPivot(next, src);
  const other = sheet.pivots.find((p) => p.id !== id && p.output && intersects(p.output, plan.rect));
  if (other) {
    ctl.alert("Bir özet tablo başka bir özet tablonun üzerine gelemez.");
    return false;
  }
  const doIt = () => {
    const rects = [plan.rect, ...(cur.output ? [cur.output] : [])];
    ctl.ops.sheetOp(
      sheet,
      label,
      () => {
        const live = sheet.pivots.find((p) => p.id === id)!;
        clearOutput(ctl.wb, sheet, live.output);
        Object.assign(live, next);
        if (plan.tableTop !== live.anchor.row) live.anchor = { ...live.anchor, row: plan.tableTop };
        writeOutput(ctl.wb, sheet, live, plan);
        // Güncelleştirmede sütun genişliklerini otomatik sığdır (Excel varsayılanı)
        if (!plan.layout.empty && ctl.sheet === sheet) {
          for (let c = plan.rect.c1; c <= plan.rect.c2; c++) {
            const w = fitColumnWidth(ctl, c);
            if (w !== null) sheet.cols.set(c, Math.max(36, w), false);
          }
        }
      },
      rects,
    );
  };
  if (overwritesData(ctl.wb, sheet, plan.rect, cur)) {
    ctl.openDialog("confirm", { message: "Hedef hücrelerin içeriğini değiştirmek istiyor musunuz?", okLabel: "Tamam", onOk: doIt });
    return false;
  }
  doIt();
  return true;
}

/** Yeni PivotTable oluşturur. target: null → yeni sayfa. */
export function createPivot(ctl: Controller, sourceSheet: SheetModel, source: Rect, target: { sheet: SheetModel; row: number; col: number } | null): void {
  const wb = ctl.wb;
  const probe = newPivotDef("x", "x", sourceSheet.id, source, { row: 2, col: 0 });
  const src = readSource(wb, probe);
  if (typeof src === "string") {
    ctl.alert(src);
    return;
  }
  const name = uniquePivotName(wb);
  const id = wb.newId("pivot");
  if (!target) {
    // Excel: etkin sayfanın önüne yeni sayfa, rapor A3'te
    ctl.ops.snapshotOp("Özet Tablo Ekle", () => {
      const sh = wb.addSheet(undefined, wb.activeIndex);
      const def = newPivotDef(id, name, sourceSheet.id, source, { row: 2, col: 0 });
      def.output = rect(2, 0, 2, 0);
      sh.pivots.push(def);
      wb.activeIndex = wb.sheets.indexOf(sh);
      sh.selection = { ranges: [rect(2, 0)], active: { row: 2, col: 0 }, anchor: { row: 2, col: 0 } };
    });
  } else {
    const sh = target.sheet;
    const def = newPivotDef(id, name, sourceSheet.id, source, { row: target.row, col: target.col });
    def.output = rect(target.row, target.col);
    if (sh.pivots.some((p) => p.output && intersects(p.output, placeholderRect(def)))) {
      ctl.alert("Bir özet tablo başka bir özet tablonun üzerine gelemez.");
      return;
    }
    ctl.ops.sheetOp(sh, "Özet Tablo Ekle", () => {
      sh.pivots.push(def);
    });
    const idx = wb.sheets.indexOf(sh);
    if (idx !== wb.activeIndex) ctl.activateSheet(idx);
    ctl.selectCell(target.row, target.col);
  }
  ctl.showFieldList = true;
  ctl.requestTab = "pivotAnalyze";
  ctl.emit();
}

/** Kaynaktan yeniden hesaplar. */
export function refreshPivot(ctl: Controller, sheet: SheetModel, id: string): boolean {
  return applyPivot(ctl, sheet, id, () => {}, "Yenile");
}

export function refreshAll(ctl: Controller): void {
  const mark = ctl.ops.history.mark();
  for (const s of ctl.wb.sheets) for (const p of [...s.pivots]) applyPivot(ctl, s, p.id, () => {}, "Tümünü Yenile");
  ctl.ops.history.squash(mark, "Tümünü Yenile");
}

export function deletePivot(ctl: Controller, sheet: SheetModel, id: string): void {
  const cur = sheet.pivots.find((p) => p.id === id);
  if (!cur) return;
  ctl.ops.sheetOp(
    sheet,
    "Özet Tabloyu Temizle",
    () => {
      clearOutput(ctl.wb, sheet, cur.output);
      sheet.pivots = sheet.pivots.filter((p) => p.id !== id);
    },
    cur.output ? [cur.output] : undefined,
  );
}

/** Tüm alanları kaldırır (Temizle > Tümünü Temizle). */
export function clearPivotFields(ctl: Controller, sheet: SheetModel, id: string): void {
  applyPivot(ctl, sheet, id, (d) => {
    d.rows = [];
    d.cols = [];
    d.values = [];
    d.filters = [];
    d.hidden = {};
  }, "Tümünü Temizle");
}

/** Raporu başka bir konuma taşır. */
export function movePivot(ctl: Controller, sheet: SheetModel, id: string, to: { row: number; col: number }): void {
  applyPivot(ctl, sheet, id, (d) => {
    d.anchor = { ...to };
  }, "Özet Tabloyu Taşı");
}

/** Hücre PivotTable raporunun parçası mı (düzenlemeyi engellemek için)? */
export function isPivotCell(sheet: SheetModel, row: number, col: number): boolean {
  if (!sheet.pivots.length) return false;
  return sheet.pivots.some((p) => p.output && contains(p.output, row, col) && !(p.values.length === 0 && p.rows.length === 0 && p.cols.length === 0));
}

export const PIVOT_EDIT_MESSAGE = "Özet tablonun bu bölümü doğrudan düzenlenemez. Düzeni Özet Tablo Alanları bölmesinden değiştirin ya da kaynak veriyi güncelleyip Yenile'ye tıklayın.";

/** Izgarada çizilecek PivotTable filtre düğmeleri. */
export function pivotButtons(sheet: SheetModel): { row: number; col: number; field: number; pivotId: string; active: boolean }[] {
  const out: { row: number; col: number; field: number; pivotId: string; active: boolean }[] = [];
  for (const p of sheet.pivots)
    for (const b of p.buttons ?? []) {
      const active = !!p.hidden?.[b.field]?.length || !!p.filters.find((f) => f.field === b.field && f.selected);
      out.push({ ...b, pivotId: p.id, active });
    }
  return out;
}
