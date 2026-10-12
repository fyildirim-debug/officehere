/** Durum Çözümlemesi işlemleri: Senaryo Yöneticisi (ekle, düzenle, sil, göster, özet) ve Veri Tablosu. */
import { addr, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import type { CellPos, Rect, Scenario } from "../model/types";
import { computeDataTable, dataTableFromSelection, scenarioResults } from "../model/whatIf";
import { STRUCTURE_MESSAGE, structureLocked } from "../model/protection";

const up = (s: string) => s.toLocaleUpperCase("tr-TR");

/** Senaryoyu ekler ya da (oldName verilirse) değiştirir; hata iletisi ya da null. */
export function saveScenario(ctl: Controller, sc: Scenario, oldName?: string): string | null {
  const sheet = ctl.sheet;
  const name = sc.name.trim();
  if (!name) return "Senaryo adı boş bırakılamaz.";
  if (name.length > 255) return "Senaryo adı en çok 255 karakter olabilir.";
  if (sheet.scenarios.some((s) => up(s.name) === up(name) && (oldName === undefined || up(s.name) !== up(oldName)))) return `“${name}” adlı bir senaryo zaten var. Başka bir ad girin.`;
  if (!sc.cells.length) return "Değişen hücreler başvurusu geçerli değil.";
  if (sc.cells.length > 32) return "Bir senaryoda en çok 32 değişen hücre olabilir.";
  ctl.ops.sheetOp(sheet, oldName === undefined ? "Senaryo Ekle" : "Senaryoyu Düzenle", () => {
    const next = { ...sc, name };
    const i = oldName === undefined ? -1 : sheet.scenarios.findIndex((s) => up(s.name) === up(oldName));
    if (i >= 0) sheet.scenarios[i] = next;
    else sheet.scenarios.push(next);
  });
  return null;
}

export function deleteScenario(ctl: Controller, name: string): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Senaryoyu Sil", () => (sheet.scenarios = sheet.scenarios.filter((s) => up(s.name) !== up(name))));
}

/** Göster: değişen hücrelere senaryonun değerleri yazılır (geri alınabilir). */
export function showScenario(ctl: Controller, name: string): void {
  const sheet = ctl.sheet;
  const sc = sheet.scenarios.find((s) => up(s.name) === up(name));
  if (!sc) return;
  ctl.ops.setInputs(sheet, sc.cells.map((c) => ({ row: c.row, col: c.col, input: c.value })), `Senaryo: ${sc.name}`);
}

/** Değişen hücre başvuru metni ("B2;B3" ya da "B2:B4"). */
export function cellsText(cells: { row: number; col: number }[]): string {
  return cells.map((c) => addr(c.row, c.col)).join(";");
}

/** "B2;B3:B4" → hücre listesi (yinelenenler atılır); geçersizse null. */
export function parseCells(text: string, parse: (s: string) => Rect | null): CellPos[] | null {
  const out: CellPos[] = [];
  const seen = new Set<string>();
  for (const part of text.split(/[;,]/).map((x) => x.trim().replace(/\$/g, "").replace(/^.*!/, "")).filter(Boolean)) {
    const r = parse(part);
    if (!r) return null;
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 100); row++)
      for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 100); col++) {
        const k = `${row}:${col}`;
        if (!seen.has(k)) {
          seen.add(k);
          out.push({ row, col });
        }
      }
  }
  return out.length ? out : null;
}

/**
 * Senaryo Özeti: yeni sayfada değişen hücreler ve sonuç hücrelerinin geçerli
 * değerleri ile her senaryodaki değerleri yan yana listelenir.
 */
export function scenarioSummary(ctl: Controller, results: CellPos[]): void {
  if (structureLocked(ctl.wb)) return ctl.alert(STRUCTURE_MESSAGE);
  const src = ctl.sheet;
  const list = src.scenarios;
  if (!list.length) return;
  const wb = ctl.wb;
  const changing: CellPos[] = [];
  for (const s of list) for (const c of s.cells) if (!changing.some((x) => x.row === c.row && x.col === c.col)) changing.push({ row: c.row, col: c.col });
  const current = changing.map((c) => wb.getValue(src, c.row, c.col));
  const currentRes = results.map((c) => wb.getValue(src, c.row, c.col));
  const resVals = scenarioResults(wb, src, list, results);
  ctl.ops.snapshotOp("Senaryo Özeti", () => {
    let n = 1;
    let name = "Senaryo Özeti";
    while (wb.sheetByName(name)) name = `Senaryo Özeti ${++n}`;
    const sh = wb.addSheet(name, wb.activeIndex);
    sh.showGridlines = false;
    const put = (r: number, c: number, v: Scalar, style?: { bold?: boolean; fill?: string; color?: string }) => {
      wb.setInputRaw(sh, r, c, v);
      if (style) wb.setStyleRaw(sh, r, c, wb.styles.intern(style));
    };
    const head = { bold: true, fill: "#D9E1F2" };
    put(1, 1, "Senaryo Özeti", { bold: true });
    put(2, 2, "Geçerli Değerler:", head);
    list.forEach((s, i) => put(2, 3 + i, s.name, head));
    let r = 4;
    put(r++, 1, "Değişen Hücreler:", { bold: true });
    changing.forEach((c, i) => {
      put(r, 1, addr(c.row, c.col));
      put(r, 2, current[i]);
      list.forEach((s, j) => {
        const v = s.cells.find((x) => x.row === c.row && x.col === c.col)?.value ?? current[i];
        put(r, 3 + j, v, { fill: "#F2F2F2" });
      });
      r++;
    });
    if (results.length) {
      put(r++, 1, "Sonuç Hücreleri:", { bold: true });
      results.forEach((c, i) => {
        put(r, 1, addr(c.row, c.col));
        put(r, 2, currentRes[i]);
        list.forEach((_, j) => put(r, 3 + j, resVals[j]?.[i] ?? null));
        r++;
      });
    }
    put(r + 1, 1, "Not: Geçerli Değerler sütunu, rapor oluşturulduğu andaki değerleri gösterir; senaryo değerleri gri hücrelerdedir.");
    sh.cols.set(1, 140);
    sh.cols.set(2, 110);
    for (let j = 0; j < list.length; j++) sh.cols.set(3 + j, 100);
    wb.activeIndex = wb.sheets.indexOf(sh);
  });
  ctl.emit();
}

/** Veri Tablosu oluşturur (Durum Çözümlemesi > Veri Tablosu); hata iletisi ya da null. */
export function createDataTable(ctl: Controller, sel: Rect, rowInput?: CellPos, colInput?: CellPos): string | null {
  const sheet = ctl.sheet;
  const def = dataTableFromSelection(sel, rowInput, colInput);
  if (typeof def === "string") return def;
  if (sheet.dataTables.some((d) => d.ref.r1 <= def.ref.r2 && def.ref.r1 <= d.ref.r2 && d.ref.c1 <= def.ref.c2 && def.ref.c1 <= d.ref.c2)) return "Seçili aralık başka bir veri tablosuyla çakışıyor.";
  ctl.ops.snapshotOp("Veri Tablosu", () => {
    const dt = { ...def, id: ctl.wb.newId("dt") };
    const vals = computeDataTable(ctl.wb, sheet, dt);
    vals.forEach((line, i) => line.forEach((v, j) => ctl.wb.setInputRaw(sheet, dt.ref.r1 + i, dt.ref.c1 + j, v)));
    sheet.dataTables.push(dt);
  });
  return null;
}
