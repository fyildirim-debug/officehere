/**
 * Tablo otomatik davranışları (Otomatik Düzelt > "Tabloya yeni satır ve
 * sütunlar ekle", "Hesaplanmış sütunlar oluşturmak için formülleri doldur"):
 * - Tablonun hemen altına/sağına yazılan değer tabloyu genişletir.
 * - Boş bir tablo sütununa yazılan formül tüm satırlara doldurulur.
 * - Yeni satırda hesaplanmış sütun formülleri sürdürülür.
 * - Başlık hücresi değişince sütun adı (ve yapısal başvurular) güncellenir.
 */
import { cellKey, qualifyStructRefs, shiftFormula } from "@officehere/formula";
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { intersects } from "../model/rect";
import type { Rect, TableDef } from "../model/types";
import type { SheetModel } from "../model/workbook";

function dataRows(t: TableDef): [number, number] {
  return [t.ref.r1 + (t.headerRow ? 1 : 0), t.ref.r2 - (t.totalRow ? 1 : 0)];
}

function syncFilter(t: TableDef): void {
  if (t.filter) t.filter.range = { ...t.ref, r2: t.totalRow ? t.ref.r2 - 1 : t.ref.r2 };
}

/** Sütun hesaplanmış sütun mu (son iki veri satırı aynı göreli formülü taşıyor)? */
function calculatedFormula(ctl: Controller, sheet: SheetModel, t: TableDef, col: number, lastRow: number): string | null {
  const [d1] = dataRows(t);
  const f = ctl.wb.getFormula(sheet, lastRow, col);
  if (!f) return null;
  if (lastRow - 1 >= d1) {
    const prev = ctl.wb.getFormula(sheet, lastRow - 1, col);
    if (!prev || shiftFormula(prev, 1, 0) !== f) return null;
  }
  return f;
}

/**
 * Tablo bir satır büyüyünce son veri satırına uygulanmış veri doğrulama ve
 * koşullu biçim aralıkları yeni satırı da kapsar (Excel'deki gibi).
 */
export function growRulesDown(sheet: SheetModel, t: TableDef, last: number, newRow: number): void {
  const grow = (r: { r1: number; c1: number; r2: number; c2: number }) => (r.r2 === last && r.r1 >= t.ref.r1 && r.c1 >= t.ref.c1 && r.c2 <= t.ref.c2 ? { ...r, r2: Math.max(r.r2, newRow) } : r);
  for (const v of sheet.validations) v.ranges = v.ranges.map(grow);
  for (const c of sheet.condFormats) c.ranges = c.ranges.map(grow);
}

/** Yeni bir başlık adı: boşsa "SütunN", aynı ad varsa sonuna sayı eklenir. */
function columnName(t: TableDef, wanted: string): string {
  let base = wanted.trim() || `Sütun${t.columns.length + 1}`;
  const taken = (n: string) => t.columns.some((c) => c.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR"));
  if (!taken(base)) return base;
  const stem = base;
  let k = 2;
  while (taken(`${stem}${k}`)) k++;
  base = `${stem}${k}`;
  return base;
}

/**
 * Tabloyu `next` aralığına getirir (değişiklik işleminin içinde çağrılır). Yeni sütunlar başlık
 * hücresinden ad alır; yeni satırlarda boş hücrelere hesaplanmış sütun formülleri sürer ve
 * (`inheritStyle` ise) son veri satırının biçimi gelir; doğrulama ve koşullu biçim kuralları büyür.
 */
function applyResize(ctl: Controller, sheet: SheetModel, cur: TableDef, next: Rect, inheritStyle: boolean): void {
  const { wb } = ctl;
  const old = cur.ref;
  const [d1] = dataRows(cur);
  // Sütunlar: eski aralıktakiler adını korur, yeniler başlık hücresinden ad alır
  if (next.c1 !== old.c1 || next.c2 !== old.c2) {
    const columns: string[] = [];
    for (let c = next.c1; c <= next.c2; c++) {
      if (c >= old.c1 && c <= old.c2) {
        columns.push(cur.columns[c - old.c1]);
        continue;
      }
      const text = cur.headerRow ? displayText(wb, sheet, next.r1, c) : "";
      const name = columnName({ ...cur, columns: [...columns, ...cur.columns.slice(Math.max(0, c - old.c1))] }, text);
      if (cur.headerRow && text !== name) wb.setInputRaw(sheet, next.r1, c, name);
      columns.push(name);
    }
    if (cur.filter) {
      const shift = old.c1 - next.c1;
      const crit: typeof cur.filter.columns = {};
      for (const [k, v] of Object.entries(cur.filter.columns)) {
        const i = Number(k) + shift;
        if (i >= 0 && i < columns.length) crit[i] = v;
      }
      cur.filter.columns = crit;
    }
    cur.columns = columns;
  }
  cur.ref = { ...next };
  syncFilter(cur);
  if (next.r2 <= old.r2) return;
  growRulesDown(sheet, { ...cur, ref: old }, old.r2, next.r2);
  const last = old.r2;
  for (let c = Math.max(next.c1, old.c1); c <= Math.min(next.c2, old.c2); c++) {
    const f = last >= d1 ? calculatedFormula(ctl, sheet, { ...cur, ref: old }, c, last) : null;
    const st = inheritStyle && last >= d1 ? sheet.styles.get(cellKey(last, c)) : undefined;
    for (let r = last + 1; r <= next.r2; r++) {
      if (f && !wb.engine.hasInput(sheet.id, r, c)) wb.setInputRaw(sheet, r, c, { formula: shiftFormula(f, r - last, 0) });
      if (st !== undefined && !sheet.styles.has(cellKey(r, c))) wb.setStyleRaw(sheet, r, c, st);
    }
  }
}

/** Boyutlandırmada yeni katılan alan (geri alma için hücre görüntüsü alınır). */
function addedArea(old: Rect, next: Rect): Rect[] {
  const out: Rect[] = [];
  if (next.r2 > old.r2) out.push({ r1: old.r2 + 1, c1: next.c1, r2: next.r2, c2: next.c2 });
  if (next.c2 > old.c2) out.push({ r1: next.r1, c1: old.c2 + 1, r2: Math.min(next.r2, old.r2), c2: next.c2 });
  if (next.c1 < old.c1) out.push({ r1: next.r1, c1: next.c1, r2: Math.min(next.r2, old.r2), c2: old.c1 - 1 });
  return out;
}

/**
 * Tabloyu Yeniden Boyutlandır (Tablo Tasarımı › Özellikler ve tablonun köşe tutamacı). Başlık
 * satırı aynı satırda kalmalı, yeni aralık eskisiyle örtüşmeli ve başka bir tabloyla ya da
 * birleştirilmiş hücreyle çakışmamalı. Yapılamıyorsa nedeni döner.
 */
export function resizeTable(ctl: Controller, t: TableDef, next: Rect): string | null {
  const sheet = ctl.sheet;
  if (next.r1 !== t.ref.r1) return "Üst bilgiler aynı satırda kalmalı ve yeni tablo aralığı özgün tablo aralığıyla örtüşmelidir.";
  if (next.r2 < t.ref.r1 + (t.headerRow ? 1 : 0) + (t.totalRow ? 1 : 0)) return "Tabloda üst bilgi satırının altında en az bir satır olmalıdır.";
  if (next.c2 < t.ref.c1 || next.c1 > t.ref.c2 || next.c2 < next.c1) return "Üst bilgiler aynı satırda kalmalı ve yeni tablo aralığı özgün tablo aralığıyla örtüşmelidir.";
  if (t.totalRow && next.r2 !== t.ref.r2) return "Toplam satırı olan bir tablonun satır sayısını değiştirmek için önce Toplam Satırı seçeneğini kapatın.";
  if (sheet.tables.some((o) => o.id !== t.id && intersects(o.ref, next))) return "Tablolar başka bir tabloyla çakışamaz.";
  const added = addedArea(t.ref, next);
  if (added.some((a) => sheet.mergesIn(a).length > 0)) return "Tablo birleştirilmiş hücreleri içeremez.";
  if (next.r1 === t.ref.r1 && next.c1 === t.ref.c1 && next.r2 === t.ref.r2 && next.c2 === t.ref.c2) return null;
  ctl.ops.sheetOp(
    sheet,
    "Tabloyu Yeniden Boyutlandır",
    () => applyResize(ctl, sheet, sheet.tables.find((x) => x.id === t.id)!, next, true),
    added,
  );
  ctl.wb.syncTables();
  return null;
}

/**
 * Yapıştırılan ya da doldurulan aralık tablonun hemen altına (ya da sağına) taşıyorsa tablo
 * o aralığı kapsayacak kadar büyür ("Tabloya yeni satır ve sütunlar ekle"). Toplam satırı olan
 * tablolar büyümez. Çağıran, işlemi önceki adımla tek geri alma adımında birleştirir.
 */
export function expandTablesForRange(ctl: Controller, sheet: SheetModel, dst: Rect): boolean {
  let changed = false;
  for (const t of [...sheet.tables]) {
    if (t.totalRow) continue;
    const cols = dst.c1 <= t.ref.c2 && dst.c2 >= t.ref.c1;
    const rows = dst.r1 <= t.ref.r2 && dst.r2 >= t.ref.r1;
    const down = cols && dst.r2 > t.ref.r2 && dst.r1 <= t.ref.r2 + 1 && dst.r1 >= t.ref.r1;
    const right = rows && dst.c2 > t.ref.c2 && dst.c1 <= t.ref.c2 + 1 && dst.c1 >= t.ref.c1 && dst.r1 >= t.ref.r1 && dst.r2 <= t.ref.r2;
    if (!down && !right) continue;
    const next: Rect = { ...t.ref, r2: down ? dst.r2 : t.ref.r2, c2: right ? dst.c2 : t.ref.c2 };
    if (sheet.tables.some((o) => o.id !== t.id && intersects(o.ref, next))) continue;
    const added = addedArea(t.ref, next);
    if (added.some((a) => sheet.mergesIn(a).length > 0)) continue;
    ctl.ops.sheetOp(sheet, "Yazma", () => applyResize(ctl, sheet, sheet.tables.find((x) => x.id === t.id)!, next, false), added);
    changed = true;
  }
  if (changed) ctl.wb.syncTables();
  return changed;
}

/** Yazma sonrası tablo davranışları; değişiklik yaptıysa true. Tek geri alma adımına katılır. */
export function afterTableEntry(ctl: Controller, sheet: SheetModel, row: number, col: number): boolean {
  const { wb } = ctl;
  const ops = ctl.ops;
  if (!sheet.tables.length) return false;
  let changed = false;
  // Tabloya katılan hücredeki tablosuz yapısal başvurular ([@Adet]) tablo adıyla nitelenir
  const qualify = (name: string) => {
    const f = wb.getFormula(sheet, row, col);
    if (f && f.includes("[")) wb.setInputRaw(sheet, row, col, { formula: qualifyStructRefs(f, name) });
  };

  // 1) Başlık hücresi: sütun adını güncelle
  const inT = sheet.tableAt(row, col);
  if (inT && inT.headerRow && row === inT.ref.r1) {
    const i = col - inT.ref.c1;
    const old = inT.columns[i];
    const text = displayText(wb, sheet, row, col);
    const next = text.trim() ? text : `Sütun${i + 1}`;
    if (next !== old) {
      ops.snapshotOp("Yazma", () => {
        const cur = sheet.tables.find((x) => x.id === inT.id)!;
        const others = { ...cur, columns: cur.columns.filter((_, j) => j !== i) };
        const name = columnName(others, next);
        if (name !== text) wb.setInputRaw(sheet, row, col, name);
        cur.columns = cur.columns.map((c, j) => (j === i ? name : c));
        wb.engine.renameTableColumn(cur.name, old, name);
      });
      return true;
    }
    return false;
  }

  // 2) Hemen alta yazılan değer: tablo bir satır büyür (toplam satırı yoksa)
  const below = sheet.tables.find((t) => !t.totalRow && t.ref.r2 + 1 === row && col >= t.ref.c1 && col <= t.ref.c2);
  // 3) Hemen sağa yazılan değer: tablo bir sütun büyür
  const right = !below ? sheet.tables.find((t) => t.ref.c2 + 1 === col && row >= t.ref.r1 && row <= t.ref.r2 - (t.totalRow ? 1 : 0)) : undefined;
  const conflict = (t: TableDef, r: { r1: number; c1: number; r2: number; c2: number }) =>
    sheet.tables.some((o) => o.id !== t.id && o.ref.r1 <= r.r2 && r.r1 <= o.ref.r2 && o.ref.c1 <= r.c2 && r.c1 <= o.ref.c2) || sheet.mergesIn(r).length > 0;
  if (below && !conflict(below, { r1: row, c1: below.ref.c1, r2: row, c2: below.ref.c2 })) {
    ops.sheetOp(
      sheet,
      "Yazma",
      () => {
        const cur = sheet.tables.find((x) => x.id === below.id)!;
        const last = cur.ref.r2;
        cur.ref = { ...cur.ref, r2: row };
        syncFilter(cur);
        growRulesDown(sheet, cur, last, row);
        qualify(cur.name);
        // Hesaplanmış sütunlar yeni satırda sürer
        for (let c = cur.ref.c1; c <= cur.ref.c2; c++) {
          if (c === col || wb.engine.hasInput(sheet.id, row, c)) continue;
          const f = calculatedFormula(ctl, sheet, { ...cur, ref: { ...cur.ref, r2: last } }, c, last);
          if (f) wb.setInputRaw(sheet, row, c, { formula: shiftFormula(f, row - last, 0) });
        }
      },
      [{ r1: row, c1: below.ref.c1, r2: row, c2: below.ref.c2 }],
    );
    changed = true;
  } else if (right && !conflict(right, { r1: right.ref.r1, c1: col, r2: right.ref.r2, c2: col })) {
    const header = right.headerRow ? right.ref.r1 : -1;
    ops.sheetOp(
      sheet,
      "Yazma",
      () => {
        const cur = sheet.tables.find((x) => x.id === right.id)!;
        const name = columnName(cur, header >= 0 ? displayText(wb, sheet, header, col) : "");
        if (header >= 0 && displayText(wb, sheet, header, col) !== name) wb.setInputRaw(sheet, header, col, name);
        cur.ref = { ...cur.ref, c2: col };
        cur.columns = [...cur.columns, name];
        syncFilter(cur);
        qualify(cur.name);
      },
      [{ r1: right.ref.r1, c1: col, r2: right.ref.r2, c2: col }],
    );
    changed = true;
  }

  // 4) Boş sütuna yazılan formül: hesaplanmış sütun olarak tüm veri satırlarına doldurulur
  const t = sheet.tableAt(row, col);
  const f = t ? wb.getFormula(sheet, row, col) : undefined;
  if (t && f) {
    const [d1, d2] = dataRows(t);
    if (row >= d1 && row <= d2 && d2 > d1) {
      let empty = true;
      for (let r = d1; r <= d2 && empty; r++) if (r !== row && wb.engine.hasInput(sheet.id, r, col)) empty = false;
      if (empty) {
        const entries: { row: number; col: number; input: { formula: string } }[] = [];
        for (let r = d1; r <= d2; r++) if (r !== row) entries.push({ row: r, col, input: { formula: shiftFormula(f, r - row, 0) } });
        ops.setInputs(sheet, entries, "Yazma");
        changed = true;
      }
    }
  }
  return changed;
}

/**
 * Tablo içinde Tab / Shift+Tab: son sütundan sonraki satırın ilk sütununa
 * (ya da ilk sütundan önceki satırın son sütununa) geçilir; son veri
 * satırının son hücresinde Tab tabloya yeni satır ekler (toplam satırı
 * aşağı kayar, hesaplanmış sütun formülleri yeni satıra sürer).
 */
export function tableTab(ctl: Controller, back: boolean): boolean {
  const { sheet, wb } = ctl;
  const r0 = ctl.range;
  if (ctl.sel.ranges.length !== 1 || r0.r1 !== r0.r2 || r0.c1 !== r0.c2) return false;
  const { row, col } = ctl.sel.active;
  const t = sheet.tableAt(row, col);
  if (!t) return false;
  const [d1, d2] = dataRows(t);
  if (t.totalRow && row === t.ref.r2) return false;
  if (back) {
    if (col !== t.ref.c1 || row <= t.ref.r1) return false;
    ctl.selectCell(row - 1, t.ref.c2);
    return true;
  }
  if (col !== t.ref.c2) return false;
  if (row < d2) {
    ctl.selectCell(row + 1, t.ref.c1);
    return true;
  }
  if (row !== d2 || d2 < d1 - 1) return false;
  const newRow = d2 + 1;
  // Yeni satırın yeri dolu ya da toplam satırıysa tablonun sütunları bir satır aşağı kayar
  let occupied = t.totalRow;
  for (let c = t.ref.c1; c <= t.ref.c2 && !occupied; c++) if (wb.engine.hasInput(sheet.id, newRow, c)) occupied = true;
  const blocked = sheet.tables.some((o) => o.id !== t.id && o.ref.r1 <= newRow && newRow <= o.ref.r2 && o.ref.c1 <= t.ref.c2 && t.ref.c1 <= o.ref.c2) || sheet.mergesIn({ r1: newRow, c1: t.ref.c1, r2: newRow, c2: t.ref.c2 }).length > 0;
  if (blocked) return false;
  ctl.ops.snapshotOp("Tablo Satırı Ekle", () => {
    const cur = sheet.tables.find((x) => x.id === t.id)!;
    if (occupied) ctl.ops.shiftCellsRaw(sheet, { r1: newRow, c1: cur.ref.c1, r2: newRow, c2: cur.ref.c2 }, "down");
    const last = d2;
    cur.ref = { ...cur.ref, r2: cur.ref.r2 + 1 };
    syncFilter(cur);
    growRulesDown(sheet, cur, last, newRow);
    for (let c = cur.ref.c1; c <= cur.ref.c2; c++) {
      const f = last >= d1 ? calculatedFormula(ctl, sheet, { ...cur, ref: { ...cur.ref, r2: last + (cur.totalRow ? 1 : 0) } }, c, last) : null;
      if (f) wb.setInputRaw(sheet, newRow, c, { formula: shiftFormula(f, 1, 0) });
    }
    // Biçim (sayı biçimi vb.) üstteki satırdan devralınır
    if (last >= d1) for (let c = cur.ref.c1; c <= cur.ref.c2; c++) wb.setStyleRaw(sheet, newRow, c, sheet.styles.get(cellKey(last, c)));
  });
  wb.syncTables();
  ctl.selectCell(newRow, t.ref.c1);
  return true;
}
