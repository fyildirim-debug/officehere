/**
 * Tabloya özgü satır/sütun ekleme-silme ve seçim komutları (sağ tık > Ekle /
 * Sil / Seç, Giriş > Hücreler). Yalnız tablonun sütun/satır bandındaki
 * hücreler kaydırılır; sayfanın geri kalanı yerinde kalır.
 */
import { cellKey, shiftFormula } from "@officehere/formula";
import type { Controller } from "../controller";
import type { Rect, TableDef } from "../model/types";
import { intersects } from "../model/rect";

function dataRows(t: TableDef): [number, number] {
  return [t.ref.r1 + (t.headerRow ? 1 : 0), t.ref.r2 - (t.totalRow ? 1 : 0)];
}

/** Seçimin tablo içindeki satır ve sütun aralığı. */
export function tableSpan(ctl: Controller, t: TableDef): { r1: number; r2: number; c1: number; c2: number } {
  const r = ctl.range;
  return { r1: Math.max(r.r1, t.ref.r1), r2: Math.min(r.r2, t.ref.r2), c1: Math.max(r.c1, t.ref.c1), c2: Math.min(r.c2, t.ref.c2) };
}

/** Banttaki diğer tablolar/birleştirmeler kaydırmayı engelliyor mu? */
function blocked(ctl: Controller, t: TableDef, band: Rect): boolean {
  const sheet = ctl.sheet;
  return sheet.tables.some((o) => o.id !== t.id && intersects(o.ref, band)) || sheet.mergesIn(band).length > 0;
}

function refuse(ctl: Controller): void {
  ctl.alert("Bu işlem çalışma sayfanızdaki bir tablodaki ya da birleştirilmiş hücrelerdeki hücreleri kaydıracağından yapılamıyor.");
}

/** Tablo satırları ekler (seçili satırların üstüne ya da altına). */
export function insertTableRows(ctl: Controller, t: TableDef, below = false): void {
  const { sheet, wb } = ctl;
  const [d1, d2] = dataRows(t);
  const span = tableSpan(ctl, t);
  const count = Math.max(1, span.r2 - span.r1 + 1);
  let at = below ? span.r2 + 1 : Math.max(span.r1, d1);
  if (at > d2 + 1) at = d2 + 1;
  const band: Rect = { r1: at, c1: t.ref.c1, r2: 1048575, c2: t.ref.c2 };
  if (blocked(ctl, t, band)) return refuse(ctl);
  ctl.ops.snapshotOp("Tablo Satırları Ekle", () => {
    const cur = sheet.tables.find((x) => x.id === t.id)!;
    ctl.ops.shiftCellsRaw(sheet, { r1: at, c1: cur.ref.c1, r2: at + count - 1, c2: cur.ref.c2 }, "down");
    cur.ref = { ...cur.ref, r2: cur.ref.r2 + count };
    if (cur.filter) cur.filter.range = { ...cur.ref, r2: cur.totalRow ? cur.ref.r2 - 1 : cur.ref.r2 };
    // Komşu veri satırının biçimi ve hesaplanmış sütun formülleri
    const src = at > d1 ? at - 1 : at + count <= d2 + count ? at + count : -1;
    if (src >= 0)
      for (let c = cur.ref.c1; c <= cur.ref.c2; c++) {
        const st = sheet.styles.get(cellKey(src, c));
        const f = wb.getFormula(sheet, src, c);
        // Hesaplanmış sütun: yapısal başvuru ya da başka bir veri satırında aynı göreli formül
        const other = src > at ? src + 1 : src - 1;
        const calc = !!f && (f.includes("[") || wb.getFormula(sheet, other, c) === shiftFormula(f, other - src, 0));
        for (let r = at; r < at + count; r++) {
          wb.setStyleRaw(sheet, r, c, st);
          if (calc) wb.setInputRaw(sheet, r, c, { formula: shiftFormula(f!, r - src, 0) });
        }
      }
    for (const v of sheet.validations) v.ranges = v.ranges.map((x) => grow(x, at, count, cur));
    for (const cf of sheet.condFormats) cf.ranges = cf.ranges.map((x) => grow(x, at, count, cur));
  });
  wb.syncTables();
  ctl.selectRange({ r1: at, c1: span.c1, r2: at + count - 1, c2: span.c2 });
}

/** Tablo sütunundaki kural aralığı yeni satırları kapsayacak biçimde büyür. */
function grow(r: Rect, at: number, count: number, t: TableDef): Rect {
  if (r.c1 > t.ref.c2 || r.c2 < t.ref.c1) return r;
  if (r.r1 < at && r.r2 >= at - 1) return { ...r, r2: r.r2 + count };
  if (r.r1 >= at) return { ...r, r1: r.r1 + count, r2: r.r2 + count };
  return r;
}

/** Seçili tablo satırlarını siler (son veri satırı silinirse tabloda boş bir satır kalır). */
export function deleteTableRows(ctl: Controller, t: TableDef): void {
  const { sheet, wb } = ctl;
  const [d1, d2] = dataRows(t);
  const span = tableSpan(ctl, t);
  const r1 = Math.max(span.r1, d1);
  const r2 = Math.min(span.r2, d2);
  if (r2 < r1) return;
  const count = r2 - r1 + 1;
  const band: Rect = { r1, c1: t.ref.c1, r2: 1048575, c2: t.ref.c2 };
  if (blocked(ctl, t, band)) return refuse(ctl);
  ctl.ops.snapshotOp("Tablo Satırlarını Sil", () => {
    const cur = sheet.tables.find((x) => x.id === t.id)!;
    ctl.ops.shiftCellsRaw(sheet, { r1, c1: cur.ref.c1, r2, c2: cur.ref.c2 }, "up");
    const all = count === d2 - d1 + 1;
    // Tablo en az bir veri satırı taşır
    cur.ref = { ...cur.ref, r2: cur.ref.r2 - count + (all ? 1 : 0) };
    if (all) ctl.ops.shiftCellsRaw(sheet, { r1: d1, c1: cur.ref.c1, r2: d1, c2: cur.ref.c2 }, "down");
    if (cur.filter) cur.filter.range = { ...cur.ref, r2: cur.totalRow ? cur.ref.r2 - 1 : cur.ref.r2 };
    const shrink = (x: Rect): Rect | null => {
      if (x.c1 > cur.ref.c2 || x.c2 < cur.ref.c1 || x.r2 < r1) return x;
      if (x.r1 > r2) return { ...x, r1: x.r1 - count, r2: x.r2 - count };
      const a = Math.min(x.r1, r1);
      const b = x.r2 - Math.min(count, x.r2 - r1 + 1);
      return b < a ? null : { ...x, r1: a, r2: b };
    };
    for (const v of sheet.validations) v.ranges = v.ranges.map(shrink).filter((x): x is Rect => !!x);
    for (const cf of sheet.condFormats) cf.ranges = cf.ranges.map(shrink).filter((x): x is Rect => !!x);
  });
  wb.syncTables();
  ctl.selectCell(Math.min(r1, sheet.tables.find((x) => x.id === t.id)?.ref.r2 ?? r1), ctl.sel.active.col);
}

/** Tablo sütunları ekler (seçili sütunların soluna ya da sağına). */
export function insertTableCols(ctl: Controller, t: TableDef, right = false): void {
  const { sheet, wb } = ctl;
  const span = tableSpan(ctl, t);
  const count = Math.max(1, span.c2 - span.c1 + 1);
  const at = right ? span.c2 + 1 : span.c1;
  const band: Rect = { r1: t.ref.r1, c1: at, r2: t.ref.r2, c2: 16383 };
  if (blocked(ctl, t, band)) return refuse(ctl);
  ctl.ops.snapshotOp("Tablo Sütunları Ekle", () => {
    const cur = sheet.tables.find((x) => x.id === t.id)!;
    ctl.ops.shiftCellsRaw(sheet, { r1: cur.ref.r1, c1: at, r2: cur.ref.r2, c2: at + count - 1 }, "right");
    const names: string[] = [];
    const taken = (n: string) => [...cur.columns, ...names].some((c) => c.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR"));
    let k = cur.columns.length + 1;
    for (let i = 0; i < count; i++) {
      while (taken(`Sütun${k}`)) k++;
      names.push(`Sütun${k}`);
    }
    const idx = at - cur.ref.c1;
    cur.columns = [...cur.columns.slice(0, idx), ...names, ...cur.columns.slice(idx)];
    cur.ref = { ...cur.ref, c2: cur.ref.c2 + count };
    if (cur.filter) {
      // Ölçütler sütunlarıyla birlikte kayar
      const cols: typeof cur.filter.columns = {};
      for (const [k, c] of Object.entries(cur.filter.columns)) cols[Number(k) >= idx ? Number(k) + count : Number(k)] = c;
      cur.filter = { range: { ...cur.ref, r2: cur.totalRow ? cur.ref.r2 - 1 : cur.ref.r2 }, columns: cols };
    }
    if (cur.headerRow) names.forEach((n, i) => wb.setInputRaw(sheet, cur.ref.r1, at + i, n));
  });
  wb.syncTables();
  ctl.selectRange({ r1: span.r1, c1: at, r2: span.r2, c2: at + count - 1 });
}

/** Seçili tablo sütunlarını siler (tüm sütunlar silinirse tablo kalkar). */
export function deleteTableCols(ctl: Controller, t: TableDef): void {
  const { sheet, wb } = ctl;
  const span = tableSpan(ctl, t);
  const count = span.c2 - span.c1 + 1;
  if (count <= 0) return;
  const band: Rect = { r1: t.ref.r1, c1: span.c1, r2: t.ref.r2, c2: 16383 };
  if (blocked(ctl, t, band)) return refuse(ctl);
  ctl.ops.snapshotOp("Tablo Sütunlarını Sil", () => {
    const cur = sheet.tables.find((x) => x.id === t.id)!;
    ctl.ops.shiftCellsRaw(sheet, { r1: cur.ref.r1, c1: span.c1, r2: cur.ref.r2, c2: span.c2 }, "left");
    if (count >= cur.columns.length) {
      sheet.tables = sheet.tables.filter((x) => x.id !== cur.id);
      return;
    }
    const idx = span.c1 - cur.ref.c1;
    cur.columns = [...cur.columns.slice(0, idx), ...cur.columns.slice(idx + count)];
    cur.ref = { ...cur.ref, c2: cur.ref.c2 - count };
    if (cur.filter) cur.filter = { range: { ...cur.ref, r2: cur.totalRow ? cur.ref.r2 - 1 : cur.ref.r2 }, columns: {} };
  });
  wb.syncTables();
  ctl.selectCell(ctl.sel.active.row, Math.min(span.c1, t.ref.c2 - count));
}

/** Seç > Tablo Sütunu Verileri / Tüm Tablo Sütunu / Tablo Satırı. */
export function selectTablePart(ctl: Controller, t: TableDef, what: "colData" | "col" | "row"): void {
  const span = tableSpan(ctl, t);
  const [d1, d2] = dataRows(t);
  if (what === "colData") ctl.selectRange({ r1: d1, c1: span.c1, r2: Math.max(d1, d2), c2: span.c2 });
  else if (what === "col") ctl.selectRange({ r1: t.ref.r1, c1: span.c1, r2: t.ref.r2, c2: span.c2 });
  else ctl.selectRange({ r1: span.r1, c1: t.ref.c1, r2: span.r2, c2: t.ref.c2 });
}
