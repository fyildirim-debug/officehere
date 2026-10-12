/**
 * Dilimleyiciler: tablo sütununa ya da özet tablo alanına bağlı düğmeli
 * filtre. Öğeye tıklamak yalnız o öğeyi, Ctrl ile tıklamak (ya da Çoklu
 * Seçim açıkken) seçime ekleyip çıkarmayı sağlar; Filtreyi Temizle tümünü seçer.
 * Tablo kaynağında tablonun otomatik filtresi, özet tabloda alanın gizli
 * öğeleri değişir (dosyada filtre durumu korunur).
 */
import type { Controller } from "../controller";
import { displayText, formatScalar } from "../model/display";
import type { SlicerDef, TableDef } from "../model/types";
import type { SheetModel } from "../model/workbook";
import { recomputeFilters } from "./data";
import { fieldItems, itemLabel } from "../pivot/compute";
import { applyPivot, readSource } from "../pivot/pivotOps";

export const BLANK_LABEL = "(boş)";

export interface SlicerItem {
  label: string;
  /** Özet tablo öğe anahtarı (tabloda görünen metin). */
  key: string;
  selected: boolean;
}

function tableOf(ctl: Controller, s: SlicerDef): { sheet: SheetModel; t: TableDef } | null {
  if (s.source.kind !== "table") return null;
  const id = s.source.tableId;
  for (const sheet of ctl.wb.sheets) {
    const t = sheet.tables.find((x) => x.id === id);
    if (t) return { sheet, t };
  }
  return null;
}

function pivotOf(ctl: Controller, s: SlicerDef) {
  if (s.source.kind !== "pivot") return null;
  const id = s.source.pivotId;
  for (const sheet of ctl.wb.sheets) {
    const p = sheet.pivots.find((x) => x.id === id);
    if (p) return { sheet, p };
  }
  return null;
}

/** Dilimleyicinin öğeleri ve seçim durumu. */
export function slicerItems(ctl: Controller, s: SlicerDef): SlicerItem[] {
  const wb = ctl.wb;
  const all = s.selected.length === 0;
  const has = new Set(s.selected);
  if (s.source.kind === "table") {
    const tt = tableOf(ctl, s);
    if (!tt) return [];
    const { sheet, t } = tt;
    const col = t.ref.c1 + t.columns.indexOf(s.source.column);
    if (col < t.ref.c1) return [];
    const d1 = t.ref.r1 + (t.headerRow ? 1 : 0);
    const d2 = t.ref.r2 - (t.totalRow ? 1 : 0);
    const seen = new Map<string, { label: string; v: unknown }>();
    let blank = false;
    for (let r = d1; r <= d2; r++) {
      const text = displayText(wb, sheet, r, col);
      if (text === "") blank = true;
      else if (!seen.has(text.toLocaleUpperCase("tr-TR"))) seen.set(text.toLocaleUpperCase("tr-TR"), { label: text, v: wb.getValue(sheet, r, col) });
    }
    const list = [...seen.values()].sort((a, b) => (typeof a.v === "number" && typeof b.v === "number" ? a.v - b.v : a.label.localeCompare(b.label, "tr")));
    const out = list.map((x) => ({ label: x.label, key: x.label, selected: all || has.has(x.label) }));
    if (blank) out.push({ label: BLANK_LABEL, key: BLANK_LABEL, selected: all || has.has(BLANK_LABEL) });
    return out;
  }
  const pp = pivotOf(ctl, s);
  if (!pp) return [];
  const src = readSource(wb, pp.p);
  if (typeof src === "string") return [];
  const field = s.source.field;
  const hidden = new Set(pp.p.hidden?.[field] ?? []);
  return fieldItems(src, field).map((it) => ({ label: formatScalar(itemLabel(it.value)), key: it.key, selected: !hidden.has(it.key) }));
}

/** Seçimi uygular (seçili öğe anahtarları; boş liste ya da hepsi = filtre yok). */
export function applySlicer(ctl: Controller, s: SlicerDef, keys: string[]): void {
  const items = slicerItems(ctl, s);
  const all = !keys.length || items.every((i) => keys.includes(i.key));
  const owner = ctl.wb.sheets.find((sh) => sh.slicers.some((x) => x.id === s.id));
  if (!owner) return;
  if (s.source.kind === "table") {
    const tt = tableOf(ctl, s);
    if (!tt) return;
    const colIdx = tt.t.columns.indexOf(s.source.column);
    ctl.ops.snapshotOp("Dilimleyici", () => {
      const sl = owner.slicers.find((x) => x.id === s.id)!;
      sl.selected = all ? [] : [...keys];
      const t = tt.sheet.tables.find((x) => x.id === tt.t.id)!;
      if (!t.filter) t.filter = { range: { ...t.ref, r2: t.totalRow ? t.ref.r2 - 1 : t.ref.r2 }, columns: {} };
      if (all) delete t.filter.columns[colIdx];
      else t.filter.columns[colIdx] = { kind: "values", values: keys.filter((k) => k !== BLANK_LABEL), blanks: keys.includes(BLANK_LABEL) };
      recomputeFilters(ctl, tt.sheet);
    });
    return;
  }
  const pp = pivotOf(ctl, s);
  if (!pp) return;
  const field = s.source.field;
  const hidden = all ? [] : items.filter((i) => !keys.includes(i.key)).map((i) => i.key);
  applyPivot(
    ctl,
    pp.sheet,
    pp.p.id,
    (d) => {
      const h = { ...(d.hidden ?? {}) };
      if (hidden.length) h[field] = hidden;
      else delete h[field];
      d.hidden = h;
    },
    "Dilimleyici",
  );
}

/** Öğeye tıklama: tek seçim ya da (çoklu) seçime ekle/çıkar. */
export function clickSlicerItem(ctl: Controller, s: SlicerDef, key: string, multi: boolean): void {
  const items = slicerItems(ctl, s);
  const selected = items.filter((i) => i.selected).map((i) => i.key);
  const allSelected = selected.length === items.length;
  let next: string[];
  if (!multi) next = allSelected || selected.length !== 1 || selected[0] !== key ? [key] : [];
  else if (allSelected) next = items.map((i) => i.key).filter((k) => k !== key);
  else next = selected.includes(key) ? selected.filter((k) => k !== key) : [...selected, key];
  applySlicer(ctl, s, next);
}

/** Dilimleyici ekler (her alan için bir tane); oluşturulanların kimlikleri. */
export function insertSlicers(ctl: Controller, source: { kind: "table"; tableId: string } | { kind: "pivot"; pivotId: string }, fields: { name: string; field?: number }[]): string[] {
  const sheet = ctl.sheet;
  const ids: string[] = [];
  // Kaynağın sağına yerleştirilir (veriyi örtmez)
  const t = source.kind === "table" ? sheet.tables.find((x) => x.id === source.tableId) : undefined;
  const p = source.kind === "pivot" ? sheet.pivots.find((x) => x.id === source.pivotId) : undefined;
  const area = t?.ref ?? p?.output;
  const row = area ? area.r1 : ctl.sel.active.row;
  const col = area ? area.c2 - 1 : ctl.sel.active.col;
  ctl.ops.sheetOp(sheet, "Dilimleyici Ekle", () => {
    fields.forEach((f, i) => {
      const id = ctl.wb.newId("slicer");
      const taken = new Set(ctl.wb.sheets.flatMap((s) => s.slicers.map((x) => x.name)));
      let name = `Dilimleyici_${f.name.replace(/[^\p{L}\p{N}_]/gu, "_")}`;
      let n = 1;
      while (taken.has(name)) name = `Dilimleyici_${f.name.replace(/[^\p{L}\p{N}_]/gu, "_")}${++n}`;
      sheet.slicers.push({
        id,
        name,
        caption: f.name,
        source: source.kind === "table" ? { kind: "table", tableId: source.tableId, column: f.name } : { kind: "pivot", pivotId: source.pivotId, field: f.field ?? 0 },
        selected: [],
        anchor: { row: row + i, col: col + 2 + i, dx: 8, dy: 8 },
        width: 180,
        height: 220,
        columns: 1,
        style: "light1",
      });
      ids.push(id);
    });
  });
  ctl.selSlicer = ids[ids.length - 1] ?? null;
  ctl.emit();
  return ids;
}

export function updateSlicer(ctl: Controller, id: string, patch: Partial<SlicerDef>, label = "Dilimleyici"): void {
  const sheet = ctl.wb.sheets.find((s) => s.slicers.some((x) => x.id === id));
  if (!sheet) return;
  ctl.ops.sheetOp(sheet, label, () => {
    sheet.slicers = sheet.slicers.map((x) => (x.id === id ? { ...x, ...patch } : x));
  });
}

/** Dilimleyiciyi siler; bağlı filtre temizlenir (Excel gibi). */
export function deleteSlicer(ctl: Controller, id: string): void {
  const sheet = ctl.wb.sheets.find((s) => s.slicers.some((x) => x.id === id));
  const s = sheet?.slicers.find((x) => x.id === id);
  if (!sheet || !s) return;
  const mark = ctl.ops.history.mark();
  if (s.selected.length || (s.source.kind === "pivot" && pivotOf(ctl, s)?.p.hidden?.[s.source.field]?.length)) applySlicer(ctl, s, []);
  ctl.ops.sheetOp(sheet, "Dilimleyiciyi Sil", () => (sheet.slicers = sheet.slicers.filter((x) => x.id !== id)));
  ctl.ops.history.squash(mark, "Dilimleyiciyi Sil");
  ctl.selSlicer = null;
  ctl.emit();
}

/** Dilimleyici eklenebilecek kaynak (etkin hücredeki tablo ya da özet tablo). */
export function slicerSourceAt(ctl: Controller): { kind: "table"; table: TableDef } | { kind: "pivot"; pivotId: string; fields: string[] } | null {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const t = sheet.tableAt(row, col);
  if (t) return { kind: "table", table: t };
  const p = sheet.pivots.find((x) => x.output && row >= x.output.r1 && row <= x.output.r2 && col >= x.output.c1 && col <= x.output.c2);
  if (p) {
    const src = readSource(ctl.wb, p);
    if (typeof src === "string") return null;
    return { kind: "pivot", pivotId: p.id, fields: src.headers.map((h) => String(h ?? "")) };
  }
  return null;
}
