/** PivotTable alan listesi işlemleri (tanım üzerinde saf dönüşümler). */
import type { PivotDef } from "../model/types";
import { isNumericField, VALUES_FIELD, type PivotSource } from "./compute";

export type PivotArea = "filters" | "rows" | "cols" | "values";

export const AREA_LABEL: Record<PivotArea, string> = {
  filters: "Filtreler",
  cols: "Sütunlar",
  rows: "Satırlar",
  values: "Değerler",
};

/** Alan raporda kullanılıyor mu? */
export function isFieldUsed(def: PivotDef, field: number): boolean {
  return def.rows.includes(field) || def.cols.includes(field) || def.values.some((v) => v.field === field) || def.filters.some((f) => f.field === field);
}

/** Bir alanın satır/sütun/filtre alanlarından çıkarılması (değerler hariç). */
function detach(def: PivotDef, field: number): void {
  def.rows = def.rows.filter((f) => f !== field);
  def.cols = def.cols.filter((f) => f !== field);
  def.filters = def.filters.filter((f) => f.field !== field);
}

/** Onay kutusu: kullanılıyorsa kaldır; değilse sayısal → Değerler, diğer → Satırlar. */
export function toggleField(def: PivotDef, src: PivotSource, field: number): void {
  if (isFieldUsed(def, field)) {
    detach(def, field);
    def.values = def.values.filter((v) => v.field !== field);
    if (def.hidden) delete def.hidden[field];
    return;
  }
  if (isNumericField(src, field)) def.values.push({ field, fn: "sum" });
  else def.rows.push(field);
}

/** Alanı bir bölgeye taşır/ekler. fromValues: Değerler'deki dizin (değer alanını taşırken). */
export function placeField(def: PivotDef, src: PivotSource, field: number, to: PivotArea, index?: number, fromValues?: number): void {
  if (field === VALUES_FIELD) {
    if (to === "rows") def.valuesOnRows = true;
    else if (to === "cols") def.valuesOnRows = false;
    return;
  }
  if (to === "values") {
    if (fromValues !== undefined) {
      // Değerler içinde yeniden sıralama
      const [vf] = def.values.splice(fromValues, 1);
      const i = index === undefined ? def.values.length : Math.min(index, def.values.length);
      def.values.splice(i, 0, vf);
      return;
    }
    const vf = { field, fn: isNumericField(src, field) ? ("sum" as const) : ("count" as const) };
    const i = index === undefined ? def.values.length : Math.min(index, def.values.length);
    def.values.splice(i, 0, vf);
    return;
  }
  if (fromValues !== undefined) def.values.splice(fromValues, 1);
  const prevFilter = def.filters.find((f) => f.field === field);
  detach(def, field);
  if (to === "filters") {
    const i = index === undefined ? def.filters.length : Math.min(index, def.filters.length);
    def.filters.splice(i, 0, { field, selected: prevFilter?.selected ?? null });
    // Rapor filtresine taşınan alanın satır/sütun gizli öğeleri seçim olarak korunmaz
    if (def.hidden) delete def.hidden[field];
    return;
  }
  const list = to === "rows" ? def.rows : def.cols;
  const i = index === undefined ? list.length : Math.min(index, list.length);
  list.splice(i, 0, field);
}

/** Bölgedeki öğeyi kaldırır. */
export function removeAt(def: PivotDef, area: PivotArea, index: number): void {
  if (area === "values") def.values.splice(index, 1);
  else if (area === "filters") def.filters.splice(index, 1);
  else (area === "rows" ? def.rows : def.cols).splice(index, 1);
}

/** Bölge içinde konum değiştirir. */
export function moveWithin(def: PivotDef, area: PivotArea, index: number, to: number): void {
  const list: unknown[] = area === "values" ? def.values : area === "filters" ? def.filters : area === "rows" ? def.rows : def.cols;
  if (index < 0 || index >= list.length) return;
  const [it] = list.splice(index, 1);
  list.splice(Math.max(0, Math.min(to, list.length)), 0, it);
}

/** Etkin hücrenin altındaki alan (Etkin Alan) — satır etiketi sütununa göre yaklaşık. */
export function areaItems(def: PivotDef): Record<PivotArea, number[]> {
  return { filters: def.filters.map((f) => f.field), rows: def.rows, cols: def.cols, values: def.values.map((v) => v.field) };
}
