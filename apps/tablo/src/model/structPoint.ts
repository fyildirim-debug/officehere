/**
 * Formül yazarken tablo hücrelerine tıklama/sürükleme: Excel gibi A1 yerine
 * yapısal başvuru üretilir ([@Tutar], Tablo1[Tutar], Tablo1[[#Başlıklar];[Tutar]]).
 */
import { renderStructRef, type StructItem, type StructRef } from "@officehere/formula";
import type { Rect } from "./types";
import type { SheetModel } from "./workbook";

/** Başvuru metni; tablo başvurusuna uymayan seçimde null (A1 kullanılır). */
export function structuredPointRef(sheet: SheetModel, r: Rect, edit: { sheetId: number; row: number; col: number }): string | null {
  const t = sheet.tables.find((x) => r.r1 >= x.ref.r1 && r.r2 <= x.ref.r2 && r.c1 >= x.ref.c1 && r.c2 <= x.ref.c2);
  if (!t) return null;
  const d1 = t.ref.r1 + (t.headerRow ? 1 : 0);
  const d2 = t.ref.r2 - (t.totalRow ? 1 : 0);
  const sameSheet = edit.sheetId === sheet.id;
  const inside = sameSheet && edit.row >= t.ref.r1 && edit.row <= t.ref.r2 && edit.col >= t.ref.c1 && edit.col <= t.ref.c2;
  let items: StructItem[];
  if (r.r1 === r.r2 && r.r1 >= d1 && r.r1 <= d2) {
    // Tek veri satırı: yalnız formülle aynı satırdaysa "bu satır"
    if (!sameSheet || edit.row !== r.r1) return null;
    items = ["#This Row"];
  } else if (r.r1 === d1 && r.r2 === d2) items = [];
  else if (r.r1 === t.ref.r1 && r.r2 === t.ref.r2) items = ["#All"];
  else if (t.headerRow && r.r1 === t.ref.r1 && r.r2 === t.ref.r1) items = ["#Headers"];
  else if (t.totalRow && r.r1 === t.ref.r2 && r.r2 === t.ref.r2) items = ["#Totals"];
  else if (t.headerRow && r.r1 === t.ref.r1 && r.r2 === d2) items = ["#Headers", "#Data"];
  else if (t.totalRow && r.r1 === d1 && r.r2 === t.ref.r2) items = ["#Data", "#Totals"];
  else return null;
  const whole = r.c1 === t.ref.c1 && r.c2 === t.ref.c2 && t.columns.length > 1;
  const ref: StructRef = { table: inside ? undefined : t.name, items };
  if (!whole) {
    ref.col1 = t.columns[r.c1 - t.ref.c1];
    if (r.c2 > r.c1) ref.col2 = t.columns[r.c2 - t.ref.c1];
  } else if (!items.length) ref.items = ["#Data"];
  if (ref.col1 === undefined && whole && items.length === 1 && items[0] === "#This Row") ref.items = ["#This Row"];
  return renderStructRef(ref, "tr");
}
