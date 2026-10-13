/**
 * Birleştir (Veri > Veri Araçları): birden çok aralığın (başka sayfalardakiler de)
 * değerlerini seçilen işlevle tek bir tabloda toplar. Etiket kullanılırsa üst satır
 * ve/veya sol sütundaki etiketler eşleştirilir (büyük/küçük harf duyarsız), aksi
 * hâlde aralıklar konumlarına göre üst üste konur. Sonuç değer olarak yazılır.
 */
import { Errors, isError, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import type { Rect } from "../model/types";
import type { SheetModel } from "../model/workbook";

export type ConsolidateFn = "SUM" | "COUNT" | "AVERAGE" | "MAX" | "MIN" | "PRODUCT" | "COUNTNUMS" | "STDEV" | "STDEVP" | "VAR" | "VARP";

export const CONSOLIDATE_FNS: [ConsolidateFn, string][] = [
  ["SUM", "Topla"],
  ["COUNT", "Say"],
  ["AVERAGE", "Ortalama"],
  ["MAX", "Maks"],
  ["MIN", "Min"],
  ["PRODUCT", "Çarpım"],
  ["COUNTNUMS", "Sayıları Say"],
  ["STDEV", "StdSapma"],
  ["STDEVP", "StdSapmaS"],
  ["VAR", "Var"],
  ["VARP", "VarS"],
];

/** "Sayfa2!$A$1:$C$5" ya da "A1:C5" → sayfa ve aralık; geçersizse null. */
export function parseSourceRef(ctl: Controller, text: string): { sheet: SheetModel; rect: Rect } | null {
  const t = text.trim().replace(/^=/, "");
  if (!t) return null;
  let sheet: SheetModel | undefined = ctl.sheet;
  let body = t;
  const bang = t.lastIndexOf("!");
  if (bang > 0) {
    let name = t.slice(0, bang);
    if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
    sheet = ctl.wb.sheetByName(name);
    body = t.slice(bang + 1);
  }
  const rect = parseRangeLoose(body.replace(/\$/g, ""));
  return sheet && rect ? { sheet, rect } : null;
}

function aggregate(fn: ConsolidateFn, vals: Scalar[]): Scalar {
  const nums = vals.filter((v): v is number => typeof v === "number");
  const n = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  const variance = (sample: boolean) => {
    if (n === 0 || (sample && n < 2)) return Errors.DIV0;
    const m = sum / n;
    return nums.reduce((a, b) => a + (b - m) * (b - m), 0) / (sample ? n - 1 : n);
  };
  switch (fn) {
    case "SUM":
      return sum;
    case "COUNT":
      return vals.filter((v) => v !== null && v !== "").length;
    case "COUNTNUMS":
      return n;
    case "AVERAGE":
      return n ? sum / n : Errors.DIV0;
    case "MAX":
      return n ? Math.max(...nums) : 0;
    case "MIN":
      return n ? Math.min(...nums) : 0;
    case "PRODUCT":
      return n ? nums.reduce((a, b) => a * b, 1) : 0;
    case "STDEV":
    case "STDEVP": {
      const v = variance(fn === "STDEV");
      return isError(v) ? v : Math.sqrt(v);
    }
    case "VAR":
    case "VARP":
      return variance(fn === "VAR");
  }
}

const key = (v: Scalar) => (v === null ? "" : String(v).trim().toLocaleUpperCase("tr-TR"));

/**
 * Birleştirmeyi yapar ve hedefe yazar. Hata metni döner (yoksa null).
 * dest: sonucun sol üst hücresi (etkin sayfada).
 */
export function consolidate(ctl: Controller, refs: string[], fn: ConsolidateFn, topRow: boolean, leftCol: boolean, dest: { row: number; col: number }): string | null {
  if (!refs.length) return "Birleştirilecek en az bir başvuru ekleyin.";
  const sources: { sheet: SheetModel; rect: Rect }[] = [];
  for (const r of refs) {
    const s = parseSourceRef(ctl, r);
    if (!s) return `Başvuru geçerli değil: ${r}`;
    sources.push(s);
  }
  const { wb } = ctl;
  const rowKeys: string[] = [];
  const colKeys: string[] = [];
  const rowLabel = new Map<string, Scalar>();
  const colLabel = new Map<string, Scalar>();
  const bucket = new Map<string, Scalar[]>();
  for (const { sheet, rect } of sources) {
    const dr0 = topRow ? 1 : 0;
    const dc0 = leftCol ? 1 : 0;
    for (let i = rect.r1 + dr0; i <= rect.r2; i++) {
      const ri = i - rect.r1 - dr0;
      let rk = String(ri);
      if (leftCol) {
        const lab = wb.getValue(sheet, i, rect.c1);
        if (lab === null || lab === "") continue;
        rk = key(lab);
        if (!rowLabel.has(rk)) rowLabel.set(rk, lab);
      }
      if (!rowKeys.includes(rk)) rowKeys.push(rk);
      for (let j = rect.c1 + dc0; j <= rect.c2; j++) {
        const cj = j - rect.c1 - dc0;
        let ck = String(cj);
        if (topRow) {
          const lab = wb.getValue(sheet, rect.r1, j);
          if (lab === null || lab === "") continue;
          ck = key(lab);
          if (!colLabel.has(ck)) colLabel.set(ck, lab);
        }
        if (!colKeys.includes(ck)) colKeys.push(ck);
        const k = rk + "\u0001" + ck;
        let b = bucket.get(k);
        if (!b) bucket.set(k, (b = []));
        b.push(wb.getValue(sheet, i, j));
      }
    }
  }
  if (!rowKeys.length || !colKeys.length) return "Birleştirilecek veri bulunamadı.";
  // Konuma göre birleştirmede satır/sütun sırası sayısaldır
  if (!leftCol) rowKeys.sort((a, b) => Number(a) - Number(b));
  if (!topRow) colKeys.sort((a, b) => Number(a) - Number(b));
  const entries: { row: number; col: number; input: Scalar }[] = [];
  const r0 = dest.row + (topRow ? 1 : 0);
  const c0 = dest.col + (leftCol ? 1 : 0);
  if (topRow) colKeys.forEach((ck, j) => entries.push({ row: dest.row, col: c0 + j, input: colLabel.get(ck) ?? null }));
  rowKeys.forEach((rk, i) => {
    if (leftCol) entries.push({ row: r0 + i, col: dest.col, input: rowLabel.get(rk) ?? null });
    colKeys.forEach((ck, j) => {
      const vals = bucket.get(rk + "\u0001" + ck);
      entries.push({ row: r0 + i, col: c0 + j, input: vals && vals.some((v) => v !== null) ? aggregate(fn, vals) : null });
    });
  });
  if (!ctl.guardSelection()) return null;
  ctl.ops.setInputs(ctl.sheet, entries, "Birleştir");
  const r2 = r0 + rowKeys.length - 1;
  const c2 = c0 + colKeys.length - 1;
  ctl.selectRange({ r1: dest.row, c1: dest.col, r2, c2 });
  return null;
}
