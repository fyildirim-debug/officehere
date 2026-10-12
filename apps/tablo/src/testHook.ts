/**
 * Uçtan uca testler için salt okunur durum: `window.__ohTablo`. Canvas
 * ızgarada hücre konumları ve değerleri DOM'dan okunamadığı için testler
 * bu küçük arayüzü kullanır. Durumu değiştiren hiçbir işlev içermez.
 */
import { colToName, isError, nameToCol } from "@officehere/formula";
import type { Controller } from "./controller";

export interface TabloTestApi {
  /** Hücrenin değeri, formülü (İngilizce, kanonik), yerel formülü ve görünen metni. */
  cell(ref: string, sheet?: string): { value: string | number | boolean | null; formula: string | null; edit: string; numFmt: string | null; bold: boolean; fill: string | null } | null;
  /** Hücrenin ekran konumu (fareyle tıklamak için). */
  cellRect(ref: string): { x: number; y: number; w: number; h: number } | null;
  /** Etkin sayfa adı, sayfa adları, seçili aralık. */
  activeSheet(): string;
  sheetNames(): string[];
  selection(): string;
  chartCount(sheet?: string): number;
  pivotCount(sheet?: string): number;
  /** Mini grafik grupları: tür ve hücreleri ("H2"). */
  sparklines(sheet?: string): { type: string; cells: string[] }[];
  isRowHidden(row: number, sheet?: string): boolean;
  undoCount(): number;
  version(): number;
}

function parseRef(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref.trim());
  if (!m) return null;
  return { row: Number(m[2]) - 1, col: nameToCol(m[1].toUpperCase()) };
}

export function installTestHook(ctl: Controller): void {
  const sheetOf = (name?: string) => (name ? ctl.wb.sheetByName(name) : ctl.sheet);
  const api: TabloTestApi = {
    cell(ref, sheetName) {
      const p = parseRef(ref);
      const sh = sheetOf(sheetName);
      if (!p || !sh) return null;
      const v = ctl.wb.getValue(sh, p.row, p.col);
      const st = ctl.wb.styleAt(sh, p.row, p.col);
      return {
        value: isError(v) ? v.code : (v as string | number | boolean | null),
        formula: ctl.wb.getFormula(sh, p.row, p.col) ?? null,
        edit: ctl.wb.editText(sh, p.row, p.col),
        numFmt: st.numFmt ?? null,
        bold: !!st.bold,
        fill: st.fill ?? null,
      };
    },
    cellRect(ref) {
      const p = parseRef(ref);
      return p ? ctl.cellClientRect(p.row, p.col) : null;
    },
    activeSheet: () => ctl.sheet.name,
    sheetNames: () => ctl.wb.sheets.map((s) => s.name),
    selection: () =>
      ctl.sel.ranges
        .map((r) => {
          const a = `${colToName(r.c1)}${r.r1 + 1}`;
          return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${colToName(r.c2)}${r.r2 + 1}`;
        })
        .join(","),
    chartCount: (name) => sheetOf(name)?.charts.length ?? 0,
    pivotCount: (name) => sheetOf(name)?.pivots.length ?? 0,
    sparklines: (name) => (sheetOf(name)?.sparklines ?? []).map((g) => ({ type: g.type, cells: g.items.map((i) => `${colToName(i.col)}${i.row + 1}`) })),
    isRowHidden: (row, name) => !!sheetOf(name)?.rows.isHidden(row),
    undoCount: () => ctl.ops.history.undoCount,
    version: () => ctl.version,
  };
  (window as unknown as { __ohTablo?: TabloTestApi }).__ohTablo = api;
}
