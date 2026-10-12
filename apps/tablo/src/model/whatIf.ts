/**
 * Durum Çözümlemesi hesapları: senaryo değerleriyle sonuç hücrelerinin
 * hesaplanması ve veri tablosu (TABLE) sonuçları. Giriş hücrelerine değerler
 * geçici olarak yazılır, hesaplanır ve özgün içerik geri yüklenir.
 */
import type { Scalar } from "@officehere/formula";
import type { CellPos, DataTableDef, Rect, Scenario } from "./types";
import type { CellInput, SheetModel, WorkbookModel } from "./workbook";

type Cell = { row: number; col: number };

/** Hücrelere geçici değerler yazıp fn'i çalıştırır; sonunda özgün içerikleri geri yükler. */
export function withTemporaryInputs<T>(wb: WorkbookModel, sheet: SheetModel, cells: Cell[], fn: (set: (values: Scalar[]) => void) => T): T {
  const originals: CellInput[] = cells.map((c) => wb.getInput(sheet, c.row, c.col));
  const set = (values: Scalar[]) => {
    cells.forEach((c, i) => wb.engine.setValue(sheet.id, c.row, c.col, values[i] ?? null));
    wb.engine.recalc();
  };
  try {
    return fn(set);
  } finally {
    cells.forEach((c, i) => wb.setInputRaw(sheet, c.row, c.col, originals[i]));
    wb.engine.recalc();
  }
}

/** Her senaryo için sonuç hücrelerinin değerleri (senaryo sırasıyla). */
export function scenarioResults(wb: WorkbookModel, sheet: SheetModel, scenarios: Scenario[], results: Cell[]): Scalar[][] {
  const cells: Cell[] = [];
  const key = (c: Cell) => `${c.row}:${c.col}`;
  const seen = new Set<string>();
  for (const s of scenarios)
    for (const c of s.cells)
      if (!seen.has(key(c))) {
        seen.add(key(c));
        cells.push({ row: c.row, col: c.col });
      }
  return withTemporaryInputs(wb, sheet, cells, (set) =>
    scenarios.map((s) => {
      // Senaryoda olmayan değişen hücreler mevcut değerinde kalır
      const current = cells.map((c) => wb.getValue(sheet, c.row, c.col));
      set(cells.map((c, i) => s.cells.find((x) => x.row === c.row && x.col === c.col)?.value ?? current[i]));
      return results.map((r) => wb.getValue(sheet, r.row, r.col));
    }),
  );
}

/** Seçimden veri tablosu tanımı: sonuç aralığı seçimin sol üst satır/sütunu dışıdır. */
export function dataTableFromSelection(sel: Rect, rowInput?: CellPos, colInput?: CellPos): DataTableDef | string {
  if (sel.r2 <= sel.r1 && sel.c2 <= sel.c1) return "Veri tablosu için en az iki satır ya da iki sütunluk bir aralık seçin.";
  if (!rowInput && !colInput) return "Giriş hücresi başvurusu geçerli değil.";
  const ref = { r1: sel.r1 + 1, c1: sel.c1 + 1, r2: sel.r2, c2: sel.c2 };
  if (ref.r2 < ref.r1 || ref.c2 < ref.c1) return "Seçili aralık veri tablosu için çok küçük.";
  const inside = (p?: CellPos) => !!p && p.row >= sel.r1 && p.row <= sel.r2 && p.col >= sel.c1 && p.col <= sel.c2;
  if (inside(rowInput) || inside(colInput)) return "Giriş hücresi veri tablosunun dışında olmalıdır.";
  return { id: "", ref, ...(rowInput ? { rowInput } : {}), ...(colInput ? { colInput } : {}) };
}

/** Veri tablosunun sonuç değerleri (ref boyutunda, satır satır). */
export function computeDataTable(wb: WorkbookModel, sheet: SheetModel, dt: DataTableDef): Scalar[][] {
  const { ref } = dt;
  const top = ref.r1 - 1;
  const left = ref.c1 - 1;
  const inputs: Cell[] = [];
  if (dt.rowInput) inputs.push(dt.rowInput);
  if (dt.colInput) inputs.push(dt.colInput);
  return withTemporaryInputs(wb, sheet, inputs, (set) => {
    const out: Scalar[][] = [];
    for (let r = ref.r1; r <= ref.r2; r++) {
      const line: Scalar[] = [];
      for (let c = ref.c1; c <= ref.c2; c++) {
        const vals: Scalar[] = [];
        if (dt.rowInput) vals.push(wb.getValue(sheet, top, c));
        if (dt.colInput) vals.push(wb.getValue(sheet, r, left));
        set(vals);
        // İki değişkenli: formül sol üst köşede; tek değişkenli sütun: formüller üst satırda; satır: sol sütunda
        const src = dt.rowInput && dt.colInput ? { row: top, col: left } : dt.colInput ? { row: top, col: c } : { row: r, col: left };
        const v = wb.getValue(sheet, src.row, src.col);
        line.push(v);
      }
      out.push(line);
    }
    return out;
  });
}

/** Tüm veri tablolarını yeniden hesaplayıp sonuçları hücrelere yazar; değişen varsa true. */
export function refreshDataTables(wb: WorkbookModel): boolean {
  let changed = false;
  for (const sheet of wb.sheets)
    for (const dt of sheet.dataTables) {
      const vals = computeDataTable(wb, sheet, dt);
      vals.forEach((line, i) =>
        line.forEach((v, j) => {
          const r = dt.ref.r1 + i;
          const c = dt.ref.c1 + j;
          if (wb.getInput(sheet, r, c) !== v) {
            wb.engine.setValue(sheet.id, r, c, v);
            changed = true;
          }
        }),
      );
    }
  if (changed) wb.engine.recalc();
  return changed;
}

/** Hücre bir veri tablosunun sonuç aralığında mı? */
export function dataTableAt(sheet: SheetModel, row: number, col: number): DataTableDef | undefined {
  return sheet.dataTables.find((d) => row >= d.ref.r1 && row <= d.ref.r2 && col >= d.ref.c1 && col <= d.ref.c2);
}
