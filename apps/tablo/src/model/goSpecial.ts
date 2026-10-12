/**
 * Özel Git (Git > Özel): boşluklar, sabitler, formüller (alt türleriyle),
 * notlar, geçerli bölge/dizi, son hücre, yalnız görünür hücreler, satır ve
 * sütun farkları, etkileyen/etkilenen hücreler, koşullu biçimler, veri
 * doğrulama. Seçim tek hücreyse kullanılan alan, değilse seçim taranır.
 */
import { cellKey, dependentsOf, isError, keyCol, keyRow, precedentsOf, type Scalar } from "@officehere/formula";
import { compressCells } from "../io/xlsxRules";
import { contains, intersects } from "./rect";
import { currentRegion } from "./selection";
import type { CellPos, Rect } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

export type GoSpecialKind =
  | "notes"
  | "constants"
  | "formulas"
  | "blanks"
  | "currentRegion"
  | "currentArray"
  | "objects"
  | "rowDiffs"
  | "colDiffs"
  | "precedents"
  | "dependents"
  | "lastCell"
  | "visible"
  | "condFormats"
  | "validation";

export interface GoSpecialOptions {
  /** Sabitler/formüller alt türleri. */
  numbers?: boolean;
  text?: boolean;
  logicals?: boolean;
  errors?: boolean;
  /** Etkileyen/etkilenen: tüm düzeyler. */
  allLevels?: boolean;
  /** Koşullu biçim / veri doğrulama: yalnız etkin hücreyle aynı olanlar. */
  same?: boolean;
}

export interface GoSpecialResult {
  ranges: Rect[];
  active: CellPos;
  /** Nesneler seçildiyse ilk nesne. */
  object?: { kind: "chart" | "image" | "shape"; id: string };
}

function typeMatch(v: Scalar, o: GoSpecialOptions): boolean {
  if (typeof v === "number") return o.numbers !== false;
  if (typeof v === "string") return o.text !== false;
  if (typeof v === "boolean") return o.logicals !== false;
  if (isError(v)) return o.errors !== false;
  return false;
}

/** Aranacak alan: tek hücre seçiliyse kullanılan alan. */
function scope(wb: WorkbookModel, sheet: SheetModel, sel: Rect[]): Rect[] {
  const single = sel.length === 1 && sel[0].r1 === sel[0].r2 && sel[0].c1 === sel[0].c2;
  if (!single) {
    // Tam satır/sütun seçimleri kullanılan alanla sınırlanır
    const used = wb.usedRange(sheet);
    if (!used) return [];
    return sel.map((r) => ({ r1: r.r1, c1: r.c1, r2: Math.min(r.r2, used.r2), c2: Math.min(r.c2, used.c2) })).filter((r) => r.r1 <= r.r2 && r.c1 <= r.c2);
  }
  const used = wb.usedRange(sheet);
  return used ? [{ r1: 0, c1: 0, r2: used.r2, c2: used.c2 }] : [];
}

export function goSpecial(wb: WorkbookModel, sheet: SheetModel, sel: Rect[], active: CellPos, kind: GoSpecialKind, o: GoSpecialOptions = {}): GoSpecialResult | null {
  const cells: [number, number][] = [];
  const area = scope(wb, sheet, sel);
  const inArea = (r: number, c: number) => area.some((a) => contains(a, r, c));
  const result = (rs: Rect[]): GoSpecialResult | null => {
    if (!rs.length) return null;
    const a = rs.some((r) => contains(r, active.row, active.col)) ? active : { row: rs[0].r1, col: rs[0].c1 };
    return { ranges: rs, active: a };
  };
  const fromCells = () => result(compressCells(cells));
  switch (kind) {
    case "notes":
      for (const k of sheet.notes.keys()) if (inArea(keyRow(k), keyCol(k))) cells.push([keyRow(k), keyCol(k)]);
      return fromCells();
    case "constants":
    case "formulas":
      wb.engine.forEachCell(sheet.id, (r, c, rec) => {
        if (!inArea(r, c)) return;
        const isF = !!rec.formula;
        if ((kind === "formulas") !== isF) return;
        if (typeMatch(isF ? wb.getValue(sheet, r, c) : rec.value, o)) cells.push([r, c]);
      });
      return fromCells();
    case "blanks":
      for (const a of area)
        for (let r = a.r1; r <= Math.min(a.r2, a.r1 + 100000); r++) for (let c = a.c1; c <= Math.min(a.c2, a.c1 + 500); c++) if (wb.isEmpty(sheet, r, c) && !sheet.mergeAt(r, c)) cells.push([r, c]);
      return fromCells();
    case "currentRegion":
      return result([currentRegion(wb, sheet, active)]);
    case "currentArray": {
      const sp = wb.engine.getSpill(sheet.id, active.row, active.col);
      if (!sp) return null;
      return result([{ r1: sp.anchorRow, c1: sp.anchorCol, r2: sp.anchorRow + sp.rows - 1, c2: sp.anchorCol + sp.cols - 1 }]);
    }
    case "objects": {
      const first = sheet.charts[0] ? { kind: "chart" as const, id: sheet.charts[0].id } : sheet.images[0] ? { kind: "image" as const, id: sheet.images[0].id } : sheet.shapes[0] ? { kind: "shape" as const, id: sheet.shapes[0].id } : null;
      return first ? { ranges: sel, active, object: first } : null;
    }
    case "rowDiffs":
    case "colDiffs": {
      // Her satırda (sütunda) etkin sütunun (satırın) değerinden farklı olan hücreler
      const key = (v: Scalar) => (v === null ? "" : typeof v === "string" ? "s" + v.toLocaleUpperCase("tr-TR") : typeof v + String(v));
      for (const a of area)
        for (let r = a.r1; r <= Math.min(a.r2, a.r1 + 100000); r++)
          for (let c = a.c1; c <= Math.min(a.c2, a.c1 + 500); c++) {
            const ref = kind === "rowDiffs" ? wb.getValue(sheet, r, active.col) : wb.getValue(sheet, active.row, c);
            const v = wb.getValue(sheet, r, c);
            if (v !== null && key(v) !== key(ref)) cells.push([r, c]);
          }
      return fromCells();
    }
    case "precedents": {
      const seen = new Set<number>();
      let frontier: CellPos[] = [];
      for (const s of sel) for (let r = s.r1; r <= Math.min(s.r2, s.r1 + 500); r++) for (let c = s.c1; c <= Math.min(s.c2, s.c1 + 100); c++) frontier.push({ row: r, col: c });
      const out: Rect[] = [];
      for (let depth = 0; frontier.length && depth < (o.allLevels ? 50 : 1); depth++) {
        const next: CellPos[] = [];
        for (const p of frontier)
          for (const pr of precedentsOf(wb.engine, sheet.id, p.row, p.col)) {
            if (pr.sheet !== sheet.id) continue;
            const r = { r1: pr.r1, c1: pr.c1, r2: Math.min(pr.r2, pr.r1 + 100000), c2: pr.c2 };
            out.push(r);
            for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 2000); i++)
              for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 200); j++) {
                const k = cellKey(i, j);
                if (!seen.has(k)) {
                  seen.add(k);
                  next.push({ row: i, col: j });
                }
              }
          }
        frontier = next;
      }
      return result(out);
    }
    case "dependents": {
      const seen = new Set<number>();
      let frontier: CellPos[] = [];
      for (const s of sel) for (let r = s.r1; r <= Math.min(s.r2, s.r1 + 500); r++) for (let c = s.c1; c <= Math.min(s.c2, s.c1 + 100); c++) frontier.push({ row: r, col: c });
      for (let depth = 0; frontier.length && depth < (o.allLevels ? 50 : 1); depth++) {
        const next: CellPos[] = [];
        for (const p of frontier)
          for (const d of dependentsOf(wb.engine, sheet.id, p.row, p.col)) {
            if (d.sheet !== sheet.id) continue;
            const k = cellKey(d.row, d.col);
            if (seen.has(k)) continue;
            seen.add(k);
            cells.push([d.row, d.col]);
            next.push({ row: d.row, col: d.col });
          }
        frontier = next;
      }
      return fromCells();
    }
    case "lastCell": {
      const u = wb.usedRange(sheet);
      return u ? result([{ r1: u.r2, c1: u.c2, r2: u.r2, c2: u.c2 }]) : null;
    }
    case "visible": {
      // Seçimi gizli satır/sütunlarda bölerek yalnız görünür parçalar
      const out: Rect[] = [];
      for (const s of sel) {
        const rowRuns: [number, number][] = [];
        let start = -1;
        const r2 = Math.min(s.r2, Math.max(s.r1, (wb.usedRange(sheet)?.r2 ?? s.r1)));
        for (let r = s.r1; r <= r2 + 1; r++) {
          const vis = r <= r2 && !sheet.rows.isHidden(r);
          if (vis && start < 0) start = r;
          if (!vis && start >= 0) {
            rowRuns.push([start, r - 1]);
            start = -1;
          }
        }
        const colRuns: [number, number][] = [];
        start = -1;
        const c2 = Math.min(s.c2, Math.max(s.c1, (wb.usedRange(sheet)?.c2 ?? s.c1)));
        for (let c = s.c1; c <= c2 + 1; c++) {
          const vis = c <= c2 && !sheet.cols.isHidden(c);
          if (vis && start < 0) start = c;
          if (!vis && start >= 0) {
            colRuns.push([start, c - 1]);
            start = -1;
          }
        }
        for (const [a, b] of rowRuns) for (const [c, d] of colRuns) out.push({ r1: a, c1: c, r2: b, c2: d });
      }
      return result(out);
    }
    case "condFormats":
    case "validation": {
      const list = kind === "condFormats" ? sheet.condFormats.map((c) => c.ranges) : sheet.validations.map((v) => v.ranges);
      const pick = o.same ? list.filter((rs) => rs.some((r) => contains(r, active.row, active.col))) : list;
      const out = pick.flat().filter((r) => area.length === 0 || area.some((a) => intersects(a, r)) || o.same);
      return result(out);
    }
  }
  return null;
}
