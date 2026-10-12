/** Bul ve Değiştir araması (saf; arayüzden bağımsız). */
import { cellKey, keyCol, keyRow, wildcardToRegExp } from "@officehere/formula";
import { displayText } from "./display";
import { contains } from "./rect";
import type { Rect } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

export interface FindOptions {
  query: string;
  matchCase: boolean;
  wholeCell: boolean;
  lookIn: "formulas" | "values" | "notes";
  byColumns: boolean;
  /** Yalnız bu aralıklarda ara (çok hücreli seçim). */
  within?: Rect[];
}

export interface FindHit {
  sheet: SheetModel;
  row: number;
  col: number;
  text: string;
}

/** Arama deseni (joker karakterler: * ? ~). */
export function findPattern(o: Pick<FindOptions, "query" | "matchCase" | "wholeCell">, global = false): RegExp {
  const re = wildcardToRegExp(o.query, o.wholeCell);
  const flags = "u" + (o.matchCase ? "" : "i") + (global ? "g" : "");
  return new RegExp(re.source, flags);
}

/** Aranacak hücre metni. */
export function cellSearchText(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, lookIn: FindOptions["lookIn"]): string {
  if (lookIn === "notes") return sheet.noteAt(row, col)?.text ?? "";
  if (lookIn === "formulas") return wb.plainText(sheet, row, col) || displayText(wb, sheet, row, col);
  return displayText(wb, sheet, row, col);
}

export function findAll(wb: WorkbookModel, sheets: SheetModel[], o: FindOptions): FindHit[] {
  if (!o.query) return [];
  const re = findPattern(o);
  const hits: FindHit[] = [];
  for (const sheet of sheets) {
    const cells: [number, number][] = [];
    if (o.lookIn === "notes") {
      for (const k of sheet.notes.keys()) cells.push([keyRow(k), keyCol(k)]);
    } else {
      const seen = new Set<number>();
      wb.engine.forEachValue(sheet.id, (r, c, v) => {
        if (v === null) return;
        const k = cellKey(r, c);
        if (!seen.has(k)) (seen.add(k), cells.push([r, c]));
      });
    }
    const sheetHits: FindHit[] = [];
    for (const [row, col] of cells) {
      if (o.within && !o.within.some((r) => contains(r, row, col))) continue;
      const text = cellSearchText(wb, sheet, row, col, o.lookIn);
      if (text && re.test(text)) sheetHits.push({ sheet, row, col, text });
    }
    sheetHits.sort((a, b) => (o.byColumns ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col));
    hits.push(...sheetHits);
  }
  return hits;
}

/** Etkin hücreden sonraki (ya da önceki) eşleşmenin dizini. */
export function nextHitIndex(hits: FindHit[], sheet: SheetModel, row: number, col: number, byColumns: boolean, backward = false): number {
  if (!hits.length) return -1;
  const order = (h: { row: number; col: number }) => (byColumns ? h.col * 2 ** 21 + h.row : h.row * 16384 + h.col);
  const cur = order({ row, col });
  const sheetIdx = (s: SheetModel) => hits.findIndex((h) => h.sheet === s);
  const start = sheetIdx(sheet);
  if (!backward) {
    for (let i = Math.max(0, start); i < hits.length; i++) if (hits[i].sheet !== sheet || order(hits[i]) > cur) return i;
    return 0;
  }
  for (let i = hits.length - 1; i >= 0; i--) if ((hits[i].sheet === sheet && order(hits[i]) < cur) || (start >= 0 && i < start)) return i;
  return hits.length - 1;
}
