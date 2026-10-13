/** ODF tablo okuyucusu (draw:frame > table:table): sütun genişlikleri, satır yükseklikleri, hücre dolgusu, kenarlar, birleştirmeler. */
import type { CellBorder, TableCell, TableRow, TableShape } from "../../model/types";
import { attr, kids } from "../xml";
import { lengthPt, lengthPx, odfColor } from "./common";
import type { ShapeBaseProps, ShapeReadCtx } from "./shapes";
import type { Props } from "./styles";
import { paragraphsOf } from "./text";

/** "0.48pt solid #ffffff" → hücre kenarı (none → null). */
function border(v: string | undefined): CellBorder | null | undefined {
  if (!v) return undefined;
  if (/^\s*none\s*$/i.test(v)) return null;
  const w = v.split(/\s+/).map((x) => lengthPt(x)).find((x) => x !== undefined) ?? 0.75;
  const c = odfColor(/#[0-9a-f]{6}/i.exec(v)?.[0]) ?? { rgb: "#000000" };
  return { color: c, width: w };
}

function cellBorders(pp: Props): TableCell["borders"] | undefined {
  const all = border(pp["border"]);
  const b = { l: border(pp["border-left"]) ?? all, r: border(pp["border-right"]) ?? all, t: border(pp["border-top"]) ?? all, b: border(pp["border-bottom"]) ?? all };
  return Object.values(b).some((x) => x !== undefined) ? b : undefined;
}

const MAX_COLS = 64;
const MAX_ROWS = 400;

export function readTable(table: Element, base: ShapeBaseProps, ctx: ShapeReadCtx): TableShape | undefined {
  const colWidths: (number | undefined)[] = [];
  const colDefault: (string | undefined)[] = [];
  const colEls = [...kids(table, "table-column"), ...kids(table, "table-columns").flatMap((c) => kids(c, "table-column"))];
  for (const c of colEls) {
    const rep = Math.max(1, Math.min(MAX_COLS, Number(attr(c, "number-columns-repeated") ?? "1") || 1));
    const cs = ctx.sheet.resolve("table-column", attr(c, "style-name"));
    const w = lengthPx(cs.graphic["column-width"]);
    for (let k = 0; k < rep && colWidths.length < MAX_COLS; k++) {
      colWidths.push(w);
      colDefault.push(attr(c, "default-cell-style-name"));
    }
  }
  const rowEls = [...kids(table, "table-row"), ...["table-rows", "table-header-rows"].flatMap((n) => kids(table, n).flatMap((c) => kids(c, "table-row")))];
  const rows: TableRow[] = [];
  let maxCols = colWidths.length;
  for (const r of rowEls) {
    if (rows.length >= MAX_ROWS) break;
    const rep = Math.max(1, Math.min(MAX_ROWS - rows.length, Number(attr(r, "number-rows-repeated") ?? "1") || 1));
    const rs = ctx.sheet.resolve("table-row", attr(r, "style-name"));
    const h = lengthPx(rs.graphic["row-height"]) ?? lengthPx(rs.graphic["min-row-height"]) ?? 40;
    const cells: TableCell[] = [];
    for (const c of kids(r)) {
      if (c.localName !== "table-cell" && c.localName !== "covered-table-cell") continue;
      const crep = Math.max(1, Math.min(MAX_COLS, Number(attr(c, "number-columns-repeated") ?? "1") || 1));
      for (let k = 0; k < crep && cells.length < MAX_COLS; k++) {
        if (c.localName === "covered-table-cell") {
          cells.push({ text: { paragraphs: [{ runs: [], level: 0, props: {} }] }, hMerge: true });
          continue;
        }
        const cs = ctx.sheet.resolve("table-cell", attr(c, "style-name") ?? attr(r, "default-cell-style-name") ?? colDefault[cells.length]);
        const cell: TableCell = { text: { paragraphs: paragraphsOf(c, { sheet: ctx.sheet, link: ctx.link }) } };
        const span = Number(attr(c, "number-columns-spanned") ?? "1");
        const rspan = Number(attr(c, "number-rows-spanned") ?? "1");
        if (span > 1) cell.gridSpan = Math.min(span, MAX_COLS);
        if (rspan > 1) cell.rowSpan = Math.min(rspan, MAX_ROWS);
        const bg = cs.graphic["fill"] === "solid" ? odfColor(cs.graphic["fill-color"]) : cs.graphic["background-color"] && cs.graphic["background-color"] !== "transparent" ? odfColor(cs.graphic["background-color"]) : undefined;
        if (bg) cell.fill = { type: "solid", color: bg };
        const borders = cellBorders(cs.paragraph);
        if (borders) cell.borders = borders;
        const va = cs.graphic["textarea-vertical-align"] ?? cs.paragraph["vertical-align"] ?? cs.graphic["vertical-align"];
        if (va === "middle") cell.anchor = "ctr";
        else if (va === "bottom") cell.anchor = "b";
        cells.push(cell);
      }
    }
    maxCols = Math.max(maxCols, cells.length);
    for (let k = 0; k < rep; k++) rows.push({ h, cells: k ? cells.map((c) => structuredClone(c)) : cells });
  }
  if (!rows.length || !maxCols) return undefined;
  // Satır içi birleştirmeler (rowSpan) için alttaki hücreler vMerge.
  rows.forEach((row) => {
    while (row.cells.length < maxCols) row.cells.push({ text: { paragraphs: [{ runs: [], level: 0, props: {} }] } });
  });
  rows.forEach((row, ri) =>
    row.cells.forEach((cell, ci) => {
      for (let k = 1; k < (cell.rowSpan ?? 1) && ri + k < rows.length; k++) rows[ri + k].cells[ci].vMerge = true;
    }),
  );
  const known = colWidths.filter((w): w is number => w !== undefined);
  const fallback = base.w / maxCols;
  const cols = Array.from({ length: maxCols }, (_, i) => colWidths[i] ?? (known.length ? known.reduce((a, b) => a + b, 0) / known.length : fallback));
  const totalW = cols.reduce((a, b) => a + b, 0);
  const scale = base.w > 0 && totalW > 0 ? base.w / totalW : 1;
  return {
    ...base,
    kind: "table",
    name: base.name === "Şekil" ? "Tablo" : base.name,
    cols: cols.map((c) => c * scale),
    rows,
    h: Math.max(base.h, rows.reduce((a, r) => a + r.h, 0)),
    style: { styleId: "none-grid", firstRow: false, lastRow: false, firstCol: false, lastCol: false, bandRow: false, bandCol: false },
  };
}
