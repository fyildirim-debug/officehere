// Hücreleri Böl (Word): seçili hücre istenen sütun × satır sayısına bölünür. Birleştirilmemiş hücrede
// tabloya yeni ızgara sütunları/satırları eklenir; diğer satır/sütunlardaki komşu hücreler genişleyerek
// görünümü korur. Birden çok hücre seçiliyse önce birleştirilir (Word: "Bölmeden önce birleştir").
import type { Node as PMNode, Schema } from "@tiptap/pm/model";
import { TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { CellSelection, TableMap, mergeCells } from "@tiptap/pm/tables";

interface GridCell {
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
  node: PMNode;
  /** Bölünen hücreden doğan yeni hücre (içerik ilkinde kalır). */
  fresh?: boolean;
}

/** İmleçteki (ya da hücre seçimindeki tek) hücrenin tablo içi konumu. */
function cellAt(state: EditorState): { tablePos: number; table: PMNode; cellPos: number } | null {
  const sel = state.selection;
  const $pos = sel instanceof CellSelection ? sel.$anchorCell : sel.$from;
  for (let d = $pos.depth; d > 0; d--) {
    const n = $pos.node(d);
    if (n.type.spec.tableRole === "table") {
      // Hücre: tablonun iki düzey altı (tablo > satır > hücre)
      const cellPos = sel instanceof CellSelection ? sel.$anchorCell.pos : $pos.before(d + 2);
      return { tablePos: $pos.before(d), table: n, cellPos };
    }
  }
  return null;
}

/** Bölme iletişim kutusunun başlangıç değerleri: hücrenin mevcut sütun/satır kapsamı. */
export function splitDefaults(state: EditorState): { cols: number; rows: number; multi: boolean } {
  const sel = state.selection;
  const multi = sel instanceof CellSelection && sel.ranges.length > 1;
  const at = cellAt(state);
  if (!at) return { cols: 2, rows: 1, multi };
  const cell = state.doc.nodeAt(at.cellPos);
  const cs = (cell?.attrs.colspan as number) ?? 1;
  const rs = (cell?.attrs.rowspan as number) ?? 1;
  return { cols: Math.max(2, cs), rows: rs, multi };
}

/** Sütun genişlikleri (px; bilinmeyen null) — hücrelerin colwidth dizilerinden. */
function gridWidths(map: TableMap, table: PMNode): (number | null)[] {
  const w: (number | null)[] = new Array(map.width).fill(null);
  const seen = new Set<number>();
  for (let i = 0; i < map.map.length; i++) {
    const off = map.map[i];
    if (seen.has(off)) continue;
    seen.add(off);
    const cell = table.nodeAt(off)!;
    const rect = map.findCell(off);
    const cw = cell.attrs.colwidth as number[] | null;
    if (!cw) continue;
    for (let k = 0; k < rect.right - rect.left; k++) if (w[rect.left + k] === null && cw[k]) w[rect.left + k] = cw[k];
  }
  return w;
}

/** Sayıyı parçalara olabildiğince eşit böler (ör. 5 → 3 parça: 2,2,1). */
function spread(total: number, parts: number): number[] {
  const base = Math.floor(total / parts);
  return Array.from({ length: parts }, (_, i) => base + (i < total % parts ? 1 : 0));
}

/** Tablo düğümünü ızgara modelinden yeniden kurar. */
function buildTable(schema: Schema, table: PMNode, cells: GridCell[], height: number, widths: (number | null)[], rowAttrs: (r: number) => Record<string, unknown> | null): PMNode {
  const rows: PMNode[] = [];
  for (let r = 0; r < height; r++) {
    const inRow = cells.filter((c) => c.row === r).sort((a, b) => a.col - b.col);
    const nodes = inRow.map((c) => {
      const cw = widths.slice(c.col, c.col + c.colspan);
      const colwidth = cw.every((x) => x !== null) ? (cw as number[]).map((x) => Math.round(x)) : null;
      const attrs = { ...c.node.attrs, colspan: c.colspan, rowspan: c.rowspan, colwidth };
      return c.fresh ? c.node.type.create(attrs, schema.nodes.paragraph.create()) : c.node.type.create(attrs, c.node.content, c.node.marks);
    });
    // Tümü üstten yayılan hücrelerle dolu satır da geçerlidir (boş satır).
    rows.push(schema.nodes.tableRow.create(rowAttrs(r), nodes));
  }
  return table.type.create(table.attrs, rows, table.marks);
}

/**
 * Seçili hücreyi `cols` sütun × `rows` satıra böler. Önce birden çok hücre seçiliyse birleştirilir.
 * Başarılıysa true; bölünemiyorsa (ör. hedef satır sayısı hücrenin kapsamını bölmüyor) false.
 */
export function splitCellGrid(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, cols: number, rows: number, mergeFirst = true): boolean {
  cols = Math.max(1, Math.min(63, Math.round(cols)));
  rows = Math.max(1, Math.min(63, Math.round(rows)));
  // Birden çok hücre: önce birleştir.
  let st = state;
  let pre: Transaction | null = null;
  if (mergeFirst && st.selection instanceof CellSelection && st.selection.ranges.length > 1) {
    let merged: Transaction | null = null;
    if (!mergeCells(st, (t) => (merged = t))) return false;
    pre = merged!;
    st = st.apply(pre);
  }
  const at = cellAt(st);
  if (!at) return false;
  const { table, tablePos } = at;
  const map = TableMap.get(table);
  const off = at.cellPos - tablePos - 1;
  const target = map.findCell(off);
  const cs = target.right - target.left;
  const rs = target.bottom - target.top;
  if (cols === 1 && rows === 1) return false;
  // Word kuralı: satır sayısı, birleştirilmiş hücrenin satır kapsamına bölünebilmeli ya da kapsam 1 olmalı.
  if (rs > 1 && rows > 1 && rs % rows !== 0 && rows < rs) return false;

  // Izgara modeli
  const seen = new Set<number>();
  const cells: GridCell[] = [];
  for (let i = 0; i < map.map.length; i++) {
    const o = map.map[i];
    if (seen.has(o)) continue;
    seen.add(o);
    const rect = map.findCell(o);
    cells.push({ row: rect.top, col: rect.left, rowspan: rect.bottom - rect.top, colspan: rect.right - rect.left, node: table.nodeAt(o)! });
  }
  const tCell = cells.find((c) => c.row === target.top && c.col === target.left)!;
  let widths = gridWidths(map, table);
  let height = map.height;

  // 1) Sütunlar
  const addC = Math.max(0, cols - cs);
  if (addC > 0) {
    const boundary = target.right - 1;
    for (const c of cells) {
      if (c === tCell) continue;
      if (c.col <= boundary && boundary < c.col + c.colspan) c.colspan += addC;
      else if (c.col > boundary) c.col += addC;
    }
    const W = widths.slice(target.left, target.right);
    const total = W.every((x) => x !== null) ? (W as number[]).reduce((a, b) => a + b, 0) : null;
    widths = [...widths.slice(0, target.left), ...new Array(cols).fill(total === null ? null : total / cols), ...widths.slice(target.right)];
  }
  const colSpans = addC > 0 ? new Array(cols).fill(1) : spread(cs, cols);
  // Hedef hücre sütun parçalarına ayrılır.
  const parts: GridCell[] = [];
  let x = target.left;
  colSpans.forEach((span, i) => {
    parts.push({ row: target.top, col: x, rowspan: rs, colspan: span, node: tCell.node, fresh: i > 0 });
    x += span;
  });
  cells.splice(cells.indexOf(tCell), 1, ...parts);

  // 2) Satırlar
  const addR = Math.max(0, rows - rs);
  if (addR > 0) {
    const boundary = target.bottom - 1;
    for (const c of cells) {
      if (parts.includes(c)) continue;
      if (c.row <= boundary && boundary < c.row + c.rowspan) c.rowspan += addR;
      else if (c.row > boundary) c.row += addR;
    }
    height += addR;
  }
  const rowSpans = addR > 0 ? new Array(rows).fill(1) : spread(rs, rows);
  for (const p of [...parts]) {
    let y = target.top;
    const stack: GridCell[] = rowSpans.map((span, i) => {
      const cell: GridCell = { row: y, col: p.col, rowspan: span, colspan: p.colspan, node: p.node, fresh: p.fresh || i > 0 };
      y += span;
      return cell;
    });
    cells.splice(cells.indexOf(p), 1, ...stack);
  }

  const srcRows: PMNode[] = [];
  table.forEach((r) => srcRows.push(r));
  const rowAttrs = (r: number) => {
    // Eklenen satırlar bölünen hücrenin son satırının özniteliklerini alır.
    const src = r < target.bottom ? r : r < target.bottom + addR ? target.bottom - 1 : r - addR;
    return srcRows[src]?.attrs ?? null;
  };
  const next = buildTable(st.schema, table, cells, height, widths, rowAttrs);
  if (dispatch) {
    // Birleştirme yapıldıysa st onun sonucudur (pre.doc === st.doc); adımlar aynı işleme eklenir.
    const base = pre ?? state.tr;
    base.replaceWith(tablePos, tablePos + table.nodeSize, next);
    // İmleç ilk parçanın içine.
    const newMap = TableMap.get(next);
    const firstOff = newMap.map[target.top * newMap.width + target.left];
    base.setSelection(TextSelection.near(base.doc.resolve(tablePos + 1 + firstOff + 1)));
    dispatch(base.scrollIntoView());
  }
  return true;
}
