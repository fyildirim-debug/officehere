// Tablo yardımcıları.
import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { CellSelection, TableMap } from "@tiptap/pm/tables";
import { NodeSelection } from "@tiptap/pm/state";

export function insertTable(editor: Editor, rows: number, cols: number, colWidthPx?: number) {
  editor.chain().focus().insertTable({ rows, cols, withHeaderRow: false }).run();
  if (colWidthPx) setAllColumnWidths(editor, colWidthPx);
}

function findTable(editor: Editor): { node: PMNode; pos: number } | null {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === "table") return { node: n, pos: $from.before(d) };
  }
  return null;
}

/** Tüm sütunlara aynı genişliği verir (px); null ise otomatik (pencereye sığdır). */
export function setAllColumnWidths(editor: Editor, w: number | null) {
  const t = findTable(editor);
  if (!t) return;
  const tr = editor.state.tr;
  t.node.descendants((cell, pos) => {
    if (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader") return true;
    const span = cell.attrs.colspan as number;
    tr.setNodeMarkup(t.pos + 1 + pos, undefined, { ...cell.attrs, colwidth: w ? Array(span).fill(Math.round(w)) : null });
    return false;
  });
  editor.view.dispatch(tr);
}

/** Sütunları dağıt: toplam genişliği eşit böler. */
export function distributeColumns(editor: Editor, contentWidth: number) {
  const t = findTable(editor);
  if (!t) return;
  const map = TableMap.get(t.node);
  setAllColumnWidths(editor, contentWidth / map.width);
}

/** İçeriğe otomatik sığdır: genişlikleri kaldırır ve tabloyu içerik kadar daraltır. */
export function autoFitContents(editor: Editor) {
  setAllColumnWidths(editor, null);
  editor.chain().focus().updateAttributes("table", { align: null }).run();
}

export function selectTablePart(editor: Editor, part: "cell" | "row" | "col" | "table") {
  const t = findTable(editor);
  if (!t) return;
  const { state } = editor;
  const $cell = (() => {
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) {
      const n = $from.node(d).type.name;
      if (n === "tableCell" || n === "tableHeader") return state.doc.resolve($from.before(d));
    }
    return null;
  })();
  if (!$cell) return;
  let sel;
  if (part === "cell") sel = new CellSelection($cell);
  else if (part === "row") sel = CellSelection.rowSelection($cell);
  else if (part === "col") sel = CellSelection.colSelection($cell);
  else sel = NodeSelection.create(state.doc, t.pos);
  editor.view.dispatch(state.tr.setSelection(sel));
  editor.commands.focus();
}

/** Hücre hizalaması: dikey (hücre) + yatay (paragraflar). */
export function alignCells(editor: Editor, v: "top" | "middle" | "bottom", h: "left" | "center" | "right") {
  editor
    .chain()
    .focus()
    .setCellAttribute("verticalAlign", v === "top" ? null : v)
    .run();
  const { state } = editor;
  const tr = state.tr;
  const sel = state.selection;
  const ranges: { from: number; to: number }[] = [];
  if (sel instanceof CellSelection) sel.forEachCell((cell, pos) => ranges.push({ from: pos, to: pos + cell.nodeSize }));
  else ranges.push({ from: sel.from, to: sel.to });
  for (const r of ranges) {
    state.doc.nodesBetween(r.from, r.to, (n, pos) => {
      if (n.type.name === "paragraph" || n.type.name === "heading") {
        tr.setNodeMarkup(pos, undefined, { ...n.attrs, textAlign: h === "left" ? null : h });
        return false;
      }
      return true;
    });
  }
  editor.view.dispatch(tr);
}

/** Satırları dağıt: tüm satırlar aynı yükseklikte (CSS ile en yüksek satıra eşitlenir). */
export function tableInfo(editor: Editor) {
  const t = findTable(editor);
  if (!t) return null;
  const map = TableMap.get(t.node);
  return { rows: map.height, cols: map.width, node: t.node, pos: t.pos };
}
