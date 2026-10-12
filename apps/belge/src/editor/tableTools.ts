// Tablo araçları: Başlık Satırlarını Yinele (sayfa sonlarında başlık satırının kopyası), tablo sıralama,
// Metni Tabloya / Tabloyu Metne Dönüştür ve varsayılan hücre kenar boşlukları.
import type { Editor } from "@tiptap/core";
import { Fragment, type Node as PMNode, type Schema } from "@tiptap/pm/model";
import { Selection, TextSelection, type EditorState } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";

/** Hücre kenar boşlukları (pt). */
export interface CellMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** Varsayılan: üst/alt 0, sol/sağ 0,19 cm (5,4 pt). */
export const DEFAULT_CELL_MARGINS: CellMargins = { top: 0, bottom: 0, left: 5.4, right: 5.4 };

export const cellPadCss = (m: CellMargins) => `${m.top}pt ${m.right}pt ${m.bottom}pt ${m.left}pt`;

export function findTable(state: EditorState): { node: PMNode; pos: number } | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === "table") return { node: n, pos: $from.before(d) };
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Başlık satırlarını yinele                                             */
/* ------------------------------------------------------------------ */

/** Satırdan önce sayfa sonu geliyorsa yeni sayfada yinelenecek başlık satırının yüksekliği (px; yoksa 0). */
export function repeatHeaderHeight(view: EditorView, rowPos: number, scale: number): number {
  const $r = view.state.doc.resolve(rowPos);
  const table = $r.parent;
  if (table.type.name !== "table" || !table.attrs.repeatHeader || $r.index() === 0) return 0;
  const tr = view.nodeDOM($r.start());
  if (!(tr instanceof HTMLElement)) return 0;
  return tr.getBoundingClientRect().height / (scale || 1);
}

const COPY_PROPS = [
  "background-color",
  "color",
  "font-weight",
  "border-top-color",
  "border-bottom-color",
  "border-left-color",
  "border-right-color",
  "border-top-width",
  "border-bottom-width",
  "border-top-style",
  "border-bottom-style",
];

/** Sayfa sonundan sonraki başlık satırı kopyası (düzenlenemez; tablonun ilk satırından kopyalanır). */
export function headerCloneWidget(): HTMLElement {
  const tr = document.createElement("tr");
  tr.className = "bl-pg bl-pg-rhead";
  tr.contentEditable = "false";
  tr.setAttribute("aria-hidden", "true");
  const fill = () => {
    const src = tr.closest("table")?.querySelector(":scope > tbody > tr:not(.bl-pg)");
    if (!src) return false;
    tr.replaceChildren();
    for (const cell of Array.from(src.children)) {
      if (!(cell instanceof HTMLElement)) continue;
      const c = cell.cloneNode(true) as HTMLElement;
      const cs = getComputedStyle(cell);
      for (const p of COPY_PROPS) c.style.setProperty(p, cs.getPropertyValue(p));
      c.classList.remove("selectedCell");
      c.querySelectorAll(".column-resize-handle, .ProseMirror-widget").forEach((e) => e.remove());
      tr.appendChild(c);
    }
    return true;
  };
  queueMicrotask(() => {
    if (!fill()) requestAnimationFrame(fill);
  });
  return tr;
}

/* ------------------------------------------------------------------ */
/* Sıralama                                                              */
/* ------------------------------------------------------------------ */

export type SortType = "text" | "number" | "date";

export interface SortOptions {
  /** Izgara sütun sırası. */
  col: number;
  type: SortType;
  desc: boolean;
  /** İlk satır üst bilgi satırı (sıralamaya katılmaz). */
  header: boolean;
}

/** Türkçe sayı ("1.234,5", "%12", "-3") → sayı; sayı yoksa NaN. */
export function parseNumber(s: string): number {
  const m = /-?[\d.,]+/.exec(s.replace(/\s/g, ""));
  if (!m) return NaN;
  let t = m[0];
  if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
  else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
  return parseFloat(t);
}

/** gg.aa.yyyy (/, - ayraçları da) → zaman damgası; tarih yoksa NaN. */
export function parseDate(s: string): number {
  const m = /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/.exec(s);
  if (!m) return NaN;
  const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
  return Date.UTC(y, Number(m[2]) - 1, Number(m[1]));
}

/** Satırın ızgara sütunundaki hücre metni. */
export function cellText(row: PMNode, col: number): string {
  let c = 0;
  let out = "";
  row.forEach((cell) => {
    const span = (cell.attrs.colspan as number) ?? 1;
    if (col >= c && col < c + span) out = cell.textContent.trim();
    c += span;
  });
  return out;
}

/** Sıralama karşılaştırıcısı: boş değerler artan sırada başa gelir (Word'deki gibi). */
export function compareKeys(a: string, b: string, type: SortType): number {
  if (type === "text") return a.localeCompare(b, "tr", { sensitivity: "base", numeric: true });
  const f = type === "number" ? parseNumber : parseDate;
  const x = f(a);
  const y = f(b);
  if (Number.isNaN(x) || Number.isNaN(y)) return Number.isNaN(x) === Number.isNaN(y) ? 0 : Number.isNaN(x) ? -1 : 1;
  return x - y;
}

/** Tablo satırlarını sıralar; dikey birleştirilmiş hücre varsa sıralanamaz (false). */
export function sortTableRows(table: PMNode, o: SortOptions): PMNode[] | null {
  const rows: PMNode[] = [];
  table.forEach((r) => rows.push(r));
  let merged = false;
  table.descendants((n) => {
    if ((n.attrs.rowspan as number) > 1) merged = true;
    return n.type.name === "table" || n.type.name === "tableRow";
  });
  if (merged) return null;
  const head = o.header ? rows.slice(0, 1) : [];
  const body = rows.slice(head.length).map((r, i) => ({ r, i, k: cellText(r, o.col) }));
  body.sort((a, b) => {
    const c = compareKeys(a.k, b.k, o.type);
    return (o.desc ? -c : c) || a.i - b.i;
  });
  return [...head, ...body.map((x) => x.r)];
}

export function sortTable(editor: Editor, o: SortOptions): boolean {
  const t = findTable(editor.state);
  if (!t) return false;
  const rows = sortTableRows(t.node, o);
  if (!rows) return false;
  const tr = editor.state.tr.replaceWith(t.pos + 1, t.pos + t.node.nodeSize - 1, Fragment.from(rows));
  // İmleç tabloda kalır (bağlamsal sekmeler açık kalsın): ilk hücrenin başı.
  tr.setSelection(TextSelection.near(tr.doc.resolve(t.pos + 4)));
  editor.view.dispatch(tr.scrollIntoView());
  return true;
}

/* ------------------------------------------------------------------ */
/* Metni Tabloya / Tabloyu Metne Dönüştür                                */
/* ------------------------------------------------------------------ */

/** Ayraç: "\t", ",", ";" ya da özel karakter; "para": her paragraf bir hücre. */
export type TextSeparator = string;

/** Seçimdeki üst düzey metin blokları ve aralığı. */
export function selectedParagraphs(state: EditorState): { from: number; to: number; nodes: PMNode[] } | null {
  // Tümünü Seç bölüm düzeyine çıkar; seçim metin bloklarına daraltılır.
  const $from = state.doc.resolve(Math.max(state.selection.from, Selection.atStart(state.doc).from));
  const $to = state.doc.resolve(Math.min(state.selection.to, Selection.atEnd(state.doc).to));
  if ($from.depth < 2 || $to.depth < 2) return null;
  if ($from.node(1) !== $to.node(1)) return null;
  const from = $from.before(2);
  const to = $to.after(2);
  const nodes: PMNode[] = [];
  state.doc.nodesBetween(from, to, (n, pos) => {
    if (pos < from) return true;
    if (n.isTextblock) nodes.push(n);
    return false;
  });
  if (!nodes.length || nodes.length !== $from.node(1).slice(from - $from.start(1), to - $from.start(1)).content.childCount) return null;
  return { from, to, nodes };
}

/** Önerilen ayraç: metinde sekme varsa sekme, yoksa noktalı virgül/virgül, yoksa paragraf. */
export function guessSeparator(nodes: PMNode[]): TextSeparator {
  const text = nodes.map((n) => n.textContent).join("\n");
  if (text.includes("\t")) return "\t";
  if (text.includes(";")) return ";";
  if (text.includes(",")) return ",";
  return "para";
}

/** Paragrafı ayraçtan böler (biçim korunur). */
function splitPara(n: PMNode, sep: string): PMNode[] {
  const text = n.textBetween(0, n.content.size, undefined, "￼");
  const parts: PMNode[] = [];
  let start = 0;
  for (let i = text.indexOf(sep); i >= 0; i = text.indexOf(sep, start)) {
    parts.push(n.type.create(n.attrs, n.content.cut(start, i)));
    start = i + sep.length;
  }
  parts.push(n.type.create(n.attrs, n.content.cut(start)));
  return parts;
}

/** Metin bloklarından tablo satırları (her satır hücre içerikleri). */
export function textRows(nodes: PMNode[], sep: TextSeparator, cols: number): PMNode[][] {
  if (sep === "para") {
    const rows: PMNode[][] = [];
    for (let i = 0; i < nodes.length; i += Math.max(1, cols)) rows.push(nodes.slice(i, i + Math.max(1, cols)));
    return rows;
  }
  return nodes.map((n) => splitPara(n, sep));
}

export function buildTable(schema: Schema, rows: PMNode[][], cols: number): PMNode {
  const n = Math.max(cols, ...rows.map((r) => r.length));
  const paragraph = (c?: PMNode) => (c ? (c.type.name === "paragraph" || c.type.name === "heading" ? c : schema.nodes.paragraph.create(null, c.content)) : schema.nodes.paragraph.create());
  const tableRows = rows.map((r) =>
    schema.nodes.tableRow.create(
      null,
      Array.from({ length: n }, (_, i) => schema.nodes.tableCell.create(null, paragraph(r[i]))),
    ),
  );
  return schema.nodes.table.create({ headerRow: false, bandedRows: false }, tableRows);
}

export function textToTable(editor: Editor, sep: TextSeparator, cols: number): boolean {
  const sel = selectedParagraphs(editor.state);
  if (!sel) return false;
  const table = buildTable(editor.schema, textRows(sel.nodes, sep, cols), cols);
  const tr = editor.state.tr.replaceWith(sel.from, sel.to, table);
  tr.setSelection(TextSelection.near(tr.doc.resolve(sel.from + 4)));
  editor.view.dispatch(tr.scrollIntoView());
  return true;
}

/** Tablo → paragraflar: hücreler ayraçla birleşir; "para" ise her hücrenin paragrafları ayrı kalır. */
export function tableToParagraphs(schema: Schema, table: PMNode, sep: TextSeparator): PMNode[] {
  const out: PMNode[] = [];
  table.forEach((row) => {
    if (sep === "para") {
      row.forEach((cell) => cell.forEach((b) => out.push(b.isTextblock ? b : schema.nodes.paragraph.create(null, schema.text(b.textContent || " ")))));
      return;
    }
    const inline: PMNode[] = [];
    let attrs: Record<string, unknown> | null = null;
    row.forEach((cell, _o, ci) => {
      if (ci > 0) inline.push(schema.text(sep));
      let first = true;
      cell.forEach((b) => {
        if (!b.isTextblock) return;
        attrs ??= b.attrs;
        if (!first && b.content.size) inline.push(schema.text(" "));
        b.content.forEach((x) => inline.push(x));
        first = false;
      });
    });
    out.push(schema.nodes.paragraph.create(attrs, inline));
  });
  return out;
}

export function tableToText(editor: Editor, sep: TextSeparator): boolean {
  const t = findTable(editor.state);
  if (!t) return false;
  const paras = tableToParagraphs(editor.schema, t.node, sep);
  const tr = editor.state.tr.replaceWith(t.pos, t.pos + t.node.nodeSize, paras);
  tr.setSelection(TextSelection.near(tr.doc.resolve(t.pos + 1)));
  editor.view.dispatch(tr.scrollIntoView());
  return true;
}
