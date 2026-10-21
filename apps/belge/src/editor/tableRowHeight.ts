// Tablo satır yüksekliği (Word'deki gibi): satırın alt kenarını sürükleme, Düzen › Hücre Boyutu › Yükseklik ve
// Satırları Dağıt. Yükseklik "en az" kuralıdır: içerik uzunsa satır yine uzar.
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { CellSelection, TableMap } from "@tiptap/pm/tables";
import type { EditorView } from "@tiptap/pm/view";

/** En küçük satır yüksekliği (px). */
export const MIN_ROW_PX = 8;

/** Satır kenarının yakalandığı uzaklık (ekran pikseli). */
const GRAB = 3;

/** Tablodaki satırların belgedeki konumları. */
export function rowPositions(table: PMNode, tablePos: number): number[] {
  const out: number[] = [];
  table.forEach((_, off) => out.push(tablePos + 1 + off));
  return out;
}

/** Öğenin yakınlaştırma ölçeği (ekran pikseli / CSS pikseli). */
export function zoomOf(el: HTMLElement): number {
  const r = el.getBoundingClientRect().width;
  return el.offsetWidth && r ? r / el.offsetWidth : 1;
}

/** Satırların ekrandaki yükseklikleri (CSS px; yakınlaştırma düzeltilmiş). */
export function measureRows(view: EditorView, tablePos: number): number[] {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table || table.type.name !== "table") return [];
  return rowPositions(table, tablePos).map((p) => {
    const el = view.nodeDOM(p) as HTMLElement | null;
    return el ? el.getBoundingClientRect().height / (zoomOf(el) || 1) : 0;
  });
}

/** Satır yüksekliklerini işleme yazar (null: yükseklik kaldırılır); değişmeyen satıra dokunmaz. */
export function rowHeightsTr(tr: Transaction, tablePos: number, heights: (number | null)[]): Transaction {
  const table = tr.doc.nodeAt(tablePos);
  if (!table || table.type.name !== "table") return tr;
  rowPositions(table, tablePos).forEach((p, i) => {
    const want = heights[i];
    if (want === undefined) return;
    const row = tr.doc.nodeAt(p);
    const h = want === null ? null : Math.max(MIN_ROW_PX, Math.round(want * 100) / 100);
    if (!row || row.attrs.height === h) return;
    tr.setNodeMarkup(p, undefined, { ...row.attrs, height: h, heightRule: h ? row.attrs.heightRule : null });
  });
  return tr;
}

/** İmlecin ya da seçimin kapsadığı satırlar; hücre seçimi yoksa (ya da tüm tablo seçiliyse) `all` true. */
export function selectedRows(state: EditorState): { tablePos: number; table: PMNode; rows: number[]; all: boolean } | null {
  const sel = state.selection;
  if (sel instanceof CellSelection) {
    const $a = sel.$anchorCell;
    const table = $a.node(-1);
    const start = $a.start(-1);
    const rect = TableMap.get(table).rectBetween($a.pos - start, sel.$headCell.pos - start);
    const rows = Array.from({ length: rect.bottom - rect.top }, (_, i) => rect.top + i);
    return { tablePos: $a.before(-1), table, rows, all: rows.length === table.childCount };
  }
  if (sel instanceof NodeSelection && sel.node.type.name === "table") {
    return { tablePos: sel.from, table: sel.node, rows: Array.from({ length: sel.node.childCount }, (_, i) => i), all: true };
  }
  const { $from } = sel;
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type.spec.tableRole !== "row") continue;
    return { tablePos: $from.before(d - 1), table: $from.node(d - 1), rows: [$from.index(d - 1)], all: false };
  }
  return null;
}

/** Seçili satırların ölçülen yüksekliği (ilk satır; CSS px) — Düzen › Yükseklik kutusunda gösterilir. */
export function currentRowHeight(view: EditorView): number | null {
  const s = selectedRows(view.state);
  if (!s) return null;
  const row = s.table.child(s.rows[0]);
  return (row.attrs.height as number | null) ?? measureRows(view, s.tablePos)[s.rows[0]] ?? null;
}

/** Seçili satırlara yükseklik verir (px; null: otomatik). */
export function setRowHeight(view: EditorView, px: number | null): boolean {
  const s = selectedRows(view.state);
  if (!s) return false;
  const heights: (number | null | undefined)[] = [];
  for (const r of s.rows) heights[r] = px;
  view.dispatch(rowHeightsTr(view.state.tr, s.tablePos, heights as (number | null)[]));
  return true;
}

/**
 * Satırları Dağıt: seçili satırlar (hücre seçimi yoksa tüm tablo) toplam yükseklikleri korunarak eşit
 * yükseklik alır. İçeriği sığmayan satır yine içeriği kadar uzun kalır.
 */
export function distributeRows(view: EditorView): boolean {
  const s = selectedRows(view.state);
  if (!s) return false;
  const rows = s.rows.length > 1 ? s.rows : Array.from({ length: s.table.childCount }, (_, i) => i);
  const measured = measureRows(view, s.tablePos);
  const avg = rows.reduce((sum, r) => sum + (measured[r] ?? 0), 0) / rows.length;
  if (!(avg > 0)) return false;
  const heights: (number | null | undefined)[] = [];
  for (const r of rows) heights[r] = Math.round(avg * 100) / 100;
  view.dispatch(rowHeightsTr(view.state.tr, s.tablePos, heights as (number | null)[]));
  return true;
}

/** Satırın (birleştirilmiş hücrede kapsadığı son satırın) öğesi; sayfa sonu ara satırları atlanır. */
function lastRowOf(cell: HTMLElement): HTMLElement | null {
  let tr = cell.closest("tr") as HTMLElement | null;
  for (let k = Number(cell.getAttribute("rowspan")) || 1; k > 1 && tr; k--) {
    let next = tr.nextElementSibling as HTMLElement | null;
    while (next && next.classList.contains("bl-pg")) next = next.nextElementSibling as HTMLElement | null;
    tr = next ?? tr;
    if (!next) break;
  }
  return tr;
}

/**
 * Fare bir satırın alt kenarında mı? Kenarın iki yanında `GRAB` piksel içinde, kenarın üstündeki hücre
 * okunur (tablonun hemen altındaki boşluk da son satırın kenarıdır). Satırın konumu ve öğesi döner.
 */
export function rowBorderAt(view: EditorView, x: number, y: number): { rowPos: number; el: HTMLElement } | null {
  const hit = document.elementFromPoint(x, y - GRAB - 1) as HTMLElement | null;
  const cell = hit?.closest?.("td, th") as HTMLElement | null;
  if (!cell || !view.dom.contains(cell) || cell.closest(".bl-pg")) return null;
  const el = lastRowOf(cell);
  if (!el || el.classList.contains("bl-pg")) return null;
  const bottom = el.getBoundingClientRect().bottom;
  if (Math.abs(bottom - y) > GRAB) return null;
  let rowPos = -1;
  try {
    rowPos = view.posAtDOM(el, 0) - 1;
  } catch {
    return null;
  }
  const row = view.state.doc.nodeAt(rowPos);
  return row && row.type.spec.tableRole === "row" ? { rowPos, el } : null;
}
