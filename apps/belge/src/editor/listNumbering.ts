// Numaralandırma denetimi (Word, LibreOffice ve Google Dokümanlar'daki sağ tık komutları):
// Numaralandırmaya Yeniden Başla, Numaralandırmaya Devam Et, Numaralandırma Değerini Ayarla.
// Her numaralı liste düğümü kendi başlangıç değerini (start) taşır; ortadaki öğeden yeniden başlatmak
// listeyi o öğeden ikiye böler.
import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorState, Transaction } from "@tiptap/pm/state";
import { canJoin, canSplit } from "@tiptap/pm/transform";

export interface OrderedAt {
  /** İmlecin içinde bulunduğu en içteki numaralı liste. */
  list: PMNode;
  pos: number;
  /** İmleçteki liste öğesinin listedeki sırası (0 tabanlı). */
  index: number;
  /** Listenin derinliği (iç içe listelerde aynı düzeydeki önceki listeyi bulmak için). */
  depth: number;
}

export function orderedListAt(state: EditorState): OrderedAt | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === "orderedList") return { list: n, pos: $from.before(d), index: $from.index(d), depth: d };
  }
  return null;
}

const startOf = (n: PMNode) => Number(n.attrs.start) || 1;

/** Öğedeki numara (listenin başlangıcı + sırası). */
export function itemNumber(at: OrderedAt): number {
  return startOf(at.list) + at.index;
}

/**
 * Aynı düzeydeki (aynı derinlikte) önceki numaralı liste: belge sırasında bu listeden önce biten son liste.
 * Araya giren paragraflar, başlıklar ve başka listeler numarayı kesmez (Word'deki "önceki listeden devam").
 */
export function previousOrderedList(state: EditorState, at: OrderedAt): { node: PMNode; pos: number } | null {
  let found: { node: PMNode; pos: number } | null = null;
  state.doc.nodesBetween(0, at.pos, (n, pos) => {
    if (pos >= at.pos) return false;
    if (n.type.name === "orderedList" && pos + n.nodeSize <= at.pos && state.doc.resolve(pos).depth + 1 === at.depth) found = { node: n, pos };
    return true;
  });
  return found;
}

/** İmleçteki öğeden itibaren numarayı `value` ile başlatır (gerekirse listeyi o öğeden böler). */
export function restartNumberingTr(state: EditorState, value = 1): Transaction | null {
  const at = orderedListAt(state);
  if (!at) return null;
  const tr = state.tr;
  const attrs = { ...at.list.attrs, start: Math.max(0, Math.round(value)) };
  if (at.index === 0) {
    tr.setNodeMarkup(at.pos, undefined, attrs);
    return tr;
  }
  // Öğeler arasındaki sınır: listenin içeriği + önceki öğelerin boyu.
  let split = at.pos + 1;
  for (let i = 0; i < at.index; i++) split += at.list.child(i).nodeSize;
  if (!canSplit(tr.doc, split, 1, [{ type: at.list.type, attrs }])) return null;
  tr.split(split, 1, [{ type: at.list.type, attrs }]);
  return tr;
}

/**
 * Önceki numaralı listeden devam eder. Liste önceki listenin hemen ardındaysa ikisi birleşir; araya
 * başka bloklar giriyorsa başlangıç değeri önceki listenin son numarasının bir fazlası olur.
 */
export function continueNumberingTr(state: EditorState): Transaction | null {
  const at = orderedListAt(state);
  if (!at) return null;
  const prev = previousOrderedList(state, at);
  if (!prev) return null;
  const tr = state.tr;
  if (prev.pos + prev.node.nodeSize === at.pos && prev.node.attrs.listStyle === at.list.attrs.listStyle && prev.node.attrs.multilevel === at.list.attrs.multilevel && canJoin(tr.doc, at.pos)) {
    tr.join(at.pos);
    return tr;
  }
  const next = startOf(prev.node) + prev.node.childCount;
  if (startOf(at.list) === next) return null;
  tr.setNodeMarkup(at.pos, undefined, { ...at.list.attrs, start: next });
  return tr;
}

export function restartNumbering(editor: Editor, value = 1): boolean {
  const tr = restartNumberingTr(editor.state, value);
  if (!tr) return false;
  editor.view.dispatch(tr.scrollIntoView());
  editor.commands.focus();
  return true;
}

export function continueNumbering(editor: Editor): boolean {
  const tr = continueNumberingTr(editor.state);
  if (!tr) return false;
  editor.view.dispatch(tr.scrollIntoView());
  editor.commands.focus();
  return true;
}

/** Önceki listeden devam edilebilir mi (menü öğesinin etkinliği). */
export function canContinueNumbering(state: EditorState): boolean {
  const at = orderedListAt(state);
  if (!at) return false;
  const prev = previousOrderedList(state, at);
  if (!prev) return false;
  return prev.pos + prev.node.nodeSize === at.pos || startOf(at.list) !== startOf(prev.node) + prev.node.childCount;
}
