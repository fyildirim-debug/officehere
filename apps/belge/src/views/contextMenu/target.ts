// Bağlam menüsünün hedefi: sağ tıklanan nesne (resim, grafik, denklem, metin kutusu, içindekiler, kaynakça),
// tablo hücresi ya da metin. Nesneye sağ tıklanınca nesne seçilir (Word'deki gibi tutamaçlar görünür);
// klavyeyle (Shift+F10, Menü tuşu) açılınca hedef geçerli seçimden okunur.
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, TextSelection, type EditorState } from "@tiptap/pm/state";
import { CellSelection } from "@tiptap/pm/tables";
import type { EditorView } from "@tiptap/pm/view";

/** Seçilerek kendi menüsünü açan nesneler. */
export type ObjectKind = "image" | "chart" | "equation" | "textBox" | "toc" | "bibliography";
export type MenuKind = "text" | "table" | ObjectKind;

const OBJECTS = new Set<string>(["image", "chart", "equation", "textBox", "toc", "bibliography"]);
/** "Alanı Güncelleştir" ile yenilenen satır içi alanlar. */
const UPDATABLE_FIELDS = new Set<string>(["seqField", "crossRef", "citation", "pageField"]);

export interface MenuTarget {
  kind: MenuKind;
  /** Nesnenin (ya da alanın) konumu ve düğümü. */
  pos: number;
  node: PMNode | null;
  /** Metin menüsü: imleç metin kutusunun içinde. */
  inTextBox: boolean;
  /** Metin menüsü: sağ tıklanan (ya da imleç yanındaki) alan. */
  field: { pos: number; node: PMNode } | null;
}

interface Hit {
  pos: number;
  node: PMNode;
}

/** DOM öğesini içeren en içteki nesne ya da alan düğümü. Metin kutusunun içeriğine tıklanınca kutu sayılmaz. */
export function nodeAtDom(view: EditorView, target: Node): { object: Hit | null; field: Hit | null } {
  let object: Hit | null = null;
  let field: Hit | null = null;
  if (!view.dom.contains(target) || target === view.dom) return { object, field };
  view.state.doc.descendants((n, pos) => {
    const name = n.type.name;
    if (!OBJECTS.has(name) && !UPDATABLE_FIELDS.has(name)) return true;
    const dom = view.nodeDOM(pos);
    if (!(dom instanceof HTMLElement) || !dom.contains(target)) return false;
    if (UPDATABLE_FIELDS.has(name)) field = { pos, node: n };
    else if (name === "textBox") {
      // Kutunun içindeki yazı metin menüsünü açar; kenar ve tutamaçlar kutunun kendisini seçer.
      const content = dom.querySelector(".bl-tb-content");
      object = content?.contains(target) ? null : { pos, node: n };
    } else object = { pos, node: n };
    return true;
  });
  return { object, field };
}

function tableDepth(state: EditorState): number {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const name = $from.node(d).type.name;
    if (name === "table") return d;
    if (name === "textBox") return -1;
  }
  return -1;
}

function inTextBox(state: EditorState): boolean {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.name === "textBox") return true;
  return false;
}

/** Geçerli seçime göre menü hedefi (klavyeyle açma ve sağ tıktan sonra). */
export function targetOfSelection(state: EditorState, field: Hit | null = null): MenuTarget {
  const sel = state.selection;
  if (sel instanceof NodeSelection) {
    const name = sel.node.type.name;
    if (OBJECTS.has(name)) return { kind: name as ObjectKind, pos: sel.from, node: sel.node, inTextBox: false, field: null };
    if (name === "table") return { kind: "table", pos: sel.from, node: sel.node, inTextBox: false, field: null };
    if (UPDATABLE_FIELDS.has(name)) field ??= { pos: sel.from, node: sel.node };
  }
  if (!field && sel.empty) {
    // İmlecin hemen yanındaki alan (klavyeyle açınca).
    const { $from } = sel;
    const after = $from.nodeAfter;
    const before = $from.nodeBefore;
    if (after && UPDATABLE_FIELDS.has(after.type.name)) field = { pos: $from.pos, node: after };
    else if (before && UPDATABLE_FIELDS.has(before.type.name)) field = { pos: $from.pos - before.nodeSize, node: before };
  }
  const table = sel instanceof CellSelection || tableDepth(state) > 0;
  return { kind: table ? "table" : "text", pos: sel.from, node: null, inTextBox: !table && inTextBox(state), field };
}

/**
 * Tıklanan belge konumu. Metin kutusu altındaki metnin üstünde serbest durduğundan kutunun içindeki
 * tıklama tarayıcının karet konumundan eşlenir (yoksa alttaki paragrafa düşebilir).
 */
function hitAt(view: EditorView, target: Node, coords: { left: number; top: number }): { pos: number } | null {
  const content = target instanceof Element ? target.closest(".bl-tb-content") : target.parentElement?.closest(".bl-tb-content");
  if (content && view.dom.contains(content)) {
    const doc = document as Document & { caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null };
    try {
      const cp = doc.caretPositionFromPoint?.(coords.left, coords.top);
      const cr = cp ? null : document.caretRangeFromPoint?.(coords.left, coords.top);
      const node = cp?.offsetNode ?? cr?.startContainer;
      if (node && content.contains(node)) return { pos: view.posAtDOM(node, cp?.offset ?? cr?.startOffset ?? 0) };
    } catch {
      // karet konumu alınamadı: düzenleyicinin eşlemesine düşülür
    }
  }
  return view.posAtCoords(coords);
}

/**
 * Sağ tık: nesneye tıklandıysa nesneyi seçer; metne tıklandıysa imleci (seçimin dışındaysa) oraya taşır.
 * Döner: açılacak menünün hedefi.
 */
export function selectContextTarget(view: EditorView, target: Node, coords: { left: number; top: number } | null): MenuTarget {
  const { object, field } = nodeAtDom(view, target);
  const { state } = view;
  const sel = state.selection;
  if (object) {
    if (!(sel instanceof NodeSelection && sel.from === object.pos)) view.dispatch(state.tr.setSelection(NodeSelection.create(state.doc, object.pos)));
    return targetOfSelection(view.state);
  }
  const hit = coords ? hitAt(view, target, coords) : null;
  if (hit) {
    const { from, to } = sel;
    // Seçili tablo (Tabloyu Seç) içindeki tıklama seçimi korur; seçili nesnenin dışına tıklama seçimi bırakır.
    const keepNode = sel instanceof NodeSelection && (sel.node.type.name === "table" || (field && sel.from === field.pos)) && hit.pos >= from && hit.pos <= to;
    const outside = hit.pos < from || hit.pos > to;
    if (sel instanceof NodeSelection ? !keepNode : outside) {
      view.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(hit.pos))));
    }
  }
  return targetOfSelection(view.state, field);
}

/** Klavyeyle açılan menünün ekran konumu: seçili nesnenin sol altı ya da imleç. */
export function menuPointForSelection(view: EditorView): { x: number; y: number } {
  const sel = view.state.selection;
  if (sel instanceof NodeSelection) {
    const dom = view.nodeDOM(sel.from);
    if (dom instanceof HTMLElement) {
      const r = dom.getBoundingClientRect();
      return { x: r.left + Math.min(24, r.width / 2), y: Math.min(r.bottom, r.top + 40) };
    }
  }
  try {
    const c = view.coordsAtPos(sel.head);
    return { x: c.left, y: c.bottom };
  } catch {
    const r = view.dom.getBoundingClientRect();
    return { x: r.left + 24, y: r.top + 24 };
  }
}
