// Yinele (F4 / Ctrl+Y): yinelenecek geri alınmış adım varsa onu yineler; yoksa son eylemi (Word'deki
// "Yinele: …") geçerli seçime yeniden uygular. Son eylem, kullanıcının son işleminden türetilir:
// karakter biçimi (işaret ekleme/kaldırma), paragraf biçimi (blok öznitelikleri / türü) ya da yazılan metin.
import { Extension, type Editor } from "@tiptap/core";
import type { Mark, MarkType, Node as PMNode, NodeType } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState, type Transaction } from "@tiptap/pm/state";
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from "@tiptap/pm/transform";
import { redoDepth } from "@tiptap/pm/history";
import { syncDomSelection } from "./pagination";

export type RepeatAction =
  | { kind: "marks"; add: Mark[]; remove: MarkType[] }
  | { kind: "block"; type: NodeType | null; attrs: Record<string, unknown> }
  | { kind: "text"; text: string; end: number };

export const repeatKey = new PluginKey<RepeatAction | null>("belgeRepeat");

/** Yalnız biçim değiştiren işlemde metin blokları: aynı sırada, aynı içerikle. */
function blockChange(tr: Transaction): RepeatAction | null {
  const a: PMNode[] = [];
  const b: PMNode[] = [];
  tr.before.descendants((n) => (n.isTextblock ? (a.push(n), false) : true));
  tr.doc.descendants((n) => (n.isTextblock ? (b.push(n), false) : true));
  if (a.length !== b.length) return null;
  let type: NodeType | null = null;
  let attrs: Record<string, unknown> | null = null;
  for (let i = 0; i < a.length; i++) {
    if (!a[i].content.eq(b[i].content)) return null;
    if (a[i].type === b[i].type && JSON.stringify(a[i].attrs) === JSON.stringify(b[i].attrs)) continue;
    const diff: Record<string, unknown> = {};
    for (const k of Object.keys(b[i].attrs)) if (JSON.stringify(a[i].attrs[k]) !== JSON.stringify(b[i].attrs[k])) diff[k] = b[i].attrs[k];
    const t = a[i].type !== b[i].type ? b[i].type : null;
    if (attrs && (JSON.stringify(attrs) !== JSON.stringify(diff) || t !== type)) return null;
    attrs = diff;
    type = t;
  }
  return attrs ? { kind: "block", type, attrs } : null;
}

/** İşlemden yinelenebilir eylem (yoksa null: önceki eylem korunmaz). */
export function actionOf(tr: Transaction, prev: RepeatAction | null): RepeatAction | null | undefined {
  if (!tr.docChanged) return undefined;
  const steps = tr.steps;
  if (steps.every((s) => s instanceof AddMarkStep || s instanceof RemoveMarkStep)) {
    const add: Mark[] = [];
    const remove: MarkType[] = [];
    for (const s of steps) {
      if (s instanceof AddMarkStep && !add.some((m) => m.eq(s.mark))) add.push(s.mark);
      if (s instanceof RemoveMarkStep && !remove.includes(s.mark.type) && !add.some((m) => m.type === s.mark.type)) remove.push(s.mark.type);
    }
    return { kind: "marks", add, remove };
  }
  // Yazma: tek adımda yalnız metin eklenir (seçili metnin yerine de olabilir).
  if (steps.length === 1 && steps[0] instanceof ReplaceStep) {
    const st = steps[0] as ReplaceStep & { from: number; to: number; slice: import("@tiptap/pm/model").Slice };
    const c = st.slice.content;
    if (st.slice.openStart === 0 && st.slice.openEnd === 0 && c.childCount > 0 && c.content.every((n) => n.isText)) {
      const text = c.textBetween(0, c.size);
      // Art arda yazılan karakterler tek eylemde birikir.
      if (prev?.kind === "text" && st.from === st.to && st.from === prev.end) return { kind: "text", text: prev.text + text, end: st.from + text.length };
      return { kind: "text", text, end: st.from + text.length };
    }
    return null;
  }
  return blockChange(tr);
}

/** Eylemi geçerli seçime uygular. */
export function applyAction(state: EditorState, a: RepeatAction): Transaction | null {
  const { from, to, empty } = state.selection;
  const tr = state.tr;
  if (a.kind === "text") {
    tr.insertText(a.text, from, to);
    return tr.scrollIntoView();
  }
  if (a.kind === "marks") {
    if (empty) {
      let set = state.storedMarks ?? state.selection.$from.marks();
      for (const t of a.remove) set = t.removeFromSet(set);
      for (const m of a.add) set = m.addToSet(set);
      return tr.setStoredMarks(set);
    }
    for (const t of a.remove) tr.removeMark(from, to, t);
    for (const m of a.add) tr.addMark(from, to, m);
    return tr;
  }
  let changed = false;
  state.doc.nodesBetween(from, to, (n, pos) => {
    if (!n.isTextblock) return true;
    const type = a.type ?? n.type;
    try {
      tr.setNodeMarkup(pos, type, { ...n.attrs, ...a.attrs });
      changed = true;
    } catch {
      // Bu blokta geçersiz (ör. tablo hücresinde başlık)
    }
    return false;
  });
  return changed ? tr : null;
}

/** F4 / Ctrl+Y: önce yinele, yoksa son eylemi yinele. */
export function repeatOrRedo(editor: Editor): boolean {
  syncDomSelection(editor.view);
  const { state, view } = editor;
  if (redoDepth(state) > 0) return editor.commands.redo();
  const a = repeatKey.getState(state);
  if (!a) return false;
  const tr = applyAction(state, a);
  if (!tr) return false;
  view.dispatch(tr.setMeta(repeatKey, "repeat"));
  return true;
}

export const Repeat = Extension.create({
  name: "belgeRepeat",
  priority: 1200,
  addKeyboardShortcuts() {
    return {
      F4: () => repeatOrRedo(this.editor),
      "Mod-y": () => repeatOrRedo(this.editor),
    };
  },
  addProseMirrorPlugins() {
    return [
      new Plugin<RepeatAction | null>({
        key: repeatKey,
        state: {
          init: () => null,
          apply(tr, prev) {
            // Yineleme, geri al/yinele ve kullanıcı dışı (sayfalama, alan güncelleme) işlemleri son eylemi değiştirmez.
            if (tr.getMeta(repeatKey) === "repeat" || tr.getMeta("history$") || tr.getMeta("addToHistory") === false || tr.getMeta("appendedTransaction")) {
              if (prev?.kind === "text" && tr.docChanged) return { ...prev, end: tr.mapping.map(prev.end) };
              return prev;
            }
            const a = actionOf(tr, prev);
            if (a === undefined) return prev;
            return a;
          },
        },
      }),
    ];
  },
});

/**
 * Kısayoldan önce bekleyen DOM seçimi okunur: Shift+End'in hemen ardından basılan Ctrl+B gibi
 * kısayollar eski (henüz okunmamış) seçime uygulanmasın. Tüm tuş eşlemlerinden önce çalışır.
 */
export const SelectionSync = Extension.create({
  name: "belgeSelectionSync",
  priority: 10000,
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("belgeSelectionSync"),
        props: {
          handleKeyDown: (view, event) => {
            if (event.ctrlKey || event.metaKey || event.altKey || /^F\d+$/.test(event.key)) syncDomSelection(view);
            return false;
          },
        },
      }),
    ];
  },
});
