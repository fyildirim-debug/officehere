// Otomatik Düzelt eklentisi: sözcük bitince (boşluk, noktalama, Enter, Sekme) son sözcüğü düzeltir;
// düz tırnakları akıllı tırnağa çevirir. Düzeltme ayrı bir geri alma adımıdır (Ctrl+Z yalnız onu geri alır).
// Paragraf başında Sekme ilk satır girintisini, Geri Sil girintiyi azaltır (Word'ün otomatik biçimi).
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { closeHistory } from "@tiptap/pm/history";
import type { EditorView } from "@tiptap/pm/view";
import { DELIMITER, correctBefore, getAutoCorrect, smartQuote } from "./autocorrectRules";

export const autoCorrectKey = new PluginKey("belgeAutoCorrect");

/** İmleçten önceki paragraf metni (satır içi düğümler tek karakter sayılır, konumlar örtüşür). */
function textBefore(state: EditorState, pos: number): { text: string; start: number; parent: PMNode; cell: boolean } | null {
  const $p = state.doc.resolve(pos);
  if (!$p.parent.isTextblock || $p.parent.type.spec.code) return null;
  const text = $p.parent.textBetween(0, $p.parentOffset, undefined, (n: PMNode) => (n.type.name === "hardBreak" ? "\n" : "￼"));
  let cell = false;
  for (let d = $p.depth - 1; d > 0; d--) {
    const n = $p.node(d);
    if (n.type.name === "tableCell" || n.type.name === "tableHeader") {
      cell = $p.index(d) === 0;
      break;
    }
  }
  return { text, start: $p.start(), parent: $p.parent, cell };
}

/** İmleçten önceki sözcük için yapılacak düzeltme (belge konumlarıyla); yoksa null. */
function pendingCorrection(state: EditorState, pos: number): { from: number; to: number; text: string } | null {
  const info = textBefore(state, pos);
  if (!info) return null;
  const c = correctBefore(info.text, getAutoCorrect(), info.cell);
  // Düzeltilecek metin satır içi düğüm içermemeli.
  if (!c || info.text.slice(c.start, c.end).includes("\ufffc")) return null;
  return { from: info.start + c.start, to: info.start + c.end, text: c.text };
}

/** Düzeltmeyi ayrı bir geri alma adımı olarak uygular. */
function dispatchCorrection(view: EditorView, fix: { from: number; to: number; text: string }) {
  const tr = view.state.tr.insertText(fix.text, fix.from, fix.to);
  closeHistory(tr);
  tr.setMeta(autoCorrectKey, true);
  view.dispatch(tr);
  // Sonraki yazma düzeltmeyle aynı geri alma adımına eklenmez.
  view.dispatch(closeHistory(view.state.tr));
}

/** İmleçteki sözcüğü düzeltir (ayrı geri alma adımı). Düzeltme yapıldıysa true. */
export function applyAutoCorrect(view: EditorView, pos: number): boolean {
  const fix = pendingCorrection(view.state, pos);
  if (fix) dispatchCorrection(view, fix);
  return !!fix;
}

export const AutoCorrect = Extension.create({
  name: "belgeAutoCorrect",
  priority: 900,
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: autoCorrectKey,
        props: {
          handleTextInput(view, from, to, text) {
            if (view.composing || text.length !== 1) return false;
            const o = getAutoCorrect();
            if (o.smartQuotes && (text === '"' || text === "'")) {
              const info = textBefore(view.state, from);
              if (!info) return false;
              const q = smartQuote(info.text[info.text.length - 1], text);
              view.dispatch(view.state.tr.insertText(q, from, to));
              return true;
            }
            if (from !== to || !DELIMITER.test(text)) return false;
            const fix = pendingCorrection(view.state, from);
            if (!fix) return false;
            // Önce ayraç yazılır (yazmayla aynı adım), sonra sözcük düzeltilir: Word'deki gibi Ctrl+Z
            // yalnız düzeltmeyi geri alır, yazılan ayraç kalır.
            view.dispatch(view.state.tr.insertText(text, from, to));
            dispatchCorrection(view, fix);
            return true;
          },
          handleKeyDown(view, event) {
            if (event.ctrlKey || event.metaKey || event.altKey || view.composing) return false;
            const sel = view.state.selection;
            if (!sel.empty) return false;
            if (event.key === "Enter" && !event.shiftKey) {
              applyAutoCorrect(view, sel.from);
              return false;
            }
            return false;
          },
        },
      }),
    ];
  },
});
