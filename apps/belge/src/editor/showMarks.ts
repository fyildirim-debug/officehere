// Tümünü Göster (¶): paragraf sonu, boşluk, sekme ve satır sonu işaretleri.
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

interface MarksState {
  on: boolean;
  deco: DecorationSet;
}

export const showMarksKey = new PluginKey<MarksState>("belgeShowMarks");

function pilcrow() {
  const s = document.createElement("span");
  s.className = "bl-pilcrow";
  s.textContent = "¶";
  return s;
}

function lineBreakMark() {
  const s = document.createElement("span");
  s.className = "bl-pilcrow";
  s.textContent = "↵";
  return s;
}

function build(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      decos.push(Decoration.widget(pos + node.nodeSize - 1, pilcrow, { side: 1, key: "pil", ignoreSelection: true }));
      return true;
    }
    if (node.isText) {
      const t = node.text ?? "";
      for (let i = 0; i < t.length; i++) {
        const ch = t[i];
        if (ch === " " || ch === " ") decos.push(Decoration.inline(pos + i, pos + i + 1, { class: "bl-ws" }));
        else if (ch === "\t") decos.push(Decoration.inline(pos + i, pos + i + 1, { class: "bl-tab" }));
      }
      return false;
    }
    if (node.type.name === "hardBreak") decos.push(Decoration.widget(pos, lineBreakMark, { side: -1, key: "br" }));
    return true;
  });
  return DecorationSet.create(doc, decos);
}

export const ShowMarks = Extension.create({
  name: "belgeShowMarks",
  addProseMirrorPlugins() {
    return [
      new Plugin<MarksState>({
        key: showMarksKey,
        state: {
          init: () => ({ on: false, deco: DecorationSet.empty }),
          apply(tr, value, _old, state) {
            const meta = tr.getMeta(showMarksKey) as boolean | undefined;
            if (meta !== undefined) return { on: meta, deco: meta ? build(state.doc) : DecorationSet.empty };
            if (value.on && tr.docChanged) return { on: true, deco: build(state.doc) };
            return value;
          },
        },
        props: {
          decorations: (state) => showMarksKey.getState(state)?.deco,
          attributes: (state): Record<string, string> => (showMarksKey.getState(state)?.on ? { class: "bl-show-marks" } : {}),
        },
      }),
    ];
  },
});

export function setShowMarks(view: EditorView, on: boolean) {
  view.dispatch(view.state.tr.setMeta(showMarksKey, on).setMeta("addToHistory", false));
}
