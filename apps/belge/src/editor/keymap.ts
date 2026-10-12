// Kelime işlemci kısayolları (editör odaklıyken).
import { Extension, type Editor } from "@tiptap/core";
import type { StyleDef } from "../model/styles";
import { applyStyle, changeIndent, clearCharFormatting, copyFormat, cycleCase, pasteFormat, setLineSpacing, setParaAttrs, stepFontSize, currentPara, type CopiedFormat } from "./commands";

export interface EditorHooks {
  styles: () => Record<string, StyleDef>;
  openDialog: (kind: "font" | "paragraph" | "link") => void;
  toggleShowMarks: () => void;
  newComment: () => void;
  toggleTrack: () => void;
  insertNote: (kind: "footnote" | "endnote") => void;
}

let copied: CopiedFormat | null = null;

/** Sonraki paragrafı Normal olan stiller (Word'deki "Sonraki paragraf stili"). */
const NEXT_NORMAL = new Set(["Title", "Subtitle", "Quote", "IntenseQuote"]);

export function createWordKeymap(hooks: EditorHooks) {
  return Extension.create({
    name: "wordKeymap",
    priority: 1000,
    addKeyboardShortcuts() {
      const ed = () => this.editor as Editor;
      const align = (a: string) => () => ed().chain().focus().setTextAlign(a).run();
      const underlineType = (type: "double") => () => {
        const cur = ed().getAttributes("underline");
        if (ed().isActive("underline") && cur.type === type) return ed().chain().focus().unsetUnderline().run();
        return ed().chain().focus().setUnderlineType({ type }).run();
      };
      const toggleTs = (attr: "caps" | "smallCaps") => () => {
        const on = !!ed().getAttributes("textStyle")[attr];
        return ed()
          .chain()
          .focus()
          .setMark("textStyle", { [attr]: !on })
          .run();
      };
      return {
        "Mod-u": () => ed().chain().focus().toggleUnderline().run(),
        "Mod-Shift-d": underlineType("double"),
        "Mod-Shift-a": toggleTs("caps"),
        "Mod-Shift-k": toggleTs("smallCaps"),
        "Mod-e": align("center"),
        "Mod-l": align("left"),
        "Mod-r": align("right"),
        "Mod-j": align("justify"),
        "Mod-1": () => (setLineSpacing(ed(), 1), true),
        "Mod-2": () => (setLineSpacing(ed(), 2), true),
        "Mod-5": () => (setLineSpacing(ed(), 1.5), true),
        "Mod-0": () => {
          const p = currentPara(ed());
          const cur = (p?.node.attrs.spaceBefore as number | null) ?? 0;
          setParaAttrs(ed(), { spaceBefore: cur >= 12 ? 0 : 12 });
          return true;
        },
        "Mod-Shift->": () => (stepFontSize(ed(), hooks.styles(), 1), true),
        "Mod-Shift-<": () => (stepFontSize(ed(), hooks.styles(), -1), true),
        "Mod-]": () => (stepFontSize(ed(), hooks.styles(), 1, true), true),
        "Mod-[": () => (stepFontSize(ed(), hooks.styles(), -1, true), true),
        "Mod-Alt-1": () => (applyStyle(ed(), hooks.styles(), "Heading1"), true),
        "Mod-Alt-2": () => (applyStyle(ed(), hooks.styles(), "Heading2"), true),
        "Mod-Alt-3": () => (applyStyle(ed(), hooks.styles(), "Heading3"), true),
        "Mod-Shift-n": () => (applyStyle(ed(), hooks.styles(), "Normal"), true),
        "Mod-Space": () => (clearCharFormatting(ed()), true),
        "Mod-Enter": () => ed().chain().focus().setPageBreak().run(),
        "Shift-F3": () => (cycleCase(ed()), true),
        "Mod-Shift-=": () => ed().chain().focus().toggleSuperscript().run(),
        "Mod-Shift-+": () => ed().chain().focus().toggleSuperscript().run(),
        "Mod-=": () => ed().chain().focus().toggleSubscript().run(),
        "Mod-d": () => (hooks.openDialog("font"), true),
        "Mod-k": () => (hooks.openDialog("link"), true),
        "Mod-Shift-8": () => (hooks.toggleShowMarks(), true),
        "Mod-*": () => (hooks.toggleShowMarks(), true),
        "Mod-Shift-l": () => ed().chain().focus().toggleBulletList().run(),
        "Mod-m": () => (changeIndent(ed(), 1), true),
        "Mod-Shift-m": () => (changeIndent(ed(), -1), true),
        "Mod-t": () => {
          const p = currentPara(ed());
          const left = ((p?.node.attrs.indentLeft as number | null) ?? 0) + 35.4;
          setParaAttrs(ed(), { indentLeft: +left.toFixed(2), indentFirst: -35.4 });
          return true;
        },
        "Mod-Shift-c": () => {
          copied = copyFormat(ed());
          return true;
        },
        "Mod-Shift-v": () => {
          if (copied) pasteFormat(ed(), copied);
          return true;
        },
        "Mod-Alt-m": () => (hooks.newComment(), true),
        "Mod-Shift-e": () => (hooks.toggleTrack(), true),
        "Mod-Alt-f": () => (hooks.insertNote("footnote"), true),
        "Mod-Alt-d": () => (hooks.insertNote("endnote"), true),
        Enter: () => {
          const e = ed();
          const { $from, empty } = e.state.selection;
          const node = $from.parent;
          const sid = node.attrs.styleId as string | null;
          if (!empty || node.type.name !== "paragraph" || !sid || !NEXT_NORMAL.has(sid) || $from.parentOffset !== node.content.size) return false;
          return e.chain().splitBlock().updateAttributes("paragraph", { styleId: null }).run();
        },
        Tab: () => {
          const e = ed();
          if (e.isActive("listItem") || e.isActive("table")) return false;
          return e.chain().focus().insertContent("\t").run();
        },
      };
    },
  });
}
