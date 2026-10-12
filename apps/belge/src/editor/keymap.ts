// Kelime işlemci kısayolları (editör odaklıyken).
import { Extension, type Editor } from "@tiptap/core";
import { Selection, TextSelection } from "@tiptap/pm/state";
import { CellSelection, TableMap, nextCell } from "@tiptap/pm/tables";
import type { StyleDef } from "../model/styles";
import { toggleList } from "./listApply";
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

/** Ctrl+Home / Ctrl+End (Shift ile seçimi genişletir). */
function docEdge(editor: Editor, dir: -1 | 1, extend: boolean): boolean {
  const { state, view } = editor;
  const target = dir < 0 ? Selection.atStart(state.doc) : Selection.atEnd(state.doc);
  let sel: Selection = target;
  if (extend) {
    const head = dir < 0 ? target.from : target.to;
    sel = TextSelection.between(state.selection.$anchor, state.doc.resolve(head));
  }
  view.dispatch(state.tr.setSelection(sel).scrollIntoView());
  return true;
}

/** Bölüm başındaki tablonun ilk hücresinin başında Enter: tablonun önüne paragraf ekler (Word davranışı). */
function paragraphBeforeTable(editor: Editor): boolean {
  const { state, view } = editor;
  const { $from, empty } = state.selection;
  // belge > bölüm > tablo > satır > hücre > paragraf
  if (!empty || $from.parentOffset !== 0 || $from.depth < 5) return false;
  if ($from.node(2).type.name !== "table" || $from.index(1) !== 0) return false;
  // Bölümün ilk bloğu olan tablonun ilk satırı, ilk hücresi, hücrenin ilk bloğu
  if ($from.index(2) !== 0 || $from.index(3) !== 0 || $from.index(4) !== 0) return false;
  const at = $from.before(2);
  const tr = state.tr.insert(at, state.schema.nodes.paragraph.create());
  tr.setSelection(TextSelection.create(tr.doc, at + 1));
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** Sonraki paragrafı Normal olan stiller (Word'deki "Sonraki paragraf stili"). */
/**
 * Hücre seçimi satır sonunda/başında Shift+Sağ/Sol ile sonraki/önceki satıra uzar (Word: tüm satırlar seçilir).
 * Tablo eklentisi bu durumda olayı tarayıcıya bırakır ve seçim çöker.
 */
function extendCellRows(editor: Editor, dir: -1 | 1): boolean {
  const { state, view } = editor;
  const sel = state.selection;
  if (!(sel instanceof CellSelection)) return false;
  const $head = sel.$headCell;
  if (nextCell($head, "horiz", dir)) return false;
  const table = $head.node(-1);
  const start = $head.start(-1);
  const map = TableMap.get(table);
  const rect = map.findCell($head.pos - start);
  const row = dir > 0 ? rect.bottom : rect.top - 1;
  if (row < 0 || row >= map.height) return true;
  const cell = start + map.map[row * map.width + (dir > 0 ? 0 : map.width - 1)];
  view.dispatch(state.tr.setSelection(CellSelection.rowSelection(sel.$anchorCell, state.doc.resolve(cell))));
  return true;
}

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
        "Mod-Shift-l": () => toggleList(ed(), "bulletList"),
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
          if (paragraphBeforeTable(e)) return true;
          const { $from, empty } = e.state.selection;
          const node = $from.parent;
          const sid = node.attrs.styleId as string | null;
          if (!empty || node.type.name !== "paragraph" || !sid || !NEXT_NORMAL.has(sid) || $from.parentOffset !== node.content.size) return false;
          return e.chain().splitBlock().updateAttributes("paragraph", { styleId: null }).run();
        },
        // Belge başı/sonu: tarayıcının yerel davranışı resim, tablo ve hücre seçiminde güvenilmez.
        "Mod-Home": () => docEdge(ed(), -1, false),
        "Mod-End": () => docEdge(ed(), 1, false),
        "Mod-Shift-Home": () => docEdge(ed(), -1, true),
        "Mod-Shift-End": () => docEdge(ed(), 1, true),
        "Shift-ArrowRight": () => extendCellRows(ed(), 1),
        "Shift-ArrowLeft": () => extendCellRows(ed(), -1),
        Tab: () => {
          const e = ed();
          if (e.isActive("listItem") || e.isActive("table")) return false;
          return e.chain().focus().insertContent("\t").run();
        },
      };
    },
  });
}
