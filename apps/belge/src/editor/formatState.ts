// Şeridin gösterdiği biçim durumu (imleç/seçimdeki etkin biçim).
import type { Editor } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import { CellSelection } from "@tiptap/pm/tables";
import { styleIdOf, type StyleDef } from "../model/styles";
import { currentPara, effectiveFont, effectiveFontSize } from "./commands";
import { isTracking } from "./trackChanges";
import { repeatKey } from "./repeat";

export interface FormatState {
  bold: boolean;
  italic: boolean;
  underline: string | null;
  strike: boolean;
  sub: boolean;
  sup: boolean;
  font: string;
  size: number;
  color: string | null;
  highlight: string | null;
  align: string;
  styleId: string;
  charStyle: string | null;
  list: "bullet" | "ordered" | null;
  listStyle: string | null;
  inTable: boolean;
  cellSelection: boolean;
  image: boolean;
  /** Metin kutusu seçili ya da imleç metin kutusunun içinde. */
  textBox: boolean;
  link: string | null;
  canUndo: boolean;
  canRedo: boolean;
  indentLeft: number;
  indentRight: number;
  indentFirst: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  lineHeight: number | null;
  lineRule: string | null;
  empty: boolean;
  tracking: boolean;
  caps: boolean;
  smallCaps: boolean;
  /** İmleçteki bölüm, bölüm sayısı ve bölümün kendi sayfa yapısının anahtarı (değişimi izlemek için). */
  section: number;
  sectionCount: number;
  sectionKey: string;
  /** Satır numaraları bu paragrafta gizli. */
  suppressLineNumbers: boolean;
}

export function computeFormat(editor: Editor, styles: Record<string, StyleDef>): FormatState {
  const { state } = editor;
  const sel = state.selection;
  const p = currentPara(editor);
  const pa = p?.node.attrs ?? {};
  const ts = editor.getAttributes("textStyle");
  const inBullet = editor.isActive("bulletList");
  const inOrdered = editor.isActive("orderedList");
  let listStyle: string | null = null;
  if (inBullet || inOrdered) {
    const { $from } = sel;
    for (let d = $from.depth; d > 0; d--) {
      const n = $from.node(d);
      if (n.type.name === "bulletList" || n.type.name === "orderedList") {
        listStyle = (n.attrs.listStyle as string | null) ?? (n.type.name === "bulletList" ? "disc" : "decimal");
        break;
      }
    }
  }
  const image = sel instanceof NodeSelection && sel.node.type.name === "image";
  const textBox = textBoxAt(editor.state) !== null;
  const deepest = (() => {
    const { $from } = sel;
    for (let d = $from.depth; d > 0; d--) {
      const n = $from.node(d).type.name;
      if (n === "bulletList") return "bullet" as const;
      if (n === "orderedList") return "ordered" as const;
    }
    return null;
  })();
  return {
    bold: editor.isActive("bold"),
    italic: editor.isActive("italic"),
    underline: editor.isActive("underline") ? ((editor.getAttributes("underline").type as string) ?? "single") : null,
    strike: editor.isActive("strike"),
    sub: editor.isActive("subscript"),
    sup: editor.isActive("superscript"),
    font: effectiveFont(editor, styles),
    size: effectiveFontSize(editor, styles),
    color: (ts.color as string | undefined) ?? null,
    highlight: editor.isActive("highlight") ? ((editor.getAttributes("highlight").color as string | undefined) ?? "#FFFF00") : null,
    align: (pa.textAlign as string | null) ?? "left",
    styleId: p ? styleIdOf(p.node.type.name, p.node.attrs) : "Normal",
    charStyle: editor.isActive("charStyle") ? (editor.getAttributes("charStyle").styleId as string) : null,
    list: deepest,
    listStyle,
    inTable: editor.isActive("table"),
    cellSelection: sel instanceof CellSelection,
    image,
    textBox,
    link: editor.isActive("link") ? ((editor.getAttributes("link").href as string) ?? "") : null,
    canUndo: editor.can().undo(),
    // Yinele: geri alınmış adım ya da yinelenebilecek son eylem varsa etkin (Word'deki "Yinele").
    canRedo: editor.can().redo() || !!repeatKey.getState(state),
    indentLeft: (pa.indentLeft as number | null) ?? 0,
    indentRight: (pa.indentRight as number | null) ?? 0,
    indentFirst: (pa.indentFirst as number | null) ?? 0,
    spaceBefore: (pa.spaceBefore as number | null) ?? null,
    spaceAfter: (pa.spaceAfter as number | null) ?? null,
    suppressLineNumbers: !!pa.suppressLineNumbers,
    lineHeight: (pa.lineHeight as number | null) ?? null,
    lineRule: (pa.lineRule as string | null) ?? null,
    empty: sel.empty,
    tracking: isTracking(state),
    caps: !!ts.caps,
    smallCaps: !!ts.smallCaps,
    section: sel.$from.depth >= 1 ? sel.$from.index(0) : 0,
    sectionCount: state.doc.childCount,
    sectionKey: (() => {
      const sec = state.doc.maybeChild(sel.$from.depth >= 1 ? sel.$from.index(0) : 0);
      return sec ? `${sec.attrs.kind}|${sec.attrs.page ? JSON.stringify(sec.attrs.page) : ""}` : "";
    })(),
  };
}

export function sameFormat(a: FormatState, b: FormatState): boolean {
  for (const k of Object.keys(a) as (keyof FormatState)[]) if (a[k] !== b[k]) return false;
  return true;
}

/** Seçili ya da imleci içeren metin kutusu (konum ve düğüm). */
export function textBoxAt(state: Editor["state"]): { pos: number; node: Editor["state"]["doc"] } | null {
  const sel = state.selection;
  if (sel instanceof NodeSelection && sel.node.type.name === "textBox") return { pos: sel.from, node: sel.node };
  const { $from } = sel;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.name === "textBox") return { pos: $from.before(d), node: $from.node(d) };
  return null;
}
