// Başvurular yardımcıları: İçindekiler / Şekiller Tablosu güncelleme, resim yazısı ekleme, çapraz
// başvuru hedefleri (resim yazıları, başlıklar, yer işaretleri) ve yer işareti oluşturma.
import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { collectHeadings } from "./commands";
import { paraText, type CrossRefKind } from "./fields";

export const CAPTION_LABELS = ["Şekil", "Tablo", "Denklem"];

export interface CaptionInfo {
  label: string;
  num: number;
  text: string;
  /** Resim yazısı paragrafının konumu. */
  pos: number;
}

/** Belgedeki resim yazıları (etiket filtresiyle). */
export function collectCaptions(doc: PMNode, label?: string): CaptionInfo[] {
  const out: CaptionInfo[] = [];
  doc.descendants((n, pos) => {
    if (!n.isTextblock) return true;
    let seq: PMNode | null = null;
    n.descendants((c) => {
      if (!seq && c.type.name === "seqField") seq = c;
      return !seq;
    });
    const s = seq as PMNode | null;
    if (s && (!label || s.attrs.label === label)) out.push({ label: s.attrs.label as string, num: s.attrs.num as number, text: paraText(n), pos });
    return false;
  });
  return out;
}

/** İçindekiler ve Şekiller Tablosu düğümlerini günceller. */
export function tocUpdate(state: EditorState, pageOf: (pos: number) => number): Transaction | null {
  const tr = state.tr;
  let found = false;
  const headings = collectHeadings(state.doc).filter((h) => h.level <= 3);
  state.doc.descendants((n, pos) => {
    if (n.type.name !== "toc") return true;
    const entries =
      n.attrs.kind === "tof"
        ? collectCaptions(state.doc, (n.attrs.label as string) || undefined).map((c) => ({ level: 1, text: c.text, page: pageOf(c.pos) + 1, pos: c.pos }))
        : headings.map((h) => ({ level: h.level, text: h.text, page: pageOf(h.pos) + 1, pos: h.pos }));
    tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries });
    found = true;
    return false;
  });
  return found ? tr : null;
}

/**
 * Resim yazısı ekler: seçili resmin/tablonun altına ya da üstüne (yoksa imleçteki paragraftan sonra)
 * "Resim Yazısı" stilinde "Etiket N metin" paragrafı.
 */
export function insertCaption(editor: Editor, label: string, text: string, position: "above" | "below", excludeLabel: boolean) {
  const { state } = editor;
  const sel = state.selection;
  const schema = state.schema;
  let at: number;
  const $f = sel.$from;
  const tableDepth = (() => {
    for (let d = $f.depth; d > 0; d--) if ($f.node(d).type.name === "table") return d;
    return -1;
  })();
  if (tableDepth > 0) at = position === "above" ? $f.before(tableDepth) : $f.after(tableDepth);
  else {
    // Resim (satır içi) ya da imleç: bulunduğu üst düzey blok.
    const pos = sel instanceof NodeSelection ? sel.from : sel.from;
    const $p = state.doc.resolve(pos);
    let d = $p.depth;
    while (d > 2 && !$p.node(d).isTextblock) d--;
    if ($p.depth < 1) return;
    const blockDepth = Math.min(Math.max(1, d), $p.depth);
    at = position === "above" ? $p.before(blockDepth) : $p.after(blockDepth);
  }
  const seq = schema.nodes.seqField.create({ label });
  const parts = [...(excludeLabel ? [] : [schema.text(`${label} `)]), seq, ...(text ? [schema.text(text.startsWith(":") || text.startsWith(".") ? text : ` ${text}`)] : [])];
  const para = schema.nodes.paragraph.create({ styleId: "Caption" }, parts);
  const tr = state.tr.insert(at, para);
  tr.setSelection(TextSelection.create(tr.doc, at + para.nodeSize - 1));
  editor.view.dispatch(tr.scrollIntoView());
  editor.view.focus();
}

export interface RefTarget {
  /** Listede gösterilen metin. */
  label: string;
  /** Başvurulacak düğümün (paragraf) konumu. */
  pos: number;
  /** Varsa paragraftaki yer işareti adı. */
  bookmark: string | null;
}

/** Çapraz başvuru türüne göre hedefler. */
export function refTargets(doc: PMNode, type: string): RefTarget[] {
  const bookmarkIn = (p: PMNode) => {
    let name: string | null = null;
    p.descendants((c) => {
      if (!name && c.type.name === "bookmark" && /^_Ref/.test(c.attrs.name as string)) name = c.attrs.name as string;
      return !name;
    });
    return name;
  };
  if (type === "Başlık") {
    const out: RefTarget[] = [];
    doc.descendants((n, pos) => {
      if (n.type.name === "heading") {
        const t = n.textContent.trim();
        if (t) out.push({ label: `${"  ".repeat((n.attrs.level as number) - 1)}${t}`, pos, bookmark: bookmarkIn(n) });
        return false;
      }
      return !n.isTextblock;
    });
    return out;
  }
  if (type === "Yer İşareti") {
    const out: RefTarget[] = [];
    doc.descendants((n, pos) => {
      if (n.type.name === "bookmark" && !/^_/.test(n.attrs.name as string)) out.push({ label: n.attrs.name as string, pos: doc.resolve(pos).before(), bookmark: n.attrs.name as string });
      return true;
    });
    return out;
  }
  return collectCaptions(doc, type).map((c) => ({ label: c.text, pos: c.pos, bookmark: bookmarkIn(doc.nodeAt(c.pos)!) }));
}

/** Hedef paragrafın başına gizli "_Ref" yer işareti koyar (yoksa) ve adını döner. */
export function ensureRefBookmark(editor: Editor, target: RefTarget): string {
  if (target.bookmark) return target.bookmark;
  const name = `_Ref${Date.now().toString().slice(-9)}${Math.floor(Math.random() * 90 + 10)}`;
  const { state } = editor;
  editor.view.dispatch(state.tr.insert(target.pos + 1, state.schema.nodes.bookmark.create({ name })).setMeta("addToHistory", true));
  return name;
}

/** İmlece çapraz başvuru ekler. */
export function insertCrossRef(editor: Editor, target: RefTarget, kind: CrossRefKind, link: boolean) {
  const sel = editor.state.selection;
  const from = sel.from;
  const name = ensureRefBookmark(editor, target);
  // Yer işareti eklendiyse imleç konumu kaymış olabilir.
  const at = target.bookmark || target.pos + 1 > from ? from : from + 1;
  const { state } = editor;
  const node = state.schema.nodes.crossRef.create({ ref: name, kind, link });
  const tr = state.tr.insert(Math.min(at, state.doc.content.size), node);
  tr.setSelection(TextSelection.create(tr.doc, Math.min(at + 1, tr.doc.content.size)));
  editor.view.dispatch(tr);
  editor.view.focus();
}
