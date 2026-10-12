// Şerit ve kısayolların kullandığı biçimlendirme komutları.
import type { Editor } from "@tiptap/core";
import type { Mark, Node as PMNode } from "@tiptap/pm/model";
import { TextSelection } from "@tiptap/pm/state";
import { FONT_SIZES } from "@officehere/core";
import { PARA_DEFAULTS, PARA_TYPES, type ParagraphAttrs } from "./paragraphFormat";
import { effectiveProps, styleIdOf, type StyleDef } from "../model/styles";

export type CaseMode = "sentence" | "lower" | "upper" | "title" | "toggle";

const TR = "tr-TR";

export function transformCase(text: string, mode: CaseMode, atSentenceStart = true): string {
  switch (mode) {
    case "lower":
      return text.toLocaleLowerCase(TR);
    case "upper":
      return text.toLocaleUpperCase(TR);
    case "toggle":
      return Array.from(text)
        .map((c) => {
          const u = c.toLocaleUpperCase(TR);
          return c === u ? c.toLocaleLowerCase(TR) : u;
        })
        .join("");
    case "title":
      return text.toLocaleLowerCase(TR).replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (_m, pre: string, ch: string) => pre + ch.toLocaleUpperCase(TR));
    case "sentence": {
      let start = atSentenceStart;
      let out = "";
      for (const c of text.toLocaleLowerCase(TR)) {
        if (start && /\p{L}/u.test(c)) {
          out += c.toLocaleUpperCase(TR);
          start = false;
        } else {
          out += c;
          if (/[.!?]/.test(c)) start = true;
        }
      }
      return out;
    }
  }
}

/** Seçimdeki metnin büyük/küçük harfini işaretleri koruyarak değiştirir. Seçim yoksa imleçteki sözcük. */
export function changeCase(editor: Editor, mode: CaseMode) {
  const { state, view } = editor;
  let { from, to } = state.selection;
  if (from === to) {
    const w = wordAt(state.doc, from);
    if (!w) return;
    from = w.from;
    to = w.to;
  }
  const tr = state.tr;
  const segments: { pos: number; text: string; marks: readonly Mark[] }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return true;
    const a = Math.max(from, pos);
    const b = Math.min(to, pos + node.nodeSize);
    segments.push({ pos: a, text: node.text!.slice(a - pos, b - pos), marks: node.marks });
    return false;
  });
  const $from = state.doc.resolve(from);
  const before = $from.parent.textBetween(0, $from.parentOffset);
  const sentenceStart = before.trim() === "" || /[.!?]\s*$/.test(before);
  for (let i = segments.length - 1; i >= 0; i--) {
    const s = segments[i];
    const prefixStart = i === 0 ? sentenceStart : /[.!?]\s*$/.test(segments[i - 1].text) || false;
    const t = transformCase(s.text, mode, prefixStart);
    if (t !== s.text) tr.replaceWith(s.pos, s.pos + s.text.length, state.schema.text(t, s.marks));
  }
  const end = from + (tr.doc.content.size - state.doc.content.size) + (to - from);
  tr.setSelection(TextSelection.create(tr.doc, from, Math.max(from, end)));
  view.dispatch(tr);
}

/** Shift+F3: BÜYÜK → küçük → Her Sözcük döngüsü. */
export function cycleCase(editor: Editor) {
  const { from, to } = editor.state.selection;
  const text = editor.state.doc.textBetween(from, to, " ") || "";
  let next: CaseMode = "upper";
  if (text && text === text.toLocaleUpperCase(TR)) next = "lower";
  else if (text && text === text.toLocaleLowerCase(TR)) next = "title";
  changeCase(editor, next);
}

export function wordAt(doc: PMNode, pos: number): { from: number; to: number } | null {
  const $p = doc.resolve(pos);
  if (!$p.parent.isTextblock) return null;
  const text = $p.parent.textContent;
  const off = $p.parentOffset;
  const isW = (c: string | undefined) => !!c && /[\p{L}\p{N}_'’]/u.test(c);
  let a = off;
  let b = off;
  while (a > 0 && isW(text[a - 1])) a--;
  while (b < text.length && isW(text[b])) b++;
  if (a === b) return null;
  const start = $p.start();
  return { from: start + a, to: start + b };
}

/* ------------------------------------------------------------------ */
/* Paragraf                                                              */
/* ------------------------------------------------------------------ */

/** Seçimdeki ilk paragrafın öznitelikleri. */
export function currentPara(editor: Editor): { node: PMNode; pos: number } | null {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d >= 0; d--) {
    const n = $from.node(d);
    if (PARA_TYPES.includes(n.type.name)) return { node: n, pos: d > 0 ? $from.before(d) : 0 };
  }
  // Düğüm seçimi (ör. resim) — üst paragraf
  let found: { node: PMNode; pos: number } | null = null;
  editor.state.doc.nodesBetween(editor.state.selection.from, editor.state.selection.to, (n, pos) => {
    if (found) return false;
    if (PARA_TYPES.includes(n.type.name)) {
      found = { node: n, pos };
      return false;
    }
    return true;
  });
  return found;
}

export function setParaAttrs(editor: Editor, attrs: Partial<ParagraphAttrs>) {
  editor.chain().focus().updateAttributes("paragraph", attrs).updateAttributes("heading", attrs).run();
}

/** Paragraf stilini uygular (Başlık n → heading düğümü). */
export function applyStyle(editor: Editor, styles: Record<string, StyleDef>, id: string) {
  const def = styles[id];
  if (!def) return;
  if (def.type === "character") {
    const active = editor.isActive("charStyle", { styleId: id });
    if (active) editor.chain().focus().unsetCharStyle().run();
    else editor.chain().focus().setCharStyle(id).run();
    return;
  }
  const keep = (n: PMNode) => {
    const a = { ...n.attrs } as Record<string, unknown>;
    delete a.level;
    a.styleId = null;
    return a;
  };
  const { state } = editor;
  const tr = state.tr;
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (!PARA_TYPES.includes(node.type.name)) return true;
    const base = keep(node);
    if (def.level) tr.setNodeMarkup(pos, state.schema.nodes.heading, { ...base, level: def.level });
    else tr.setNodeMarkup(pos, state.schema.nodes.paragraph, { ...base, styleId: id === "Normal" ? null : id });
    return false;
  });
  editor.view.dispatch(tr);
  editor.commands.focus();
}

/** Düzey girintisi: 1,25 cm. */
export const INDENT_STEP = 35.4;

export function changeIndent(editor: Editor, dir: 1 | -1) {
  if (editor.isActive("listItem")) {
    if (dir > 0) editor.chain().focus().sinkListItem("listItem").run();
    else editor.chain().focus().liftListItem("listItem").run();
    return;
  }
  const { state } = editor;
  const tr = state.tr;
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (!PARA_TYPES.includes(node.type.name)) return true;
    const cur = (node.attrs.indentLeft as number | null) ?? 0;
    const next = dir > 0 ? Math.floor(cur / INDENT_STEP + 1.001) * INDENT_STEP : Math.max(0, Math.ceil(cur / INDENT_STEP - 1.001) * INDENT_STEP);
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: next ? +next.toFixed(2) : null });
    return false;
  });
  editor.view.dispatch(tr);
  editor.commands.focus();
}

export function setLineSpacing(editor: Editor, value: number) {
  setParaAttrs(editor, { lineRule: "auto", lineHeight: value });
}

/** Paragrafları alfabetik sıralar (Türkçe). */
export function sortParagraphs(editor: Editor, desc: boolean, numeric = false) {
  const { state } = editor;
  const { from, to } = state.selection;
  const $a = state.doc.resolve(from);
  const $b = state.doc.resolve(to);
  const range = $a.blockRange($b);
  if (!range) return;
  const nodes: PMNode[] = [];
  for (let i = range.startIndex; i < range.endIndex; i++) nodes.push(range.parent.child(i));
  if (nodes.length < 2) return;
  const key = (n: PMNode) => n.textContent.trim();
  const sorted = [...nodes].sort((x, y) => {
    const r = numeric ? (parseFloat(key(x).replace(",", ".")) || 0) - (parseFloat(key(y).replace(",", ".")) || 0) : key(x).localeCompare(key(y), "tr", { sensitivity: "base" });
    return desc ? -r : r;
  });
  const tr = state.tr.replaceWith(range.start, range.end, sorted);
  editor.view.dispatch(tr);
}

/* ------------------------------------------------------------------ */
/* Yazı tipi                                                             */
/* ------------------------------------------------------------------ */

/** İmleçteki etkin yazı tipi boyutu (pt). */
export function effectiveFontSize(editor: Editor, styles: Record<string, StyleDef>): number {
  const fs = editor.getAttributes("textStyle").fontSize as string | undefined;
  if (fs) return parseFloat(fs);
  const p = currentPara(editor);
  const id = p ? styleIdOf(p.node.type.name, p.node.attrs) : "Normal";
  return effectiveProps(styles, id).size ?? 11;
}

export function effectiveFont(editor: Editor, styles: Record<string, StyleDef>): string {
  const ff = editor.getAttributes("textStyle").fontFamily as string | undefined;
  if (ff) return ff.replace(/^["']|["']$/g, "").split(",")[0].replace(/["']/g, "").trim();
  const p = currentPara(editor);
  const id = p ? styleIdOf(p.node.type.name, p.node.attrs) : "Normal";
  return effectiveProps(styles, id).font ?? "Aptos";
}

export function setFontSize(editor: Editor, pt: number) {
  const v = Math.max(1, Math.min(1638, Math.round(pt * 2) / 2));
  editor.chain().focus().setFontSize(`${v}pt`).run();
}

/** Ctrl+Shift+> / < : listedeki sonraki boyut; Ctrl+] / [ : 1 nk. */
export function stepFontSize(editor: Editor, styles: Record<string, StyleDef>, dir: 1 | -1, fine = false) {
  const cur = effectiveFontSize(editor, styles);
  let next: number;
  if (fine) next = cur + dir;
  else if (dir > 0) next = FONT_SIZES.find((s) => s > cur + 0.01) ?? Math.ceil((cur + 10) / 10) * 10;
  else next = [...FONT_SIZES].reverse().find((s) => s < cur - 0.01) ?? Math.max(1, cur - 1);
  setFontSize(editor, next);
}

/** Tüm Biçimlendirmeyi Temizle: karakter biçimi ve paragraf stili Normal'e döner. */
export function clearFormatting(editor: Editor) {
  const { state } = editor;
  const { from, to } = state.selection;
  const keep = new Set(["link", "comment", "insertion", "deletion", "highlight"]);
  const tr = state.tr;
  for (const mt of Object.values(state.schema.marks)) if (!keep.has(mt.name)) tr.removeMark(from, to, mt);
  if (from === to) tr.setStoredMarks([]);
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!PARA_TYPES.includes(node.type.name)) return true;
    tr.setNodeMarkup(pos, state.schema.nodes.paragraph, { ...PARA_DEFAULTS, textAlign: null });
    return false;
  });
  editor.view.dispatch(tr);
  editor.commands.focus();
}

/** Ctrl+Space: yalnız karakter biçimini temizler. */
export function clearCharFormatting(editor: Editor) {
  const { state } = editor;
  const { from, to } = state.selection;
  const keep = new Set(["link", "comment", "insertion", "deletion"]);
  const tr = state.tr;
  for (const mt of Object.values(state.schema.marks)) if (!keep.has(mt.name)) tr.removeMark(from, to, mt);
  if (from === to) tr.setStoredMarks([]);
  editor.view.dispatch(tr);
}

/* ------------------------------------------------------------------ */
/* Biçim Boyacısı                                                        */
/* ------------------------------------------------------------------ */

export interface CopiedFormat {
  marks: Mark[];
  para: { type: string; attrs: Record<string, unknown> } | null;
}

export function copyFormat(editor: Editor): CopiedFormat {
  const { state } = editor;
  const { $from, empty } = state.selection;
  const marks = empty ? (state.storedMarks ?? $from.marks()) : state.doc.resolve(Math.min(state.selection.from + 1, state.selection.to)).marks();
  const p = currentPara(editor);
  const skip = new Set(["link", "comment", "insertion", "deletion"]);
  return { marks: marks.filter((m) => !skip.has(m.type.name)), para: p ? { type: p.node.type.name, attrs: { ...p.node.attrs } } : null };
}

export function pasteFormat(editor: Editor, f: CopiedFormat) {
  const { state } = editor;
  let { from, to } = state.selection;
  if (from === to) {
    const w = wordAt(state.doc, from);
    if (w) ({ from, to } = w);
  }
  const tr = state.tr;
  const skip = new Set(["link", "comment", "insertion", "deletion"]);
  for (const mt of Object.values(state.schema.marks)) if (!skip.has(mt.name)) tr.removeMark(from, to, mt);
  for (const m of f.marks) tr.addMark(from, to, m);
  // Paragraf biçimi: seçim tüm paragrafı kapsıyorsa ya da birden çok paragraf varsa.
  const $a = state.doc.resolve(from);
  const $b = state.doc.resolve(to);
  const whole = $a.parent !== $b.parent || ($a.parentOffset === 0 && $b.parentOffset === $b.parent.content.size);
  if (f.para && whole) {
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!PARA_TYPES.includes(node.type.name)) return true;
      tr.setNodeMarkup(pos, state.schema.nodes[f.para!.type], f.para!.attrs);
      return false;
    });
  }
  editor.view.dispatch(tr);
}

/* ------------------------------------------------------------------ */
/* Ekle                                                                  */
/* ------------------------------------------------------------------ */

export const DATE_FORMATS: { id: string; fmt: (d: Date) => string }[] = [
  { id: "short", fmt: (d) => d.toLocaleDateString("tr-TR") },
  { id: "long", fmt: (d) => d.toLocaleDateString("tr-TR", { day: "numeric", month: "long", year: "numeric" }) },
  { id: "weekday", fmt: (d) => d.toLocaleDateString("tr-TR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }) },
  { id: "iso", fmt: (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` },
  { id: "monthYear", fmt: (d) => d.toLocaleDateString("tr-TR", { month: "long", year: "numeric" }) },
  { id: "shortTime", fmt: (d) => d.toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }) },
  { id: "time", fmt: (d) => d.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" }) },
  { id: "timeSec", fmt: (d) => d.toLocaleTimeString("tr-TR") },
];

export function insertText(editor: Editor, text: string) {
  editor.chain().focus().insertContent(text).run();
}

/** Belgedeki başlıklar (gezinti bölmesi ve içindekiler için). */
export function collectHeadings(doc: PMNode): { level: number; text: string; pos: number }[] {
  const out: { level: number; text: string; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "heading") {
      const text = node.textContent.trim();
      if (text) out.push({ level: node.attrs.level as number, text, pos });
      return false;
    }
    return node.type.name !== "paragraph";
  });
  return out;
}

/** Sözcük sayımı (Türkçe harfleri ve kesme işaretli sözcükleri tanır). */
export function countWords(text: string): number {
  const m = text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-.,]*/gu);
  return m ? m.length : 0;
}

export function docStats(doc: PMNode) {
  const text = doc.textBetween(0, doc.content.size, "\n", " ");
  let paragraphs = 0;
  doc.descendants((n) => {
    if (n.isTextblock) {
      if (n.textContent.trim()) paragraphs++;
      return false;
    }
    return true;
  });
  return {
    words: countWords(text),
    chars: text.replace(/\n/g, "").length,
    charsNoSpaces: text.replace(/\s/g, "").length,
    paragraphs,
  };
}
