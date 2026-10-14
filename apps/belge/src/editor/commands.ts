// Şerit ve kısayolların kullandığı biçimlendirme komutları.
import type { Editor } from "@tiptap/core";
import type { Mark, Node as PMNode } from "@tiptap/pm/model";
import { Selection, TextSelection } from "@tiptap/pm/state";
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
  // Seçim (Tümünü Seç dahil) aynı türüyle korunur; metin boyu değişse de eşleme ile kayar.
  const end = from + (tr.doc.content.size - state.doc.content.size) + (to - from);
  if (state.selection.empty) tr.setSelection(TextSelection.between(tr.doc.resolve(from), tr.doc.resolve(Math.max(from, end))));
  else tr.setSelection(state.selection.map(tr.doc, tr.mapping));
  view.dispatch(tr);
}

/** Shift+F3 döngüsündeki sonraki kip: BÜYÜK → küçük → Her Sözcük (seçim tam cümle içeriyorsa Cümle düzeni). */
export function nextCaseMode(text: string): CaseMode {
  if (text && text === text.toLocaleUpperCase(TR) && text !== text.toLocaleLowerCase(TR)) return "lower";
  if (text && text === text.toLocaleLowerCase(TR)) return /[.!?…](\s|$)/.test(text.trim()) ? "sentence" : "title";
  return "upper";
}

/** Shift+F3: BÜYÜK → küçük → Her Sözcük / Cümle düzeni döngüsü (seçim yoksa imleçteki sözcük). */
export function cycleCase(editor: Editor) {
  const { from, to, empty } = editor.state.selection;
  let text = editor.state.doc.textBetween(from, to, " ") || "";
  if (empty) {
    const w = wordAt(editor.state.doc, from);
    if (w) text = editor.state.doc.textBetween(w.from, w.to);
  }
  changeCase(editor, nextCaseMode(text));
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
  // Tümünü Seç bölüm düzeyine çıkar; seçim metin bloklarına daraltılır.
  const from = Math.max(state.selection.from, Selection.atStart(state.doc).from);
  const to = Math.min(state.selection.to, Selection.atEnd(state.doc).to);
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
  const keep = new Set(["link", "comment", "insertion", "deletion", "formatChange", "highlight"]);
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
  const keep = new Set(["link", "comment", "insertion", "deletion", "formatChange"]);
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
  const skip = new Set(["link", "comment", "insertion", "deletion", "formatChange"]);
  return { marks: marks.filter((m) => !skip.has(m.type.name)), para: p ? { type: p.node.type.name, attrs: { ...p.node.attrs } } : null };
}

export function pasteFormat(editor: Editor, f: CopiedFormat, range?: { from: number; to: number }) {
  const { state } = editor;
  let { from, to } = range ?? state.selection;
  if (from === to) {
    const w = wordAt(state.doc, from);
    if (w) ({ from, to } = w);
  }
  const tr = state.tr;
  const skip = new Set(["link", "comment", "insertion", "deletion", "formatChange"]);
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
    // Anahat düzeyi verilmiş paragraflar da başlık gibi sayılır (Gezinti bölmesi, İçindekiler).
    if (node.type.name === "paragraph") {
      const lvl = node.attrs.outlineLevel as number | null;
      const text = node.textContent.trim();
      if (lvl && text) out.push({ level: lvl, text, pos });
      return false;
    }
    return true;
  });
  return out;
}

/** Sözcük sayımı (Türkçe harfleri ve kesme işaretli sözcükleri tanır). */
export function countWords(text: string): number {
  const m = text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-.,]*/gu);
  return m ? m.length : 0;
}

export interface TextStats {
  words: number;
  chars: number;
  charsNoSpaces: number;
  /** Boş olmayan paragraflar. */
  paragraphs: number;
}

/** Bir düğümün gövde ve not (dipnot/son not) istatistikleri. */
interface NodeStats {
  body: TextStats;
  notes: TextStats;
}

const NO_TEXT: TextStats = { words: 0, chars: 0, charsNoSpaces: 0, paragraphs: 0 };

function textStats(text: string, paragraph: boolean): TextStats {
  return { words: countWords(text), chars: text.replace(/\n/g, "").length, charsNoSpaces: text.replace(/\s/g, "").length, paragraphs: paragraph && text.trim() ? 1 : 0 };
}

function addStats(a: TextStats, b: TextStats): TextStats {
  return { words: a.words + b.words, chars: a.chars + b.chars, charsNoSpaces: a.charsNoSpaces + b.charsNoSpaces, paragraphs: a.paragraphs + b.paragraphs };
}

/** Paragrafın içerikteki [a, b) aralığı. Bloklar arasına ayraç girdiğinden sayımlar bloklar üzerinden toplanabilir. */
function textblockStats(n: PMNode, a: number, b: number): NodeStats {
  const from = Math.max(0, a);
  const to = Math.min(n.content.size, b);
  let notes = NO_TEXT;
  n.forEach((c, off) => {
    if (c.type.name === "note" && off >= from && off < to && c.attrs.text) notes = addStats(notes, textStats(String(c.attrs.text), true));
  });
  return { body: textStats(n.textBetween(from, Math.max(from, to), "\n", " "), true), notes };
}

/**
 * Düğüm başına sayım önbelleği. Düğümler değişmez ve düzenlemede değişmeyen alt ağaçlar paylaşıldığından
 * bir tuştan sonra yalnız değişen paragraf ve onun üst düğümleri yeniden sayılır.
 */
const statsCache = new WeakMap<PMNode, NodeStats>();

function nodeStats(n: PMNode): NodeStats {
  let s = statsCache.get(n);
  if (s) return s;
  if (n.isTextblock) s = textblockStats(n, 0, n.content.size);
  else if (n.isLeaf) s = { body: textStats(" ", false), notes: NO_TEXT };
  else {
    let body = NO_TEXT;
    let notes = NO_TEXT;
    n.forEach((c) => {
      const cs = nodeStats(c);
      body = addStats(body, cs.body);
      notes = addStats(notes, cs.notes);
    });
    s = { body, notes };
  }
  statsCache.set(n, s);
  return s;
}

/** Düğüm içeriğinin [from, to) aralığı: tamamen içerilen alt düğümler önbellekten gelir. */
function rangeStats(n: PMNode, from: number, to: number): NodeStats {
  if (from <= 0 && to >= n.content.size) return nodeStats(n);
  if (n.isTextblock) return textblockStats(n, from, to);
  let body = NO_TEXT;
  let notes = NO_TEXT;
  n.forEach((c, off) => {
    if (off >= to || off + c.nodeSize <= from) return;
    const cs = c.isLeaf ? nodeStats(c) : rangeStats(c, from - off - 1, to - off - 1);
    body = addStats(body, cs.body);
    notes = addStats(notes, cs.notes);
  });
  return { body, notes };
}

/**
 * Belge (ya da from–to aralığı) istatistikleri. withNotes: dipnot ve son not metinleri de sayılır
 * (Sözcük Sayısı'ndaki "Metin kutularını, dipnotları ve son notları dahil et"; metin kutuları zaten gövdededir).
 * Değişmeyen bloklar önbellekten gelir; sonuç tüm metnin tek seferde sayımıyla aynıdır.
 */
export function docStats(doc: PMNode, range?: { from: number; to: number } | null, withNotes = false): TextStats {
  const from = range ? Math.max(0, range.from) : 0;
  const to = range ? Math.min(doc.content.size, range.to) : doc.content.size;
  if (to <= from) return { ...NO_TEXT };
  const s = rangeStats(doc, from, to);
  // Önbellekteki nesne dışarı verilmez.
  return withNotes ? addStats(s.body, s.notes) : { ...s.body };
}
