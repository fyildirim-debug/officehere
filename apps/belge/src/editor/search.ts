// Bul / Değiştir: eşleşmeleri dekorasyonla vurgular; Türkçe büyük/küçük harf duyarsız arama.
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

export interface SearchOptions {
  query: string;
  caseSensitive: boolean;
  wholeWord: boolean;
}

export interface SearchMatch {
  from: number;
  to: number;
}

interface SearchState extends SearchOptions {
  matches: SearchMatch[];
  current: number;
  deco: DecorationSet;
}

export const searchKey = new PluginKey<SearchState>("belgeSearch");

/** Türkçe duyarlı katlama: İ/I/ı/i aynı kabul edilir; uzunluk korunur. */
export function foldTr(s: string): string {
  let out = "";
  for (const ch of s) {
    const l = ch.toLocaleLowerCase("tr-TR");
    out += l === "ı" ? "i" : l.length === ch.length ? l : ch;
  }
  return out;
}

const WORD_CHAR = /[\p{L}\p{N}_]/u;

/** Belgedeki eşleşmeleri bulur (yalnız metin blokları içinde). */
export function findMatches(doc: PMNode, opts: SearchOptions): SearchMatch[] {
  const out: SearchMatch[] = [];
  if (!opts.query) return out;
  const q = opts.caseSensitive ? opts.query : foldTr(opts.query);
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Metni konum eşlemesiyle topla (satır içi düğümler yer tutucu karakterle).
    let text = "";
    const map: number[] = [];
    node.forEach((child, off) => {
      if (child.isText) {
        const t = child.text ?? "";
        for (let i = 0; i < t.length; i++) map.push(pos + 1 + off + i);
        text += t;
      } else {
        map.push(pos + 1 + off);
        text += "￼";
      }
    });
    const hay = opts.caseSensitive ? text : foldTr(text);
    let idx = hay.indexOf(q);
    while (idx >= 0) {
      const end = idx + q.length;
      const okWord = !opts.wholeWord || ((idx === 0 || !WORD_CHAR.test(hay[idx - 1])) && (end >= hay.length || !WORD_CHAR.test(hay[end])));
      if (okWord) out.push({ from: map[idx], to: map[end - 1] + 1 });
      idx = hay.indexOf(q, idx + Math.max(1, q.length));
    }
    return false;
  });
  return out;
}

function build(doc: PMNode, matches: SearchMatch[], current: number) {
  return DecorationSet.create(
    doc,
    matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === current ? "bl-hit is-current" : "bl-hit" })),
  );
}

export const Search = Extension.create({
  name: "belgeSearch",
  addProseMirrorPlugins() {
    return [
      new Plugin<SearchState>({
        key: searchKey,
        state: {
          init: () => ({ query: "", caseSensitive: false, wholeWord: false, matches: [], current: -1, deco: DecorationSet.empty }),
          apply(tr, value, _old, state) {
            const meta = tr.getMeta(searchKey) as Partial<SearchState> | undefined;
            if (meta) {
              const next = { ...value, ...meta };
              if ("query" in meta || "caseSensitive" in meta || "wholeWord" in meta) next.matches = findMatches(state.doc, next);
              if (next.current >= next.matches.length) next.current = next.matches.length ? 0 : -1;
              return { ...next, deco: build(state.doc, next.matches, next.current) };
            }
            if (tr.docChanged && value.query) {
              const matches = findMatches(state.doc, value);
              const current = Math.min(value.current, matches.length - 1);
              return { ...value, matches, current, deco: build(state.doc, matches, current) };
            }
            return value;
          },
        },
        props: {
          decorations: (state) => searchKey.getState(state)?.deco,
        },
      }),
    ];
  },
});

export function getSearch(state: EditorState): SearchState | undefined {
  return searchKey.getState(state);
}

export function setSearch(view: EditorView, opts: Partial<SearchOptions> & { current?: number }) {
  view.dispatch(view.state.tr.setMeta(searchKey, opts).setMeta("addToHistory", false));
}

export function clearSearch(view: EditorView) {
  setSearch(view, { query: "", current: -1 });
}

function selectMatch(view: EditorView, m: SearchMatch, index: number) {
  const tr: Transaction = view.state.tr.setSelection(TextSelection.create(view.state.doc, m.from, m.to)).setMeta(searchKey, { current: index }).scrollIntoView();
  view.dispatch(tr);
}

/** Sonraki/önceki eşleşmeye gider (imleç konumundan). */
export function gotoMatch(view: EditorView, dir: 1 | -1): boolean {
  const st = getSearch(view.state);
  if (!st || !st.matches.length) return false;
  const { from, to } = view.state.selection;
  let idx: number;
  if (dir > 0) {
    idx = st.matches.findIndex((m) => m.from >= to && !(m.from === from && m.to === to));
    if (idx < 0) idx = 0;
  } else {
    idx = -1;
    for (let i = st.matches.length - 1; i >= 0; i--) {
      if (st.matches[i].to <= from && !(st.matches[i].from === from)) {
        idx = i;
        break;
      }
    }
    if (idx < 0) idx = st.matches.length - 1;
  }
  selectMatch(view, st.matches[idx], idx);
  return true;
}

export function gotoMatchIndex(view: EditorView, idx: number) {
  const st = getSearch(view.state);
  if (!st?.matches[idx]) return;
  selectMatch(view, st.matches[idx], idx);
}

/**
 * Büyük/küçük harf eşleştirme kapalıyken küçük harfli yeni metin bulunan metnin
 * biçimine uyar (Word davranışı): "ELMA" → "AYVA", "Elma" → "Ayva".
 */
export function adaptCase(found: string, replacement: string, caseSensitive: boolean): string {
  if (caseSensitive || !replacement || replacement !== replacement.toLocaleLowerCase("tr-TR")) return replacement;
  const letters = found.replace(/[^\p{L}]/gu, "");
  if (!letters) return replacement;
  if (letters.length > 1 && letters === letters.toLocaleUpperCase("tr-TR") && letters !== letters.toLocaleLowerCase("tr-TR")) return replacement.toLocaleUpperCase("tr-TR");
  const first = letters[0];
  if (first !== first.toLocaleLowerCase("tr-TR")) return replacement.charAt(0).toLocaleUpperCase("tr-TR") + replacement.slice(1);
  return replacement;
}

/** Seçili eşleşmeyi değiştirir ve sonrakine geçer. */
export function replaceCurrent(view: EditorView, replacement: string): boolean {
  const st = getSearch(view.state);
  if (!st) return false;
  const { from, to } = view.state.selection;
  const hit = st.matches.find((m) => m.from === from && m.to === to);
  if (!hit) return gotoMatch(view, 1);
  const marks = view.state.doc.resolve(from).marksAcross(view.state.doc.resolve(to)) ?? undefined;
  const tr = view.state.tr;
  const text = adaptCase(view.state.doc.textBetween(from, to), replacement, st.caseSensitive);
  if (text) tr.replaceWith(from, to, view.state.schema.text(text, marks));
  else tr.delete(from, to);
  tr.setSelection(TextSelection.create(tr.doc, from + text.length));
  view.dispatch(tr);
  gotoMatch(view, 1);
  return true;
}

/** Tüm eşleşmeleri değiştirir; değiştirilen sayısını döner. */
export function replaceAll(view: EditorView, replacement: string): number {
  const st = getSearch(view.state);
  if (!st || !st.matches.length) return 0;
  const tr = view.state.tr;
  for (let i = st.matches.length - 1; i >= 0; i--) {
    const m = st.matches[i];
    const marks = tr.doc.resolve(m.from).marksAcross(tr.doc.resolve(m.to)) ?? undefined;
    const text = adaptCase(tr.doc.textBetween(m.from, m.to), replacement, st.caseSensitive);
    if (text) tr.replaceWith(m.from, m.to, view.state.schema.text(text, marks));
    else tr.delete(m.from, m.to);
  }
  const n = st.matches.length;
  view.dispatch(tr);
  return n;
}
