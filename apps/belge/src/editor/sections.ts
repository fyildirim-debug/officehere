// Bölümler (OOXML sectPr karşılığı): belge = bölüm+, her bölüm kendi sayfa yapısını ve sütun
// düzenini taşır. Bölüm sonu = iki bölüm düğümü arasındaki sınır. Sütun sonu ayrı bir düğümdür.
import { Extension, Node, mergeAttributes, type ChainedCommands, type Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { AllSelection, Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import { SECTION_KIND_LABEL, pageMetrics, sectionPage, type PageSetup, type SectionAttrs, type SectionKind } from "../model/document";
import { insertBreak } from "./nodes";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    sections: {
      /** İmleçte bölüm sonu ekler (Düzen > Kesmeler > Bölüm Sonları). */
      insertSectionBreak: (kind: SectionKind) => ReturnType;
      /** Sütun sonu (Ctrl+Shift+Enter). */
      insertColumnBreak: () => ReturnType;
    };
  }
}

/** Belge kökü: yalnız bölümlerden oluşur. */
export const BelgeDocument = Node.create({
  name: "doc",
  topNode: true,
  content: "section+",
});

export const Section = Node.create({
  name: "section",
  content: "block+",
  // Bölüm sınırı yalnız açık komutlarla (bölüm sonunu silme) birleştirilir/bölünür.
  isolating: true,
  defining: true,
  addAttributes() {
    return {
      kind: {
        default: "nextPage",
        parseHTML: (el) => (el.getAttribute("data-kind") as SectionKind | null) ?? "nextPage",
        renderHTML: (a) => ({ "data-kind": a.kind }),
      },
      page: {
        default: null,
        parseHTML: (el) => {
          const raw = el.getAttribute("data-page");
          if (!raw) return null;
          try {
            return JSON.parse(raw) as PageSetup;
          } catch {
            return null;
          }
        },
        renderHTML: (a) => (a.page ? { "data-page": JSON.stringify(a.page) } : {}),
      },
    };
  },
  parseHTML() {
    return [{ tag: "section[data-kind]" }, { tag: "section.bl-section" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["section", mergeAttributes({ class: "bl-section" }, HTMLAttributes), 0];
  },
  addCommands() {
    return {
      insertSectionBreak:
        (kind) =>
        ({ state, dispatch }) =>
          insertSectionBreak(state, dispatch, kind),
      insertColumnBreak:
        () =>
        ({ state, dispatch }) =>
          insertBreak(state, dispatch, state.schema.nodes.columnBreak.create()),
    };
  },
  addProseMirrorPlugins() {
    return [sectionEndPlugin()];
  },
});

/** Bölüm kısayolları (temel Geri Sil/Sil davranışından önce çalışır). */
export const SectionKeys = Extension.create({
  name: "belgeSectionKeys",
  priority: 1100,
  addKeyboardShortcuts() {
    return {
      Backspace: ({ editor }) => removeSectionBreak(editor.state, editor.view.dispatch, -1),
      Delete: ({ editor }) => removeSectionBreak(editor.state, editor.view.dispatch, 1),
      "Mod-Shift-Enter": ({ editor }) => editor.commands.insertColumnBreak(),
    };
  },
});

/** Sütun sonu: çok sütunlu bölümde sonraki sütuna, tek sütunda sonraki sayfaya geçer. */
export const ColumnBreak = Node.create({
  name: "columnBreak",
  group: "block",
  atom: true,
  selectable: true,
  parseHTML() {
    return [{ tag: "div[data-column-break]" }];
  },
  renderHTML() {
    return ["div", { "data-column-break": "", class: "bl-colbreak", contenteditable: "false" }, ["span", {}, "Sütun Sonu"]];
  },
});

/* ------------------------------------------------------------------ */
/* Komutlar                                                            */
/* ------------------------------------------------------------------ */

/** İmlecin bulunduğu bölüm. */
export function sectionAt(doc: PMNode, pos: number): { node: PMNode; pos: number; index: number } | null {
  const $p = doc.resolve(Math.max(0, Math.min(pos, doc.content.size)));
  if ($p.depth >= 1 && $p.node(1).type.name === "section") return { node: $p.node(1), pos: $p.before(1), index: $p.index(0) };
  // Bölümler arası (konum 0 ya da sınır): sonraki bölüm
  const index = Math.min($p.index(0), doc.childCount - 1);
  let off = 0;
  for (let i = 0; i < index; i++) off += doc.child(i).nodeSize;
  return doc.childCount ? { node: doc.child(index), pos: off, index } : null;
}

/** Bölüm sonu ekler: bölüm imleçte ikiye ayrılır, yeni bölüm verilen türle başlar. */
export function insertSectionBreak(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, kind: SectionKind): boolean {
  const tr = state.tr;
  if (!state.selection.empty) tr.deleteSelection();
  const $pos = tr.selection.$from;
  if ($pos.depth < 2 || $pos.node(1).type.name !== "section") return false;
  const section = $pos.node(1);
  const attrs: SectionAttrs = { kind, page: (section.attrs.page as PageSetup | null) ?? null };
  const schema = state.schema;
  const para = $pos.depth === 2 && $pos.parent.isTextblock ? $pos.parent : null;
  let cursor: number;
  if (para && $pos.parentOffset > 0 && $pos.parentOffset < para.content.size) {
    // Paragraf ortası: paragraf ve bölüm birlikte bölünür.
    // typesAfter dıştan içe: önce bölüm, sonra paragraf.
    tr.split($pos.pos, 2, [{ type: schema.nodes.section, attrs }, { type: para.type, attrs: { ...para.attrs, pageBreakBefore: false } }]);
    cursor = $pos.pos + 4;
  } else if (para && $pos.parentOffset === 0 && para.content.size > 0) {
    // Dolu paragrafın başı: paragraf yeni bölüme geçer; ilk paragrafsa önüne boş paragraf eklenir.
    let at = $pos.before(2);
    if ($pos.index(1) === 0) {
      tr.insert(at, schema.nodes.paragraph.create());
      at += 2;
    }
    tr.split(at, 1, [{ type: schema.nodes.section, attrs }]);
    cursor = at + 3;
  } else {
    // Paragraf sonu, boş paragraf ya da liste/tablo içi: üst düzey bloktan sonra bölünür, yeni bölüm boş paragrafla başlar.
    const after = $pos.after(2);
    const next = para && para.type.name === "heading" ? schema.nodes.paragraph.create() : schema.nodes.paragraph.create(para ? { ...para.attrs, pageBreakBefore: false } : null);
    tr.insert(after, next);
    tr.split(after, 1, [{ type: schema.nodes.section, attrs }]);
    cursor = after + 3;
  }
  // Yeni bölümün türünü (bölünen ilk parçaya değil ikinciye) ayarla.
  const $c = tr.doc.resolve(Math.min(cursor, tr.doc.content.size));
  if ($c.depth >= 1 && $c.node(1).type.name === "section") tr.setNodeMarkup($c.before(1), undefined, attrs);
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(cursor, tr.doc.content.size))));
  if (dispatch) dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Bölüm sonunu siler (Geri Sil bölüm başında / Sil bölüm sonunda). Word'deki gibi birleşen
 * bölüm sonraki bölümün biçimini alır; sınırdaki paragraflar birleşir.
 */
export function removeSectionBreak(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, dir: -1 | 1): boolean {
  const sel = state.selection;
  if (!sel.empty) return false;
  const $c = sel.$from;
  if ($c.depth !== 2 || !$c.parent.isTextblock) return false;
  const secIndex = $c.index(0);
  if (dir < 0) {
    if ($c.parentOffset !== 0 || $c.index(1) !== 0 || secIndex === 0) return false;
  } else if ($c.parentOffset !== $c.parent.content.size || $c.index(1) !== $c.node(1).childCount - 1 || secIndex === state.doc.childCount - 1) return false;
  const a = state.doc.child(dir < 0 ? secIndex - 1 : secIndex);
  const b = state.doc.child(dir < 0 ? secIndex : secIndex + 1);
  let aPos = 0;
  for (let i = 0; i < (dir < 0 ? secIndex - 1 : secIndex); i++) aPos += state.doc.child(i).nodeSize;
  const boundary = aPos + a.nodeSize;
  const tr = state.tr;
  tr.join(boundary, 1);
  tr.setNodeMarkup(aPos, undefined, { kind: a.attrs.kind, page: b.attrs.page });
  // Sınırdaki iki metin bloğu birleşir (bölüm işareti paragraf işaretinin yerindeydi).
  const inner = boundary - 1;
  const $j = tr.doc.resolve(inner);
  const before = $j.nodeBefore;
  const after = $j.nodeAfter;
  if (before?.isTextblock && after?.isTextblock && before.type === after.type) {
    tr.join(inner);
    tr.setSelection(TextSelection.create(tr.doc, inner - 1));
  } else tr.setSelection(TextSelection.near(tr.doc.resolve(inner)));
  if (dispatch) dispatch(tr.scrollIntoView());
  return true;
}

/* ------------------------------------------------------------------ */
/* Her bölüm bir paragrafla biter                                      */
/* ------------------------------------------------------------------ */

const END_OK = new Set(["paragraph", "heading", "bulletList", "orderedList"]);

function sectionEndPlugin() {
  return new Plugin({
    key: new PluginKey("belgeSectionEnd"),
    appendTransaction(trs, _old, state) {
      if (!trs.some((t) => t.docChanged)) return null;
      let tr: Transaction | null = null;
      let off = 0;
      const ends: number[] = [];
      state.doc.forEach((sec) => {
        const last = sec.lastChild;
        if (!last || !END_OK.has(last.type.name)) ends.push(off + sec.nodeSize - 1);
        off += sec.nodeSize;
      });
      for (let i = ends.length - 1; i >= 0; i--) {
        tr ??= state.tr;
        tr.insert(ends[i], state.schema.nodes.paragraph.create());
      }
      return tr?.setMeta("addToHistory", false) ?? null;
    },
  });
}

/* ------------------------------------------------------------------ */
/* Bölüm geometrisi                                                    */
/* ------------------------------------------------------------------ */

export interface SecGeom {
  index: number;
  /** Bölüm düğümünün konumu. */
  pos: number;
  end: number;
  kind: SectionKind;
  page: PageSetup;
  w: number;
  h: number;
  mt: number;
  mb: number;
  ml: number;
  mr: number;
  headerY: number;
  footerY: number;
  contentW: number;
  cols: number;
  colGap: number;
  colW: number;
  line: boolean;
  titlePage: boolean;
}

export function geomOf(page: PageSetup, index = 0, pos = 0, end = 0, kind: SectionKind = "nextPage"): SecGeom {
  const m = pageMetrics(page);
  const cols = Math.max(1, Math.min(13, page.columns || 1));
  const colGap = cols > 1 ? page.columnSpace / 15 : 0;
  const colW = cols > 1 ? (m.contentW - (cols - 1) * colGap) / cols : m.contentW;
  return { index, pos, end, kind, page, w: m.w, h: m.h, mt: m.mt, mb: m.mb, ml: m.ml, mr: m.mr, headerY: m.headerY, footerY: m.footerY, contentW: m.contentW, cols, colGap, colW, line: !!page.columnLine && cols > 1, titlePage: !!page.titlePage };
}

/** Belgedeki bölümlerin etkin geometrisi. */
export function sectionGeoms(doc: PMNode, def: PageSetup): SecGeom[] {
  const out: SecGeom[] = [];
  let off = 0;
  doc.forEach((sec, _o, i) => {
    out.push(geomOf(sectionPage(sec.attrs as SectionAttrs, def), i, off, off + sec.nodeSize, (sec.attrs.kind as SectionKind) ?? "nextPage"));
    off += sec.nodeSize;
  });
  if (!out.length) out.push(geomOf(def));
  return out;
}

/** Gövde genişliği: en geniş sayfa (dar sayfalar bu genişlikte ortalanır). */
export function bodyWidth(geoms: SecGeom[]): number {
  return Math.max(...geoms.map((g) => g.w));
}

/** Bölüm kutusunun yatay dolgusu (gövde genişliğinde ortalanan sayfanın kenar boşlukları). */
export function sectionPadding(g: SecGeom, bodyW: number) {
  const left = (bodyW - g.w) / 2;
  return { pl: left + g.ml, pr: bodyW - left - g.w + g.mr, pageLeft: left };
}

/* ------------------------------------------------------------------ */
/* Bölüm görünümü: kenar boşlukları, sütunlar, bölüm sonu işaretleri  */
/* ------------------------------------------------------------------ */

export interface SectionLayoutSource {
  /** Yazdırma Düzeni etkin mi (Web Düzeni'nde bölüm kenar boşlukları uygulanmaz). */
  enabled(): boolean;
  defaultPage(): PageSetup;
}

export const sectionLayoutKey = new PluginKey<DecorationSet>("belgeSectionLayout");

function secMark(label: string) {
  return () => {
    const d = document.createElement("div");
    d.className = "bl-secmark";
    d.contentEditable = "false";
    const s = document.createElement("span");
    const b = document.createElement("b");
    b.textContent = `Bölüm Sonu (${label})`;
    s.appendChild(b);
    d.appendChild(s);
    return d;
  };
}

export function buildSectionDecos(doc: PMNode, src: SectionLayoutSource): DecorationSet {
  const geoms = sectionGeoms(doc, src.defaultPage());
  const decos: Decoration[] = [];
  const paged = src.enabled();
  const bodyW = bodyWidth(geoms);
  geoms.forEach((g, i) => {
    if (paged) {
      const { pl, pr } = sectionPadding(g, bodyW);
      let style = `padding-left:${pl.toFixed(2)}px;padding-right:${pr.toFixed(2)}px`;
      const attrs: Record<string, string> = {};
      if (g.cols > 1) {
        // Ölçüm sırasında bölüm tek ve dar bir sütun olarak dizilir (--bl-npr).
        style += `;column-count:${g.cols};column-gap:${g.colGap.toFixed(2)}px;--bl-npr:${(pr + (g.contentW - g.colW)).toFixed(2)}px`;
        if (g.line) style += ";column-rule:0.75pt solid #000";
        attrs.class = "is-cols";
      }
      attrs.style = style;
      decos.push(Decoration.node(g.pos, g.end, attrs));
    }
    if (i < geoms.length - 1) {
      const next = geoms[i + 1];
      decos.push(Decoration.widget(g.end - 1, secMark(SECTION_KIND_LABEL[next.kind]), { side: 1, key: `secmark:${next.kind}`, ignoreSelection: true }));
    }
  });
  return DecorationSet.create(doc, decos);
}

/** Bölüm düzeni eklentisi; kaynak (denetleyici) değişince `sectionLayoutKey` meta ile yenilenir. */
export function createSectionLayout(src: SectionLayoutSource) {
  return Extension.create({
    name: "belgeSectionLayout",
    addProseMirrorPlugins() {
      return [
        new Plugin<DecorationSet>({
          key: sectionLayoutKey,
          state: {
            init: (_c, state) => buildSectionDecos(state.doc, src),
            apply: (tr, old, _o, state) => (tr.docChanged || tr.getMeta(sectionLayoutKey) ? buildSectionDecos(state.doc, src) : old),
          },
          props: {
            decorations: (state) => sectionLayoutKey.getState(state),
          },
        }),
      ];
    },
  });
}

/**
 * Seçili metne sütun uygulanırken (Word gibi) seçimin önüne ve arkasına sürekli bölüm sonu
 * eklenir; ortadaki yeni bölümün konumu döner. Seçim tek bölüm içinde değilse null.
 */
export function isolateSelection(state: EditorState, dispatch: (tr: Transaction) => void, page: PageSetup | null): number | null {
  const { $from, $to, empty } = state.selection;
  if (empty || $from.depth < 2 || $to.depth < 2 || $from.index(0) !== $to.index(0)) return null;
  const secPos = $from.before(1);
  const sec = $from.node(1);
  const start = $from.before(2);
  const end = $to.after(2);
  const tr = state.tr;
  const attrs = { kind: "continuous" as SectionKind, page };
  // Önce son (konumlar kaymasın), sonra baş.
  if (end < secPos + sec.nodeSize - 1) tr.split(end, 1, [{ type: state.schema.nodes.section, attrs: { kind: "continuous", page: sec.attrs.page } }]);
  let mid = secPos;
  if (start > secPos + 1) {
    tr.split(start, 1, [{ type: state.schema.nodes.section, attrs }]);
    mid = start + 1;
  }
  tr.setSelection(state.selection.map(tr.doc, tr.mapping));
  dispatch(tr);
  return mid;
}

/**
 * Blok saran komutları (listeler) bölüm sınırını aşamaz: Tümünü Seç ya da birden çok bölüme yayılan
 * seçimde komut her bölümün seçili kısmına ayrı ayrı (sondan başa, konumlar kaymasın diye) uygulanır.
 * Hepsi tek işlemde yapılır (tek Geri Al adımı).
 */
export function sectionWise(editor: Editor, apply: (c: ChainedCommands) => ChainedCommands): boolean {
  const sel = editor.state.selection;
  const all = sel instanceof AllSelection;
  const count = editor.state.doc.childCount;
  const first = sel.$from.depth ? sel.$from.index(0) : 0;
  const last = Math.min(count - 1, sel.$to.depth ? sel.$to.index(0) : count - 1);
  if (!all && first === last) return apply(editor.chain().focus()).run();
  const size = editor.state.doc.content.size;
  let c = editor.chain().focus();
  for (let i = last; i >= first; i--) {
    c = c.command(({ tr }) => {
      let start = 0;
      for (let k = 0; k < i; k++) start += tr.doc.child(k).nodeSize;
      const end = start + tr.doc.child(i).nodeSize;
      const from = i === first ? Math.max(sel.from, start + 1) : start + 1;
      const to = i === last ? Math.min(sel.to + (end - (start + editor.state.doc.child(i).nodeSize)), end - 1) : end - 1;
      tr.setSelection(TextSelection.between(tr.doc.resolve(from), tr.doc.resolve(Math.max(from, to))));
      return true;
    });
    c = apply(c);
  }
  // Seçim geri kurulur (Tümünü Seç korunur).
  c = c.command(({ tr }) => {
    const doc = tr.doc;
    tr.setSelection(all ? new AllSelection(doc) : TextSelection.between(doc.resolve(Math.min(sel.from, doc.content.size)), doc.resolve(Math.max(0, doc.content.size - (size - sel.to)))));
    return true;
  });
  return c.run();
}
