// Alanlar: resim yazısı numarası (SEQ), çapraz başvuru (REF/PAGEREF), alıntı (CITATION), kaynakça
// (BIBLIOGRAPHY) ve adres mektup birleştirme alanı (MERGEFIELD). Alan sonuçları düğüm özniteliğinde
// saklanır ve belge değiştikçe güncellenir (Word'deki alan sonucu gibi); DOCX'e alan olarak yazılır.
import { Extension, Node, mergeAttributes } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState, type Transaction } from "@tiptap/pm/state";
import { bibEntry, citationText, orderSources, type BibEntry, type BibStyle, type Source } from "../model/sources";

/* ------------------------------------------------------------------ */
/* Alan bağlamı (belge ayarlarındaki kaynaklar, birleştirme kaydı)     */
/* ------------------------------------------------------------------ */

interface FieldContext {
  sources: Source[];
  bibStyle: BibStyle;
  /** Adres mektup birleştirmede önizlenen kayıt (null: alan adları gösterilir). */
  record: Record<string, string> | null;
}

const fieldCtx: FieldContext = { sources: [], bibStyle: "apa", record: null };
export const fieldsKey = new PluginKey("belgeFields");
/** Önizleme kaydı değişince birleştirme alanı görünümleri yeniden çizilir (belge değişmez). */
const recordListeners = new Set<() => void>();

/** Kaynaklar / önizleme kaydı değişince alanları yeniler. */
export function setFieldContext(view: { state: EditorState; dispatch: (tr: Transaction) => void; isDestroyed?: boolean } | null, patch: Partial<FieldContext>) {
  Object.assign(fieldCtx, patch);
  if ("record" in patch) recordListeners.forEach((l) => l());
  if (view && !view.isDestroyed) view.dispatch(view.state.tr.setMeta(fieldsKey, { refresh: true }).setMeta("addToHistory", false));
}

export function getFieldContext(): Readonly<FieldContext> {
  return fieldCtx;
}

/* ------------------------------------------------------------------ */
/* Düğümler                                                             */
/* ------------------------------------------------------------------ */

const atomInline = { group: "inline", inline: true, atom: true, selectable: true } as const;

/** SEQ alanı (resim yazısı numarası): aynı etiketteki sırası. */
export const SeqField = Node.create({
  name: "seqField",
  ...atomInline,
  addAttributes() {
    return {
      label: { default: "Şekil", parseHTML: (el) => el.getAttribute("data-seq") ?? "Şekil", renderHTML: (a) => ({ "data-seq": a.label }) },
      num: { default: 1, parseHTML: (el) => Number(el.getAttribute("data-num") ?? el.textContent ?? 1) || 1, renderHTML: (a) => ({ "data-num": String(a.num) }) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-seq]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-fieldval bl-seq" }), String(node.attrs.num)];
  },
  renderText({ node }) {
    return String(node.attrs.num);
  },
});

export type CrossRefKind = "full" | "label" | "text" | "page" | "position";

export const CROSSREF_KIND_LABEL: Record<CrossRefKind, string> = {
  full: "Tüm metin",
  label: "Yalnızca etiket ve numara",
  text: "Yalnızca metin",
  page: "Sayfa numarası",
  position: "Yukarıda/aşağıda",
};

/** Çapraz başvuru: bir yer işaretine (resim yazısı, başlık) başvurur. */
export const CrossRef = Node.create({
  name: "crossRef",
  ...atomInline,
  addAttributes() {
    return {
      ref: { default: "", parseHTML: (el) => el.getAttribute("data-ref") ?? "", renderHTML: (a) => ({ "data-ref": a.ref }) },
      kind: { default: "full", parseHTML: (el) => el.getAttribute("data-kind") ?? "full", renderHTML: (a) => ({ "data-kind": a.kind }) },
      text: { default: "", parseHTML: (el) => el.textContent ?? "", renderHTML: () => ({}) },
      link: { default: true, parseHTML: (el) => el.getAttribute("data-link") !== "false", renderHTML: (a) => ({ "data-link": String(a.link) }) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-ref]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-fieldval bl-xref" }), (node.attrs.text as string) || "Hata! Başvuru kaynağı bulunamadı."];
  },
  renderText({ node }) {
    return node.attrs.text as string;
  },
});

/** Metin içi alıntı (kaynak etiketiyle). */
export const Citation = Node.create({
  name: "citation",
  ...atomInline,
  addAttributes() {
    return {
      tag: { default: "", parseHTML: (el) => el.getAttribute("data-cite") ?? "", renderHTML: (a) => ({ "data-cite": a.tag }) },
      pages: { default: "", parseHTML: (el) => el.getAttribute("data-pages") ?? "", renderHTML: (a) => (a.pages ? { "data-pages": a.pages } : {}) },
      text: { default: "", parseHTML: (el) => el.textContent ?? "", renderHTML: () => ({}) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-cite]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-fieldval bl-cite" }), (node.attrs.text as string) || "(Kaynak)"];
  },
  renderText({ node }) {
    return node.attrs.text as string;
  },
});

/** Adres mektup birleştirme alanı: «Ad»; önizlemede kayıttaki değer. */
export const MergeField = Node.create({
  name: "mergeField",
  ...atomInline,
  addAttributes() {
    return { name: { default: "", parseHTML: (el) => el.getAttribute("data-merge") ?? "", renderHTML: (a) => ({ "data-merge": a.name }) } };
  },
  parseHTML() {
    return [{ tag: "span[data-merge]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-fieldval bl-merge" }), `«${node.attrs.name as string}»`];
  },
  renderText({ node }) {
    return `«${node.attrs.name as string}»`;
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement("span");
      dom.className = "bl-fieldval bl-merge";
      dom.setAttribute("data-merge", node.attrs.name as string);
      let cur = node;
      const show = () => {
        const rec = fieldCtx.record;
        const name = cur.attrs.name as string;
        dom.textContent = rec ? (rec[name] ?? "") : `«${name}»`;
        dom.classList.toggle("is-preview", !!rec);
      };
      show();
      recordListeners.add(show);
      return {
        dom,
        update: (n) => {
          if (n.type.name !== "mergeField") return false;
          cur = n;
          show();
          return true;
        },
        destroy: () => recordListeners.delete(show),
      };
    };
  },
});

/** Kaynakça bloğu: kaynaklardan biçimlendirilmiş girdiler (Kaynakçayı Güncelleştir). */
export const Bibliography = Node.create({
  name: "bibliography",
  group: "block",
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      title: { default: "Kaynakça", parseHTML: (el) => el.getAttribute("data-title") ?? "Kaynakça", renderHTML: (a) => ({ "data-title": a.title }) },
      entries: {
        default: [],
        parseHTML: (el) => {
          try {
            return JSON.parse(el.getAttribute("data-entries") ?? "[]") as BibEntry[];
          } catch {
            return [];
          }
        },
        renderHTML: (a) => ({ "data-entries": JSON.stringify(a.entries) }),
      },
    };
  },
  parseHTML() {
    return [{ tag: "div[data-bibliography]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    const entries = (node.attrs.entries as BibEntry[]) ?? [];
    return [
      "div",
      mergeAttributes(HTMLAttributes, { "data-bibliography": "", class: "bl-bib", contenteditable: "false" }),
      ["h1", { class: "bl-bib-title" }, node.attrs.title as string],
      ...(entries.length
        ? entries.map((e) => ["p", { class: "bl-bib-entry", "data-style": "Bibliography" }, e.before, ["i", {}, e.italic], e.after])
        : [["p", { class: "bl-bib-empty" }, "Belgede kaynak yok."]]),
    ] as never;
  },
});

/* ------------------------------------------------------------------ */
/* Alan sonuçlarının güncellenmesi                                      */
/* ------------------------------------------------------------------ */

/** Yer işaretinin bulunduğu paragraf ve içindeki ilk SEQ alanı. */
function refTarget(doc: PMNode, name: string) {
  let found: { para: PMNode; paraPos: number; pos: number } | null = null;
  doc.descendants((n, pos, parent) => {
    if (found) return false;
    if (n.type.name === "bookmark" && n.attrs.name === name && parent) {
      const $p = doc.resolve(pos);
      found = { para: $p.parent, paraPos: $p.before(), pos };
    }
    return true;
  });
  return found as { para: PMNode; paraPos: number; pos: number } | null;
}

/** Paragrafın görünen metni (alan sonuçlarıyla). */
export function paraText(p: PMNode): string {
  let s = "";
  p.descendants((n) => {
    if (n.isText) s += n.text;
    else if (n.type.name === "seqField") s += String(n.attrs.num);
    else if (n.type.name === "crossRef" || n.type.name === "citation") s += n.attrs.text as string;
    else if (n.type.name === "hardBreak") s += " ";
    return true;
  });
  return s.trim();
}

/** Çapraz başvurunun metni (sayfa numarası türü dışarıdan, sayfalamadan gelir). */
export function crossRefText(doc: PMNode, ref: string, kind: CrossRefKind, at: number, pageOf?: (pos: number) => number): string {
  const t = refTarget(doc, ref);
  if (!t) return "Hata! Başvuru kaynağı bulunamadı.";
  if (kind === "position") return t.pos < at ? "yukarıda" : "aşağıda";
  if (kind === "page") return pageOf ? String(pageOf(t.pos) + 1) : "1";
  let seq: PMNode | null = null;
  t.para.descendants((n) => {
    if (!seq && n.type.name === "seqField") seq = n;
    return !seq;
  });
  const full = paraText(t.para);
  const s = seq as PMNode | null;
  const label = s ? `${s.attrs.label as string} ${s.attrs.num as number}` : "";
  if (kind === "label") return label || full;
  if (kind === "text") {
    if (!label) return full;
    const i = full.indexOf(label);
    return i >= 0 ? full.slice(i + label.length).replace(/^[\s:.\-–—]+/, "") : full;
  }
  return full;
}

/** Belgedeki alanların güncel sonuçlarını hesaplar; değişen varsa işlem döner. */
export function updateFields(state: EditorState, pageOf?: (pos: number) => number): Transaction | null {
  const tr = state.tr;
  const counters = new Map<string, number>();
  const cited: string[] = [];
  let changed = false;
  let hasXref = false;
  let hasBib = false;
  state.doc.descendants((n, pos) => {
    switch (n.type.name) {
      case "seqField": {
        const label = n.attrs.label as string;
        const num = (counters.get(label) ?? 0) + 1;
        counters.set(label, num);
        if (n.attrs.num !== num) {
          tr.setNodeMarkup(pos, undefined, { ...n.attrs, num });
          changed = true;
        }
        return false;
      }
      case "citation":
        if (!cited.includes(n.attrs.tag as string)) cited.push(n.attrs.tag as string);
        return false;
      case "crossRef":
        hasXref = true;
        return false;
      case "bibliography":
        hasBib = true;
        return false;
    }
    return true;
  });
  if (!cited.length && !hasXref && !hasBib) return changed ? tr.setMeta(fieldsKey, { applied: true }).setMeta("addToHistory", false) : null;
  const after = changed ? tr.doc : state.doc;
  const order = orderSources(fieldCtx.sources, fieldCtx.bibStyle, cited);
  after.descendants((n, pos) => {
    if (n.type.name === "citation") {
      const i = order.findIndex((s) => s.tag === n.attrs.tag);
      const text = citationText(order[i], fieldCtx.bibStyle, i, (n.attrs.pages as string) || undefined);
      if (text !== n.attrs.text) {
        tr.setNodeMarkup(pos, undefined, { ...n.attrs, text });
        changed = true;
      }
      return false;
    }
    if (hasXref && n.type.name === "crossRef") {
      if (n.attrs.kind === "page" && !pageOf) return false;
      const text = crossRefText(after, n.attrs.ref as string, n.attrs.kind as CrossRefKind, pos, pageOf);
      if (text !== n.attrs.text) {
        tr.setNodeMarkup(pos, undefined, { ...n.attrs, text });
        changed = true;
      }
      return false;
    }
    if (hasBib && n.type.name === "bibliography") {
      const entries = order.map((s, i) => bibEntry(s, fieldCtx.bibStyle, i));
      if (JSON.stringify(entries) !== JSON.stringify(n.attrs.entries)) {
        tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries });
        changed = true;
      }
      return false;
    }
    return true;
  });
  return changed ? tr.setMeta(fieldsKey, { applied: true }).setMeta("addToHistory", false) : null;
}

/** Alan güncelleme eklentisi: her belge değişikliğinde (ve bağlam yenilenince) sonuçlar tazelenir. */
export const FieldUpdater = Extension.create({
  name: "belgeFieldUpdater",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: fieldsKey,
        appendTransaction(trs, _old, state) {
          if (!trs.some((t) => (t.docChanged && !t.getMeta(fieldsKey)) || (t.getMeta(fieldsKey) as { refresh?: boolean } | undefined)?.refresh)) return null;
          return updateFields(state);
        },
      }),
    ];
  },
});
