// Özel düğümler: sayfa sonu, içindekiler, sayfa numarası alanı, yer işareti, liste stilleri.
import { Extension, Node, mergeAttributes } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    pageBreak: { setPageBreak: () => ReturnType };
    toc: { insertToc: (entries: TocEntry[]) => ReturnType };
    pageField: { insertPageField: (field: PageFieldKind) => ReturnType };
    bookmark: { insertBookmark: (name: string) => ReturnType };
  }
}

/**
 * Blok düzeyinde kesme (sayfa / bölüm / sütun sonu) ekler. Word'deki gibi: paragraf sonunda
 * kesme + yeni paragraf, başında kesme paragrafın önüne, ortada paragraf bölünür; boş
 * paragrafın yerini kesme alır. Liste/tablo içinde içerik ekleme kuralına düşer.
 */
export function insertBreak(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, node: PMNode): boolean {
  const { schema } = state;
  const tr = state.tr;
  if (!state.selection.empty) tr.deleteSelection();
  const $pos = tr.selection.$from;
  // Üst düzey blok: belge > bölüm > blok (derinlik 2).
  const D = 2;
  if ($pos.depth !== D || !$pos.parent.isTextblock) {
    // Tablo içinde: tablo geçerli satırdan bölünür ve kesme araya girer.
    let td = -1;
    for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === "table") td = d;
    if (td === D) {
      const tableStart = $pos.before(D);
      const rowStart = $pos.before(D + 1);
      if (rowStart === tableStart + 1) tr.insert(tableStart, node);
      else {
        tr.split(rowStart, 1);
        tr.insert(rowStart + 1, node);
      }
    } else {
      // Liste öğesi vb.: kesme öğenin içinde yeni blok olarak yer alır.
      try {
        tr.replaceSelectionWith(node, false);
      } catch {
        return false;
      }
    }
    if (dispatch) dispatch(tr.scrollIntoView());
    return true;
  }
  const para = $pos.parent;
  const before = $pos.before(D);
  const after = $pos.after(D);
  const marks = $pos.marks();
  const nextPara = () => (para.type.name === "heading" ? schema.nodes.paragraph.create() : para.type.create({ ...para.attrs, pageBreakBefore: false }));
  if (para.content.size === 0) {
    tr.replaceWith(before, after, [node, para.type.create({ ...para.attrs, pageBreakBefore: false })]);
    tr.setSelection(TextSelection.create(tr.doc, before + node.nodeSize + 1));
  } else if ($pos.parentOffset === 0) {
    tr.insert(before, node);
  } else if ($pos.parentOffset === para.content.size) {
    tr.insert(after, [node, nextPara()]);
    tr.setSelection(TextSelection.create(tr.doc, after + node.nodeSize + 1));
  } else {
    tr.split($pos.pos);
    const at = $pos.pos + 1;
    tr.insert(at, node);
    tr.setSelection(TextSelection.create(tr.doc, at + node.nodeSize + 1));
  }
  if (marks.length) tr.setStoredMarks(marks);
  if (dispatch) dispatch(tr.scrollIntoView());
  return true;
}

/** Ctrl+Enter sayfa sonu. */
export const PageBreak = Node.create({
  name: "pageBreak",
  group: "block",
  atom: true,
  selectable: true,
  parseHTML() {
    return [{ tag: "div[data-page-break]" }, { tag: "hr[data-page-break]" }];
  },
  renderHTML() {
    return ["div", { "data-page-break": "", class: "bl-pagebreak", contenteditable: "false" }, ["span", {}, "Sayfa Sonu"]];
  },
  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ state, dispatch }) =>
          insertBreak(state, dispatch, state.schema.nodes.pageBreak.create()),
    };
  },
});

export interface TocEntry {
  level: number;
  text: string;
  page: number;
  /** Başlığın belge konumu (gezinme için). */
  pos: number;
}

/** İçindekiler tablosu: başlıklardan üretilen, "Tabloyu Güncelleştir" ile yenilenen blok. */
export const TableOfContents = Node.create({
  name: "toc",
  group: "block",
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      entries: {
        default: [],
        parseHTML: (el) => {
          try {
            return JSON.parse(el.getAttribute("data-entries") ?? "[]");
          } catch {
            return [];
          }
        },
        renderHTML: (a) => ({ "data-entries": JSON.stringify(a.entries) }),
      },
      title: { default: "İçindekiler", parseHTML: (el) => el.getAttribute("data-title") ?? "İçindekiler", renderHTML: (a) => ({ "data-title": a.title }) },
      /** "toc": başlıklardan İçindekiler; "tof": resim yazılarından Şekiller Tablosu (label etiketiyle). */
      kind: { default: "toc", parseHTML: (el) => el.getAttribute("data-kind") ?? "toc", renderHTML: (a) => (a.kind !== "toc" ? { "data-kind": a.kind } : {}) },
      label: { default: null, parseHTML: (el) => el.getAttribute("data-label"), renderHTML: (a) => (a.label ? { "data-label": a.label } : {}) },
      showPages: { default: true, parseHTML: (el) => el.getAttribute("data-pages") !== "false", renderHTML: (a) => (a.showPages === false ? { "data-pages": "false" } : {}) },
    };
  },
  parseHTML() {
    return [{ tag: "div[data-toc]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    const entries = (node.attrs.entries as TocEntry[]) ?? [];
    return [
      "div",
      mergeAttributes(HTMLAttributes, { "data-toc": "", class: "bl-toc", contenteditable: "false" }),
      ...(node.attrs.title ? [["div", { class: "bl-toc-title" }, node.attrs.title as string]] : []),
      ...(entries.length
        ? entries.map((e) => [
            "div",
            { class: `bl-toc-entry bl-toc-l${e.level}`, "data-pos": String(e.pos) },
            ["span", { class: "bl-toc-text" }, e.text],
            ["span", { class: "bl-toc-dots" }],
            ["span", { class: "bl-toc-page" }, node.attrs.showPages === false ? "" : String(e.page)],
          ])
        : [["div", { class: "bl-toc-empty" }, node.attrs.kind === "tof" ? "Şekiller tablosu girdisi bulunamadı." : "İçindekiler girdisi bulunamadı."]]),
    ] as unknown as [string, ...unknown[]] as never;
  },
  addCommands() {
    return {
      insertToc:
        (entries) =>
        ({ commands }) =>
          commands.insertContent({ type: "toc", attrs: { entries } }),
    };
  },
});

/** Sayfa/tarih alanları: PAGE, NUMPAGES, DATE (gg.aa.yyyy), TIME (SS:dd). */
export type PageFieldKind = "PAGE" | "NUMPAGES" | "DATE" | "TIME";

/** Tarih/saat alanının güncel değeri (diğer alanlar için null). */
export function fieldNowText(field: unknown, d = new Date()): string | null {
  if (field === "DATE") return d.toLocaleDateString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" });
  if (field === "TIME") return d.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" });
  return null;
}

/** Alanı Güncelleştir: tarih/saat alanları şimdiki zamanla yeniden yazılır. */
export const FIELD_REFRESH_EVENT = "officehere:refresh-fields";

/** Üst/alt bilgideki sayfa numarası alanı. Statik çizimde jeton metni sayfa numarasıyla değiştirilir. */
export const PAGE_TOKEN = "⁠#SAYFA#⁠";
export const NUMPAGES_TOKEN = "⁠#TOPLAM#⁠";

export const PageField = Node.create({
  name: "pageField",
  group: "inline",
  inline: true,
  atom: true,
  addAttributes() {
    return { field: { default: "PAGE", parseHTML: (el) => el.getAttribute("data-field") ?? "PAGE", renderHTML: (a) => ({ "data-field": a.field }) } };
  },
  parseHTML() {
    return [{ tag: "span[data-field]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-field" }), fieldNowText(node.attrs.field) ?? (node.attrs.field === "NUMPAGES" ? NUMPAGES_TOKEN : PAGE_TOKEN)];
  },
  addNodeView() {
    return ({ node, editor, getPos }) => {
      const dom = document.createElement("span");
      dom.className = "bl-field is-editing";
      const now = fieldNowText(node.attrs.field);
      if (now) {
        dom.textContent = now;
        const refresh = () => (dom.textContent = fieldNowText(node.attrs.field) ?? dom.textContent);
        window.addEventListener(FIELD_REFRESH_EVENT, refresh);
        return { dom, update: (n) => n.type.name === "pageField" && n.attrs.field === node.attrs.field, destroy: () => window.removeEventListener(FIELD_REFRESH_EVENT, refresh) };
      }
      const total = node.attrs.field === "NUMPAGES";
      dom.textContent = total ? "#" : "1";
      // Belge gövdesinde alan, bulunduğu sayfanın numarasını gösterir (sayfalama denetleyicisinden).
      const ctrl = editor.extensionManager.extensions.find((e) => e.name === "belgePagination")?.options?.controller as
        | { pageOf(pos: number): number; info: { pageCount: number }; subscribe(fn: () => void): () => void }
        | undefined;
      if (!ctrl) return { dom };
      const update = () => {
        const pos = typeof getPos === "function" ? getPos() : undefined;
        dom.textContent = total ? String(ctrl.info.pageCount) : String(pos === undefined ? 1 : ctrl.pageOf(pos) + 1);
      };
      update();
      const off = ctrl.subscribe(update);
      return { dom, update: (n) => n.type.name === "pageField" && n.attrs.field === node.attrs.field && (update(), true), destroy: off };
    };
  },
  renderText({ node }) {
    return fieldNowText(node.attrs.field) ?? (node.attrs.field === "NUMPAGES" ? "#" : "1");
  },
  addCommands() {
    return {
      insertPageField:
        (field) =>
        ({ commands }) =>
          commands.insertContent({ type: "pageField", attrs: { field } }),
    };
  },
});

/** Yer işareti (görünmez çapa). */
export const Bookmark = Node.create({
  name: "bookmark",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,
  addAttributes() {
    return { name: { default: "", parseHTML: (el) => el.getAttribute("data-bookmark") ?? el.getAttribute("id"), renderHTML: (a) => ({ "data-bookmark": a.name, id: a.name }) } };
  },
  parseHTML() {
    return [{ tag: "a[data-bookmark]" }, { tag: "span[data-bookmark]" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-bookmark" })];
  },
  renderText() {
    return "";
  },
  addCommands() {
    return {
      insertBookmark:
        (name) =>
        ({ commands }) =>
          commands.insertContent({ type: "bookmark", attrs: { name } }),
    };
  },
});

/* ------------------------------------------------------------------ */
/* Liste stilleri                                                       */
/* ------------------------------------------------------------------ */

export const BULLET_STYLES: { id: string; char: string; label: string }[] = [
  { id: "disc", char: "●", label: "Dolu Yuvarlak Madde İşareti" },
  { id: "circle", char: "○", label: "Boş Yuvarlak Madde İşareti" },
  { id: "square", char: "■", label: "Dolu Kare Madde İşareti" },
  { id: "diamond", char: "❖", label: "Yıldız Madde İşareti" },
  { id: "arrow", char: "➢", label: "Ok Madde İşareti" },
  { id: "check", char: "✓", label: "Onay İşareti Madde İşareti" },
  { id: "dash", char: "–", label: "Tire Madde İşareti" },
];

export const NUMBER_STYLES: { id: string; sample: string[]; label: string }[] = [
  { id: "decimal", sample: ["1.", "2.", "3."], label: "1. 2. 3." },
  { id: "decimal-paren", sample: ["1)", "2)", "3)"], label: "1) 2) 3)" },
  { id: "upper-roman", sample: ["I.", "II.", "III."], label: "I. II. III." },
  { id: "upper-alpha", sample: ["A.", "B.", "C."], label: "A. B. C." },
  { id: "lower-alpha-paren", sample: ["a)", "b)", "c)"], label: "a) b) c)" },
  { id: "lower-alpha", sample: ["a.", "b.", "c."], label: "a. b. c." },
  { id: "lower-roman", sample: ["i.", "ii.", "iii."], label: "i. ii. iii." },
];

export const MULTILEVEL_STYLES: { id: string; label: string; sample: string[] }[] = [
  { id: "outline", label: "1. 1.1. 1.1.1.", sample: ["1.", "1.1.", "1.1.1."] },
  { id: "mixed", label: "1. a. i.", sample: ["1.", "a.", "i."] },
  { id: "bullets", label: "● ○ ■", sample: ["●", "○", "■"] },
];

/** Madde ve numaralı listelere stil öznitelikleri ekler. */
export const ListStyles = Extension.create({
  name: "listStyles",
  addGlobalAttributes() {
    return [
      {
        types: ["bulletList", "orderedList"],
        attributes: {
          listStyle: {
            default: null,
            parseHTML: (el) => el.getAttribute("data-list-style"),
            renderHTML: (a) => (a.listStyle ? { "data-list-style": a.listStyle } : {}),
          },
          multilevel: {
            default: null,
            parseHTML: (el) => el.getAttribute("data-multilevel"),
            renderHTML: (a) => (a.multilevel ? { "data-multilevel": a.multilevel } : {}),
          },
        },
      },
    ];
  },
});
