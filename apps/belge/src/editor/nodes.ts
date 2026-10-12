// Özel düğümler: sayfa sonu, içindekiler, sayfa numarası alanı, yer işareti, liste stilleri.
import { Extension, Node, mergeAttributes } from "@tiptap/core";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    pageBreak: { setPageBreak: () => ReturnType };
    toc: { insertToc: (entries: TocEntry[]) => ReturnType };
    pageField: { insertPageField: (field: "PAGE" | "NUMPAGES") => ReturnType };
    bookmark: { insertBookmark: (name: string) => ReturnType };
  }
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
        ({ chain }) =>
          chain()
            .insertContent([{ type: "pageBreak" }, { type: "paragraph" }])
            .run(),
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
      ["div", { class: "bl-toc-title" }, node.attrs.title as string],
      ...(entries.length
        ? entries.map((e) => [
            "div",
            { class: `bl-toc-entry bl-toc-l${e.level}`, "data-pos": String(e.pos) },
            ["span", { class: "bl-toc-text" }, e.text],
            ["span", { class: "bl-toc-dots" }],
            ["span", { class: "bl-toc-page" }, String(e.page)],
          ])
        : [["div", { class: "bl-toc-empty" }, "İçindekiler girdisi bulunamadı."]]),
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
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-field" }), node.attrs.field === "NUMPAGES" ? NUMPAGES_TOKEN : PAGE_TOKEN];
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement("span");
      dom.className = "bl-field is-editing";
      dom.textContent = node.attrs.field === "NUMPAGES" ? "#" : "1";
      return { dom };
    };
  },
  renderText({ node }) {
    return node.attrs.field === "NUMPAGES" ? "#" : "1";
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
