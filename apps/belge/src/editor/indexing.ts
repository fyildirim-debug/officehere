// Dizin (Başvurular › Dizin): Girdiyi İşaretle (XE alanı; yalnız Tümünü Göster açıkken görünür) ve
// Dizin Ekle (INDEX): işaretli girdiler Türkçe alfabe sırasıyla, alt girdileri ve sayfa numaralarıyla listelenir.
import { Node, mergeAttributes } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorState, Transaction } from "@tiptap/pm/state";

export interface IndexItem {
  main: string;
  sub?: string;
  pages: number[];
}

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    indexing: {
      /** İmlece dizin girdisi (XE) ekler. */
      markIndexEntry: (main: string, sub?: string) => ReturnType;
      /** İmlece dizin ekler. */
      insertIndex: (columns: number) => ReturnType;
    };
  }
}

/** XE alan kodu metni. */
export const xeCode = (main: string, sub?: string) => `XE "${main}${sub ? `:${sub}` : ""}"`;

export const IndexEntry = Node.create({
  name: "indexEntry",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,
  addAttributes() {
    return {
      main: { default: "", parseHTML: (el) => el.getAttribute("data-xe") ?? "", renderHTML: (a) => ({ "data-xe": a.main }) },
      sub: { default: null, parseHTML: (el) => el.getAttribute("data-xe-sub"), renderHTML: (a) => (a.sub ? { "data-xe-sub": a.sub } : {}) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-xe]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, { class: "bl-xe", contenteditable: "false" }), `{ ${xeCode(String(node.attrs.main), (node.attrs.sub as string | null) ?? undefined)} }`];
  },
  renderText() {
    return "";
  },
  addCommands() {
    return {
      markIndexEntry:
        (main, sub) =>
        ({ state, dispatch }) => {
          if (!main.trim()) return false;
          // Word gibi girdi seçimin sonuna eklenir.
          const at = state.selection.to;
          if (dispatch) dispatch(state.tr.insert(at, state.schema.nodes.indexEntry.create({ main: main.trim(), sub: sub?.trim() || null })));
          return true;
        },
    };
  },
});

export const IndexBlock = Node.create({
  name: "indexBlock",
  group: "block",
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      entries: {
        default: [],
        parseHTML: (el) => {
          try {
            return JSON.parse(el.getAttribute("data-index") ?? "[]") as IndexItem[];
          } catch {
            return [];
          }
        },
        renderHTML: (a) => ({ "data-index": JSON.stringify(a.entries) }),
      },
      columns: { default: 2, parseHTML: (el) => Number(el.getAttribute("data-cols") ?? 2), renderHTML: (a) => ({ "data-cols": String(a.columns) }) },
    };
  },
  parseHTML() {
    return [{ tag: "div[data-index]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    const entries = (node.attrs.entries as IndexItem[]) ?? [];
    const rows = entries.length
      ? entries.map((e) => ["div", { class: e.sub ? "bl-ix-2" : "bl-ix-1" }, indexLine(e)])
      : [["div", { class: "bl-ix-empty" }, "Dizin girdisi bulunamadı."]];
    return ["div", mergeAttributes(HTMLAttributes, { class: "bl-index", contenteditable: "false", style: `column-count:${Math.max(1, Number(node.attrs.columns) || 1)}` }), ...rows];
  },
  addCommands() {
    return {
      insertIndex:
        (columns) =>
        ({ commands }) =>
          commands.insertContent({ type: "indexBlock", attrs: { columns, entries: [] } }),
    };
  },
});

/** "Girdi, 3, 5" biçimi (alt girdide ad yalnız alt girdidir). */
export function indexLine(e: IndexItem): string {
  return `${e.sub ?? e.main}${e.pages.length ? `, ${e.pages.join(", ")}` : ""}`;
}

const collator = new Intl.Collator("tr", { sensitivity: "base" });

/** Belgedeki XE girdilerinden sıralı dizin (ana girdiler ve altında alt girdiler). */
export function collectIndex(doc: PMNode, pageOf: (pos: number) => number): IndexItem[] {
  const map = new Map<string, { main: string; sub?: string; pages: Set<number> }>();
  doc.descendants((n, pos) => {
    if (n.type.name !== "indexEntry") return true;
    const main = String(n.attrs.main);
    const sub = (n.attrs.sub as string | null) ?? undefined;
    const page = pageOf(pos) + 1;
    // Alt girdisi olan ana girdi de listelenir (sayfası olmadan).
    if (sub && !map.has(main)) map.set(main, { main, pages: new Set() });
    const key = sub ? `${main}\u0000${sub}` : main;
    const e = map.get(key) ?? { main, sub, pages: new Set<number>() };
    e.pages.add(page);
    map.set(key, e);
    return false;
  });
  return [...map.values()]
    .sort((a, b) => collator.compare(a.main, b.main) || (a.sub ? (b.sub ? collator.compare(a.sub, b.sub) : 1) : b.sub ? -1 : 0))
    .map((e) => ({ main: e.main, ...(e.sub ? { sub: e.sub } : {}), pages: [...e.pages].sort((x, y) => x - y) }));
}

/** Tüm dizinleri günceller (değişiklik yoksa null). */
export function indexUpdate(state: EditorState, pageOf: (pos: number) => number): Transaction | null {
  const entries = collectIndex(state.doc, pageOf);
  const tr = state.tr;
  let changed = false;
  state.doc.descendants((n, pos) => {
    if (n.type.name !== "indexBlock") return true;
    if (JSON.stringify(n.attrs.entries) !== JSON.stringify(entries)) {
      tr.setNodeMarkup(pos, undefined, { ...n.attrs, entries });
      changed = true;
    }
    return false;
  });
  return changed ? tr : null;
}
