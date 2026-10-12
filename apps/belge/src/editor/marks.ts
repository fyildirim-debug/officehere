// Karakter işaretleri: altı çizili türleri, karakter stilleri, ek yazı tipi özellikleri, açıklama ve izlenen değişiklikler.
import { Extension, Mark, mergeAttributes } from "@tiptap/core";

export type UnderlineType = "single" | "double" | "thick" | "dotted" | "dashed" | "wave" | "words";

export const UNDERLINE_TYPES: { id: UnderlineType; label: string }[] = [
  { id: "single", label: "Altı Çizili" },
  { id: "double", label: "Çift Altı Çizili" },
  { id: "thick", label: "Kalın Altı Çizili" },
  { id: "dotted", label: "Noktalı Altı Çizili" },
  { id: "dashed", label: "Kesik Çizgili Altı Çizili" },
  { id: "wave", label: "Dalgalı Altı Çizili" },
];

function underlineCss(type: string, color: string | null): string {
  const style = type === "double" ? "double" : type === "dotted" ? "dotted" : type === "dashed" ? "dashed" : type === "wave" ? "wavy" : "solid";
  const parts = [`text-decoration-line:underline`, `text-decoration-style:${style}`];
  if (type === "thick") parts.push("text-decoration-thickness:2px");
  if (color) parts.push(`text-decoration-color:${color}`);
  return parts.join(";");
}

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    belgeUnderline: {
      setUnderlineType: (attrs: { type?: UnderlineType; color?: string | null }) => ReturnType;
    };
    charStyle: {
      setCharStyle: (styleId: string) => ReturnType;
      unsetCharStyle: () => ReturnType;
    };
    comment: {
      setComment: (id: string) => ReturnType;
      unsetComment: (id: string) => ReturnType;
    };
  }
}

export const Underline = Mark.create({
  name: "underline",
  addAttributes() {
    return {
      type: { default: "single", parseHTML: (el) => el.getAttribute("data-u") ?? "single", renderHTML: () => ({}) },
      color: {
        default: null,
        parseHTML: (el) => el.getAttribute("data-u-color"),
        renderHTML: (a) => ({ "data-u": a.type, ...(a.color ? { "data-u-color": a.color } : {}), style: underlineCss(a.type as string, a.color as string | null) }),
      },
    };
  },
  parseHTML() {
    return [{ tag: "u" }, { style: "text-decoration", consuming: false, getAttrs: (s) => ((s as string).includes("underline") ? {} : false) }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["u", mergeAttributes(HTMLAttributes), 0];
  },
  addCommands() {
    return {
      setUnderline:
        () =>
        ({ commands }) =>
          commands.setMark(this.name),
      setUnderlineType:
        (attrs) =>
        ({ commands }) =>
          commands.setMark(this.name, attrs),
      toggleUnderline:
        () =>
        ({ commands }) =>
          commands.toggleMark(this.name),
      unsetUnderline:
        () =>
        ({ commands }) =>
          commands.unsetMark(this.name),
    };
  },
});

/** Karakter stili (Vurgu, Güçlü, Hafif Vurgu…). */
export const CharStyle = Mark.create({
  name: "charStyle",
  addAttributes() {
    return { styleId: { default: "Emphasis", parseHTML: (el) => el.getAttribute("data-cstyle"), renderHTML: (a) => ({ "data-cstyle": a.styleId }) } };
  },
  parseHTML() {
    return [{ tag: "span[data-cstyle]" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["span", HTMLAttributes, 0];
  },
  addCommands() {
    return {
      setCharStyle:
        (styleId) =>
        ({ commands }) =>
          commands.setMark(this.name, { styleId }),
      unsetCharStyle:
        () =>
        ({ commands }) =>
          commands.unsetMark(this.name),
    };
  },
});

/** TextStyle işaretine ek öznitelikler: karakter aralığı, büyük harf, küçük büyük harf, gizli, çift üstü çizili. */
export const TextStyleExtras = Extension.create({
  name: "textStyleExtras",
  addGlobalAttributes() {
    return [
      {
        types: ["textStyle"],
        attributes: {
          letterSpacing: {
            default: null,
            parseHTML: (el) => {
              const v = el.style.letterSpacing;
              return v && v.endsWith("pt") ? parseFloat(v) : v && v.endsWith("px") ? +(parseFloat(v) * 0.75).toFixed(2) : null;
            },
            renderHTML: (a) => (a.letterSpacing ? { style: `letter-spacing:${a.letterSpacing}pt` } : {}),
          },
          caps: {
            default: false,
            parseHTML: (el) => el.style.textTransform === "uppercase",
            renderHTML: (a) => (a.caps ? { style: "text-transform:uppercase" } : {}),
          },
          smallCaps: {
            default: false,
            parseHTML: (el) => el.style.fontVariant === "small-caps" || el.style.fontVariantCaps === "small-caps",
            renderHTML: (a) => (a.smallCaps ? { style: "font-variant:small-caps" } : {}),
          },
          hidden: {
            default: false,
            parseHTML: (el) => el.hasAttribute("data-hidden"),
            renderHTML: (a) => (a.hidden ? { "data-hidden": "" } : {}),
          },
          doubleStrike: {
            default: false,
            parseHTML: (el) => el.hasAttribute("data-dstrike"),
            renderHTML: (a) => (a.doubleStrike ? { "data-dstrike": "", style: "text-decoration-line:line-through;text-decoration-style:double" } : {}),
          },
          // Metin efektleri (Yazı Tipi › Metin Efektleri; DOCX w:shadow, w:outline, w:emboss, w:imprint).
          shadow: { default: false, parseHTML: (el) => el.hasAttribute("data-shadow"), renderHTML: (a) => (a.shadow ? { "data-shadow": "" } : {}) },
          outline: { default: false, parseHTML: (el) => el.hasAttribute("data-outline-text"), renderHTML: (a) => (a.outline ? { "data-outline-text": "" } : {}) },
          emboss: { default: false, parseHTML: (el) => el.hasAttribute("data-emboss"), renderHTML: (a) => (a.emboss ? { "data-emboss": "" } : {}) },
          imprint: { default: false, parseHTML: (el) => el.hasAttribute("data-imprint"), renderHTML: (a) => (a.imprint ? { "data-imprint": "" } : {}) },
        },
      },
    ];
  },
});

/** Açıklama bağlantısı: metin aralığını bir açıklama kimliğine bağlar. */
export const CommentMark = Mark.create({
  name: "comment",
  inclusive: false,
  excludes: "",
  addAttributes() {
    return { id: { default: "", parseHTML: (el) => el.getAttribute("data-comment"), renderHTML: (a) => ({ "data-comment": a.id }) } };
  },
  parseHTML() {
    return [{ tag: "span[data-comment]" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["span", mergeAttributes({ class: "bl-comment" }, HTMLAttributes), 0];
  },
  addCommands() {
    return {
      setComment:
        (id) =>
        ({ commands }) =>
          commands.setMark(this.name, { id }),
      unsetComment:
        (id) =>
        ({ tr, state, dispatch }) => {
          const type = state.schema.marks.comment;
          state.doc.descendants((node, pos) => {
            node.marks.forEach((m) => {
              if (m.type === type && m.attrs.id === id) tr.removeMark(pos, pos + node.nodeSize, m);
            });
          });
          if (dispatch) dispatch(tr);
          return true;
        },
    };
  },
});

/** İzlenen değişiklik: eklenen metin. */
export const Insertion = Mark.create({
  name: "insertion",
  inclusive: true,
  excludes: "deletion",
  addAttributes() {
    return {
      author: { default: "", parseHTML: (el) => el.getAttribute("data-author"), renderHTML: (a) => ({ "data-author": a.author }) },
      date: { default: "", parseHTML: (el) => el.getAttribute("data-date"), renderHTML: (a) => ({ "data-date": a.date }) },
    };
  },
  parseHTML() {
    return [{ tag: "ins" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["ins", mergeAttributes({ class: "bl-ins" }, HTMLAttributes), 0];
  },
});

/** İzlenen değişiklik: silinen metin (kabul edilene dek belgede kalır). */
export const Deletion = Mark.create({
  name: "deletion",
  inclusive: false,
  excludes: "insertion",
  addAttributes() {
    return {
      author: { default: "", parseHTML: (el) => el.getAttribute("data-author"), renderHTML: (a) => ({ "data-author": a.author }) },
      date: { default: "", parseHTML: (el) => el.getAttribute("data-date"), renderHTML: (a) => ({ "data-date": a.date }) },
    };
  },
  parseHTML() {
    return [{ tag: "del" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["del", mergeAttributes({ class: "bl-del" }, HTMLAttributes), 0];
  },
});

/** İzlenen değişiklik: biçim değişikliği (DOCX w:rPrChange). "old", önceki biçim işaretlerinin JSON'udur. */
export const FormatChange = Mark.create({
  name: "formatChange",
  inclusive: false,
  addAttributes() {
    return {
      author: { default: "", parseHTML: (el) => el.getAttribute("data-author"), renderHTML: (a) => ({ "data-author": a.author }) },
      date: { default: "", parseHTML: (el) => el.getAttribute("data-date"), renderHTML: (a) => ({ "data-date": a.date }) },
      old: { default: "[]", parseHTML: (el) => el.getAttribute("data-old") ?? "[]", renderHTML: (a) => ({ "data-old": a.old }) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-fmtchange]" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["span", mergeAttributes({ class: "bl-fmtchg", "data-fmtchange": "" }, HTMLAttributes), 0];
  },
});
