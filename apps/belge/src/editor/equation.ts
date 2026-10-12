// Denklem düğümü (Ekle › Denklem): doğrusal (LaTeX benzeri) girişten MathML ile çizilir. Boş paragrafa
// eklenen denklem görüntülü (ortalanmış, kendi satırında), metin içindeki satır içidir. Çift tıklama
// Denklem iletişim kutusunu açar. DOCX'te OMML (m:oMath / m:oMathPara) olarak yazılır ve okunur.
import { Node, mergeAttributes } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import { parseMath, toLinear } from "./math/ast";
import { mathElement } from "./math/render";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    equation: {
      /** Denklem ekler; seçili denklem varsa günceller. */
      insertEquation: (latex: string, display?: boolean) => ReturnType;
    };
  }
}

/** Denklemi MathML olarak çizer (ayrıştırılamayan giriş metin olarak gösterilir). */
export function renderEquation(latex: string, display: boolean): string {
  try {
    return mathElement(parseMath(latex), display);
  } catch {
    return `<math><mtext>${latex.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</mtext></math>`;
  }
}

/** Düz metin karşılığı (dışa aktarma, sözcük sayımı). */
export function equationText(latex: string): string {
  return latex;
}

export const EQUATION_EVENT = "officehere:edit-equation";

/** Hazır denklemler (Ekle › Denklem menüsü). */
export const BUILTIN_EQUATIONS: { label: string; latex: string }[] = [
  { label: "Dairenin alanı", latex: "A=\\pi r^2" },
  { label: "İkinci dereceden denklemin kökleri", latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}" },
  { label: "Pisagor teoremi", latex: "a^2+b^2=c^2" },
  { label: "Binom açılımı", latex: "(x+a)^n=\\sum_{k=0}^{n} \\left(\\begin{matrix}n\\\\k\\end{matrix}\\right)x^k a^{n-k}" },
  { label: "Fourier serisi", latex: "f(x)=a_0+\\sum_{n=1}^{\\infty} \\left(a_n\\cos \\frac{n\\pi x}{L}+b_n\\sin \\frac{n\\pi x}{L}\\right)" },
  { label: "Üstel fonksiyonun Taylor açılımı", latex: "e^x=1+\\frac{x}{1!}+\\frac{x^2}{2!}+\\frac{x^3}{3!}+\\cdots" },
  { label: "Belirli integral", latex: "\\int_0^1 x^2 dx=\\frac{1}{3}" },
];

export const Equation = Node.create({
  name: "equation",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      latex: { default: "", parseHTML: (el) => el.getAttribute("data-latex") ?? "", renderHTML: (a) => ({ "data-latex": a.latex }) },
      display: { default: false, parseHTML: (el) => el.hasAttribute("data-display"), renderHTML: (a) => (a.display ? { "data-display": "" } : {}) },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-equation]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    const attrs = mergeAttributes(HTMLAttributes, { "data-equation": "", class: `bl-eq${node.attrs.display ? " is-display" : ""}` });
    if (typeof document === "undefined") return ["span", attrs, String(node.attrs.latex)];
    const span = document.createElement("span");
    for (const [k, v] of Object.entries(attrs)) span.setAttribute(k, String(v));
    span.innerHTML = renderEquation(String(node.attrs.latex), !!node.attrs.display);
    return span;
  },
  renderText({ node }) {
    return equationText(String(node.attrs.latex));
  },
  addNodeView() {
    return ({ node, getPos, editor }) => {
      const dom = document.createElement("span");
      dom.className = `bl-eq${node.attrs.display ? " is-display" : ""}`;
      dom.contentEditable = "false";
      dom.setAttribute("data-equation", "");
      dom.title = "Düzenlemek için çift tıklayın";
      let cur = node;
      const draw = () => {
        dom.innerHTML = renderEquation(String(cur.attrs.latex), !!cur.attrs.display);
        dom.classList.toggle("is-display", !!cur.attrs.display);
      };
      draw();
      dom.addEventListener("dblclick", (e) => {
        e.preventDefault();
        const pos = typeof getPos === "function" ? getPos() : undefined;
        window.dispatchEvent(new CustomEvent(EQUATION_EVENT, { detail: { pos, editor } }));
      });
      return {
        dom,
        update: (n) => {
          if (n.type.name !== "equation") return false;
          cur = n;
          draw();
          return true;
        },
        selectNode: () => dom.classList.add("is-selected"),
        deselectNode: () => dom.classList.remove("is-selected"),
      };
    };
  },
  addCommands() {
    return {
      insertEquation:
        (latex, display) =>
        ({ state, dispatch }) => {
          const sel = state.selection;
          const type = state.schema.nodes.equation;
          if (sel instanceof NodeSelection && sel.node.type === type) {
            if (dispatch) {
              const tr = state.tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, latex, ...(display !== undefined ? { display } : {}) });
              dispatch(tr.setSelection(NodeSelection.create(tr.doc, sel.from)));
            }
            return true;
          }
          const $f = sel.$from;
          // Boş paragrafa eklenen denklem görüntülüdür (Word gibi).
          const disp = display ?? ($f.parent.isTextblock && $f.parent.content.size === 0);
          if (dispatch) {
            const tr = state.tr.replaceSelectionWith(type.create({ latex, display: disp }), false);
            dispatch(tr.scrollIntoView());
          }
          return true;
        },
    };
  },
});

/** DOCX'ten okunan denklemin doğrusal girişi. */
export { toLinear };
