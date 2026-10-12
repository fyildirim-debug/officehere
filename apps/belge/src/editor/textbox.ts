// Metin kutusu: paragrafa bağlı, serbest konumlu (metnin önünde), taşınabilir ve boyutlandırılabilir
// zengin metin alanı. Belge akışında sıfır yükseklikli bir çapa bloğu olarak yer alır; kutu çapaya
// göre (x, y) ötelenir. DOCX'te wps:txbx (yedek: v:textbox) olarak yazılır.
import { Node, mergeAttributes } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import type { EditorView, NodeView } from "@tiptap/pm/view";

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    textBox: {
      /** İmleçteki paragrafa bağlı basit metin kutusu ekler. */
      insertTextBox: (attrs?: Partial<TextBoxAttrs>, text?: string) => ReturnType;
    };
  }
}

export interface TextBoxAttrs {
  /** Çapaya göre konum ve boyut (px). */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Kenarlık rengi (null: kenarlık yok) ve kalınlığı (pt). */
  border: string | null;
  borderWidth: number;
  /** Dolgu (null: dolgu yok). */
  fill: string | null;
}

export const TEXTBOX_PLACEHOLDER = "[Belgeden bir alıntı yapmak ya da ilginç bir noktayı vurgulamak için bu metin kutusunu kullanın. Kutuyu sayfada istediğiniz yere taşımak için kenarından sürükleyin.]";

const num = (v: string | null, d: number) => {
  const n = v === null ? NaN : parseFloat(v);
  return Number.isFinite(n) ? n : d;
};

export const TextBox = Node.create({
  name: "textBox",
  group: "block",
  content: "(paragraph | heading | bulletList | orderedList)+",
  isolating: true,
  defining: true,
  selectable: true,
  draggable: false,
  addAttributes() {
    return {
      x: { default: 0, parseHTML: (el) => num(el.getAttribute("data-x"), 0), renderHTML: (a) => ({ "data-x": a.x }) },
      y: { default: 0, parseHTML: (el) => num(el.getAttribute("data-y"), 0), renderHTML: (a) => ({ "data-y": a.y }) },
      w: { default: 240, parseHTML: (el) => num(el.getAttribute("data-w"), 240), renderHTML: (a) => ({ "data-w": a.w }) },
      h: { default: 120, parseHTML: (el) => num(el.getAttribute("data-h"), 120), renderHTML: (a) => ({ "data-h": a.h }) },
      border: { default: "#000000", parseHTML: (el) => (el.hasAttribute("data-border") ? el.getAttribute("data-border") || null : "#000000"), renderHTML: (a) => ({ "data-border": a.border ?? "" }) },
      borderWidth: { default: 0.75, parseHTML: (el) => num(el.getAttribute("data-bw"), 0.75), renderHTML: (a) => ({ "data-bw": a.borderWidth }) },
      fill: { default: "#FFFFFF", parseHTML: (el) => (el.hasAttribute("data-fill") ? el.getAttribute("data-fill") || null : "#FFFFFF"), renderHTML: (a) => ({ "data-fill": a.fill ?? "" }) },
    };
  },
  parseHTML() {
    return [{ tag: "div[data-textbox]", contentElement: ".bl-tb-content" }, { tag: "div[data-textbox]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    // Statik çizim (HTML dışa aktarma, okuma modu): çapa + mutlak konumlu kutu.
    return ["div", mergeAttributes(HTMLAttributes, { "data-textbox": "", class: "bl-tb-anchor" }), ["div", { class: "bl-textbox", style: boxStyle(node.attrs as TextBoxAttrs) }, ["div", { class: "bl-tb-content" }, 0]]];
  },
  addNodeView() {
    return ({ node, view, getPos }) => new TextBoxView(node, view, getPos as () => number | undefined);
  },
  addCommands() {
    return {
      insertTextBox:
        (attrs = {}, text = TEXTBOX_PLACEHOLDER) =>
        ({ state, tr, dispatch }) => {
          const schema = state.schema;
          const box = schema.nodes.textBox.create({ ...attrs }, schema.nodes.paragraph.create(null, text ? schema.text(text) : undefined));
          const $from = state.selection.$from;
          // Çapa: imleçteki üst düzey bloğun önü (belge > bölüm > blok).
          const depth = Math.min(2, $from.depth);
          const at = depth >= 2 ? $from.before(2) : $from.pos;
          if (dispatch) {
            tr.insert(at, box);
            // Kutudaki metin seçili gelir: yazınca yer tutucu metnin yerini alır.
            const start = at + 2;
            tr.setSelection(TextSelection.create(tr.doc, start, start + (text?.length ?? 0)));
            tr.scrollIntoView();
          }
          return true;
        },
    };
  },
});

export function boxStyle(a: TextBoxAttrs): string {
  const s = [`left:${a.x}px`, `top:${a.y}px`, `width:${a.w}px`, `height:${a.h}px`];
  s.push(a.border ? `border:${a.borderWidth}pt solid ${a.border}` : "border:0");
  s.push(`background:${a.fill ?? "transparent"}`);
  return s.join(";");
}

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;

/** Kutu düğüm görünümü: içerik düzenlenebilir; kenarlık taşır, tutamaçlar boyutlandırır. */
class TextBoxView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  private box: HTMLElement;

  constructor(
    private node: PMNode,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "bl-tb-anchor";
    this.dom.setAttribute("data-textbox", "");
    this.box = document.createElement("div");
    this.box.className = "bl-textbox";
    this.contentDOM = document.createElement("div");
    this.contentDOM.className = "bl-tb-content";
    const frame = document.createElement("div");
    frame.className = "bl-tb-handles";
    frame.contentEditable = "false";
    // Taşıma kenarları (dört kenar şeridi) ve boyutlandırma tutamaçları
    for (const side of ["top", "right", "bottom", "left"]) {
      const edge = document.createElement("div");
      edge.className = `bl-tb-edge is-${side}`;
      edge.addEventListener("pointerdown", (e) => this.startMove(e));
      frame.appendChild(edge);
    }
    for (const h of HANDLES) {
      const el = document.createElement("div");
      el.className = `bl-tb-h is-${h}`;
      el.addEventListener("pointerdown", (e) => this.startResize(e, h));
      frame.appendChild(el);
    }
    this.box.append(this.contentDOM, frame);
    this.dom.appendChild(this.box);
    // Kutu seçiliyken içine tıklanınca düğüm seçimi metin imlecine döner.
    this.contentDOM.addEventListener("mousedown", (e) => {
      const sel = this.view.state.selection;
      if (!(sel instanceof NodeSelection) || sel.node !== this.node) return;
      // Tıklanan karakter konumu (kutunun altındaki metne değil, kutu içeriğine eşlenir).
      const doc = document as Document & { caretPositionFromPoint?: (x: number, y: number) => { offsetNode: globalThis.Node; offset: number } | null };
      let pos: number | null = null;
      try {
        const cp = doc.caretPositionFromPoint?.(e.clientX, e.clientY);
        const cr = cp ? null : document.caretRangeFromPoint?.(e.clientX, e.clientY);
        const node = cp?.offsetNode ?? cr?.startContainer;
        const off = cp?.offset ?? cr?.startOffset ?? 0;
        if (node && this.contentDOM.contains(node)) pos = this.view.posAtDOM(node, off);
      } catch {
        pos = null;
      }
      const start = this.getPos();
      if (start === undefined) return;
      if (pos === null || pos <= start || pos >= start + this.node.nodeSize) pos = start + this.node.nodeSize - 2;
      e.preventDefault();
      this.view.dispatch(this.view.state.tr.setSelection(TextSelection.near(this.view.state.doc.resolve(pos))));
      this.view.focus();
    });
    this.render();
  }

  private render() {
    this.box.setAttribute("style", boxStyle(this.node.attrs as TextBoxAttrs));
  }

  private scale() {
    const dom = this.view.dom as HTMLElement;
    const z = dom.offsetWidth ? dom.getBoundingClientRect().width / dom.offsetWidth : 1;
    return z || 1;
  }

  private pickNode() {
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
  }

  private drag(e: PointerEvent, onMove: (dx: number, dy: number) => void, onEnd: () => void) {
    e.preventDefault();
    e.stopPropagation();
    this.pickNode();
    const sx = e.clientX;
    const sy = e.clientY;
    const z = this.scale();
    const move = (ev: PointerEvent) => onMove((ev.clientX - sx) / z, (ev.clientY - sy) / z);
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      onEnd();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  private startMove(e: PointerEvent) {
    const a = this.node.attrs as TextBoxAttrs;
    let nx = a.x;
    let ny = a.y;
    this.drag(
      e,
      (dx, dy) => {
        nx = Math.round(a.x + dx);
        ny = Math.round(a.y + dy);
        this.box.style.left = `${nx}px`;
        this.box.style.top = `${ny}px`;
      },
      () => this.commit({ x: nx, y: ny }),
    );
  }

  private startResize(e: PointerEvent, h: (typeof HANDLES)[number]) {
    const a = this.node.attrs as TextBoxAttrs;
    let r = { x: a.x, y: a.y, w: a.w, h: a.h };
    this.drag(
      e,
      (dx, dy) => {
        r = { x: a.x, y: a.y, w: a.w, h: a.h };
        if (h.includes("e")) r.w = Math.max(24, a.w + dx);
        if (h.includes("s")) r.h = Math.max(16, a.h + dy);
        if (h.includes("w")) {
          r.w = Math.max(24, a.w - dx);
          r.x = a.x + (a.w - r.w);
        }
        if (h.includes("n")) {
          r.h = Math.max(16, a.h - dy);
          r.y = a.y + (a.h - r.h);
        }
        r = { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.w), h: Math.round(r.h) };
        Object.assign(this.box.style, { left: `${r.x}px`, top: `${r.y}px`, width: `${r.w}px`, height: `${r.h}px` });
      },
      () => this.commit(r),
    );
  }

  private commit(patch: Partial<TextBoxAttrs>) {
    const pos = this.getPos();
    if (pos === undefined) return;
    const tr = this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, ...patch });
    tr.setSelection(NodeSelection.create(tr.doc, pos));
    this.view.dispatch(tr);
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add("ProseMirror-selectednode");
  }

  deselectNode() {
    this.dom.classList.remove("ProseMirror-selectednode");
  }

  stopEvent(e: Event) {
    const t = e.target as HTMLElement;
    return !!t.closest?.(".bl-tb-handles");
  }

  ignoreMutation(m: MutationRecord | { type: "selection"; target: globalThis.Node }) {
    if (m.type === "selection") return false;
    // Kutunun stil/konum değişiklikleri ve tutamaçlar içerik değildir.
    return !this.contentDOM.contains(m.target) || (m.type === "attributes" && m.target === this.contentDOM);
  }
}
