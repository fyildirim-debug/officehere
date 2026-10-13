// Büyük İlk Harf (Ekle › Metin): paragrafın ilk harfi birkaç satır boyunca büyük yazılır. Metin içinde
// (harf metne gömülür) ya da kenar boşluğunda (harf sol boşluğa taşar, metin hizası değişmez).
// Çizim tarayıcının initial-letter özelliğiyle yapılır; kenar boşluğu kipinde harfin genişliği ölçülüp
// düğüm süslemesiyle (--dc-w) verilir. DOCX'e çerçeveli paragraf (w:framePr w:dropCap), ODT'ye
// style:drop-cap olarak yazılır.
import { Extension, type Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";
import { PARA_TYPES } from "./paragraphFormat";

export type DropCapMode = "drop" | "margin";

export interface DropCap {
  mode: DropCapMode;
  /** Bırakılacak satır sayısı (1-10). */
  lines: number;
  /** Metinden uzaklık (pt). */
  distance: number;
  /** Harfin yazı tipi (null: paragrafınki). */
  font: string | null;
}

export const DROP_CAP_DEFAULT: DropCap = { mode: "drop", lines: 3, distance: 0, font: null };

/** Öznitelik değerini doğrular (dosyadan ya da yapıştırmadan gelen bozuk değerler atılır). */
export function normalizeDropCap(v: unknown): DropCap | null {
  if (!v || typeof v !== "object") return null;
  const o = v as Partial<DropCap>;
  if (o.mode !== "drop" && o.mode !== "margin") return null;
  const lines = Math.min(10, Math.max(1, Math.round(Number(o.lines) || 3)));
  const distance = Math.min(144, Math.max(0, Number(o.distance) || 0));
  return { mode: o.mode, lines, distance: +distance.toFixed(2), font: typeof o.font === "string" && o.font ? o.font : null };
}

/** Paragrafın satır içi CSS değişkenleri. */
export function dropCapCss(dc: DropCap): string {
  return `--dc-lines:${dc.lines};--dc-dist:${dc.distance}pt${dc.font ? `;--dc-font:"${dc.font.replace(/["\\]/g, "")}"` : ""}`;
}

export const DropCapAttr = Extension.create({
  name: "dropCap",
  addGlobalAttributes() {
    return [
      {
        types: PARA_TYPES,
        attributes: {
          dropCap: {
            default: null,
            parseHTML: (el) => {
              const mode = el.getAttribute("data-dropcap");
              if (!mode) return null;
              return normalizeDropCap({ mode, lines: el.getAttribute("data-dropcap-lines"), distance: el.getAttribute("data-dropcap-dist"), font: el.getAttribute("data-dropcap-font") });
            },
            renderHTML: (a) => {
              const dc = normalizeDropCap(a.dropCap);
              if (!dc) return {};
              return {
                "data-dropcap": dc.mode,
                "data-dropcap-lines": String(dc.lines),
                "data-dropcap-dist": String(dc.distance),
                ...(dc.font ? { "data-dropcap-font": dc.font } : {}),
                style: dropCapCss(dc),
              };
            },
          },
        },
      },
    ];
  },
});

/** Kenar boşluğu kipindeki harf genişliği ölçümü (yalnız ekranda çizilen düzenleyicilerde). */
export const DropCapMeasure = Extension.create({
  name: "dropCapMeasure",
  addProseMirrorPlugins() {
    return [marginMeasure()];
  },
});

/* ------------------------------------------------------------------ */
/* Kenar boşluğu kipi: harf genişliği ölçümü                            */
/* ------------------------------------------------------------------ */

const measureKey = new PluginKey<{ widths: Map<number, number>; deco: DecorationSet }>("belgeDropCapMeasure");

/** Kenar boşluğu kipindeki paragrafların konumları. */
function marginParas(doc: PMNode): number[] {
  const out: number[] = [];
  doc.descendants((n, pos) => {
    if (n.isTextblock) {
      if ((n.attrs.dropCap as DropCap | null)?.mode === "margin") out.push(pos);
      return false;
    }
    return true;
  });
  return out;
}

/** İlk harfin (::first-letter) çizilen genişliği (px). */
function letterWidth(view: EditorView, pos: number): number | null {
  const dom = view.nodeDOM(pos);
  if (!(dom instanceof HTMLElement)) return null;
  const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
  let t = walker.nextNode() as Text | null;
  while (t && !t.data.trim()) t = walker.nextNode() as Text | null;
  if (!t) return null;
  const range = document.createRange();
  const at = t.data.search(/\S/);
  range.setStart(t, at);
  range.setEnd(t, Math.min(t.data.length, at + (t.data.codePointAt(at)! > 0xffff ? 2 : 1)));
  const w = range.getBoundingClientRect().width;
  return w > 0 ? w : null;
}

function marginMeasure() {
  return new Plugin<{ widths: Map<number, number>; deco: DecorationSet }>({
    key: measureKey,
    state: {
      init: () => ({ widths: new Map(), deco: DecorationSet.empty }),
      apply(tr, prev, _old, state: EditorState) {
        const meta = tr.getMeta(measureKey) as Map<number, number> | undefined;
        if (meta) {
          const deco = DecorationSet.create(
            state.doc,
            [...meta].map(([pos, w]) => {
              const n = state.doc.nodeAt(pos);
              const dc = n?.attrs.dropCap as DropCap | null;
              return Decoration.node(pos, pos + (n?.nodeSize ?? 1), { style: `--dc-w:${(w + (dc?.distance ?? 0) * (96 / 72)).toFixed(2)}px` });
            }),
          );
          return { widths: meta, deco };
        }
        if (!tr.docChanged) return prev;
        return { widths: prev.widths, deco: prev.deco.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => measureKey.getState(state)?.deco,
    },
    view: (view) => {
      let frame = 0;
      const measure = () => {
        frame = 0;
        if (view.isDestroyed) return;
        const st = measureKey.getState(view.state)!;
        const next = new Map<number, number>();
        for (const pos of marginParas(view.state.doc)) {
          const w = letterWidth(view, pos);
          if (w !== null) next.set(pos, Math.round(w * 4) / 4);
        }
        const same = next.size === st.widths.size && [...next].every(([k, v]) => st.widths.get(k) === v);
        if (!same) view.dispatch(view.state.tr.setMeta(measureKey, next).setMeta("addToHistory", false));
      };
      const schedule = () => {
        if (!frame) frame = requestAnimationFrame(measure);
      };
      schedule();
      return {
        update: (v, prevState) => {
          if (v.state.doc !== prevState.doc) schedule();
        },
        destroy: () => cancelAnimationFrame(frame),
      };
    },
  });
}

/* ------------------------------------------------------------------ */
/* Komutlar                                                             */
/* ------------------------------------------------------------------ */

/** Büyük ilk harf alamayan kapsayıcılar (Word'deki gibi listeler, tablolar ve metin kutuları). */
const NO_DROP = new Set(["listItem", "tableCell", "tableHeader", "textBox"]);

/** İmleçteki paragraf (büyük ilk harf uygulanabilir: liste, tablo ve metin kutusu dışındaki paragraf/başlık). */
export function dropCapTarget(state: EditorState): { node: PMNode; pos: number } | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if (NO_DROP.has($from.node(d).type.name)) return null;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (PARA_TYPES.includes(n.type.name)) return { node: n, pos: $from.before(d) };
  }
  return null;
}

/** Paragraf büyük ilk harf alabilir mi: ilk satır içi düğüm boş olmayan metin olmalı. */
export function canDropCap(node: PMNode): boolean {
  const first = node.firstChild;
  return !!first && first.isText && !!first.text?.trim();
}

/** İmleçteki paragrafa büyük ilk harf uygular (null: kaldırır). */
export function setDropCap(editor: Editor, dc: DropCap | null): boolean {
  const t = dropCapTarget(editor.state);
  if (!t) return false;
  const value = dc ? normalizeDropCap(dc) : null;
  if (value && !canDropCap(t.node)) return false;
  editor.view.dispatch(editor.state.tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, dropCap: value }).scrollIntoView());
  editor.commands.focus();
  return true;
}

/** Metnin ilk harfi (birleşik işaretler ve vekil çiftlerle). */
export function firstGrapheme(text: string): string {
  const m = /^\s*(\P{M}\p{M}*)/u.exec(text);
  return m ? text.slice(0, m.index + m[0].length) : "";
}
