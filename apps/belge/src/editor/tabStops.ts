// Sekme durakları (Biçim › Sekmeler, cetvel, DOCX w:tabs): paragrafa özel sol/orta/sağ/ondalık/çubuk
// duraklar ve öncü karakterler. Paragraftaki her sekme karakteri ölçülerek bir sonraki durağa uzatılır.
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

export type TabAlign = "left" | "center" | "right" | "decimal" | "bar";
export type TabLeader = "none" | "dot" | "hyphen" | "underscore";

export interface TabStop {
  /** Sol kenar boşluğundan uzaklık (pt). */
  pos: number;
  align: TabAlign;
  leader: TabLeader;
}

/** Varsayılan sekme durağı aralığı: 1,25 cm (pt). */
export const DEFAULT_TAB = 35.4;
const PT = 4 / 3;

export const TAB_ALIGN_LABEL: Record<TabAlign, string> = { left: "Sol", center: "Orta", right: "Sağ", decimal: "Ondalık", bar: "Çubuk" };

export function normalizeTabs(tabs: TabStop[]): TabStop[] {
  const seen = new Map<number, TabStop>();
  for (const t of tabs) if (Number.isFinite(t.pos) && t.pos >= 0) seen.set(Math.round(t.pos * 10) / 10, { ...t, pos: Math.round(t.pos * 10) / 10 });
  return [...seen.values()].sort((a, b) => a.pos - b.pos);
}

/** Paragraf özniteliği: tabs (yalnız özel duraklar; yoksa null). */
export const TabStopsAttr = Extension.create({
  name: "belgeTabStops",
  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          tabs: {
            default: null,
            parseHTML: (el) => {
              const raw = el.getAttribute("data-tabs");
              if (!raw) return null;
              try {
                return normalizeTabs(JSON.parse(raw) as TabStop[]);
              } catch {
                return null;
              }
            },
            renderHTML: (a) => (a.tabs && (a.tabs as TabStop[]).length ? { "data-tabs": JSON.stringify(a.tabs) } : {}),
          },
          outlineLevel: {
            default: null,
            parseHTML: (el) => {
              const v = el.getAttribute("data-outline");
              return v ? Number(v) : null;
            },
            renderHTML: (a) => (a.outlineLevel ? { "data-outline": String(a.outlineLevel) } : {}),
          },
        },
      },
    ];
  },
  addProseMirrorPlugins() {
    return [tabLayoutPlugin()];
  },
});

interface TabWidth {
  pos: number;
  width: number;
  leader: TabLeader;
}

const tabKey = new PluginKey<{ widths: TabWidth[]; deco: DecorationSet }>("belgeTabLayout");

function buildDeco(doc: PMNode, widths: TabWidth[]): DecorationSet {
  const decos: Decoration[] = [];
  for (const w of widths) {
    if (w.pos < 0 || w.pos + 1 > doc.content.size) continue;
    decos.push(Decoration.inline(w.pos, w.pos + 1, { class: `bl-tab${w.leader !== "none" ? ` is-${w.leader}` : ""}`, style: `width:${w.width.toFixed(2)}px` }));
  }
  return DecorationSet.create(doc, decos);
}

/** Sekmenin bitiş konumu ve sonrasındaki metin parçası (sonraki sekmeye ya da satır sonuna dek) genişliği. */
function segmentWidth(view: EditorView, from: number, to: number, decimal: boolean): number {
  if (to <= from) return 0;
  try {
    const a = view.domAtPos(from);
    const b = view.domAtPos(to);
    const r = document.createRange();
    r.setStart(a.node, a.offset);
    r.setEnd(b.node, b.offset);
    const rects = Array.from(r.getClientRects()).filter((x) => x.width > 0);
    if (!rects.length) return 0;
    // Yalnız ilk satırdaki parça.
    const top = rects[0].top;
    const line = rects.filter((x) => Math.abs(x.top - top) < 2);
    void decimal;
    return Math.max(...line.map((x) => x.right)) - Math.min(...line.map((x) => x.left));
  } catch {
    return 0;
  }
}

/** Belgedeki özel duraklı paragrafların sekme genişliklerini ölçer. */
function measure(view: EditorView, scale: number): TabWidth[] {
  const out: TabWidth[] = [];
  view.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const tabs = node.attrs.tabs as TabStop[] | null;
    if (!tabs?.length || !node.textContent.includes("\t")) return false;
    const el = view.nodeDOM(pos) as HTMLElement | null;
    if (!el) return false;
    const cs = getComputedStyle(el);
    // Duraklar paragrafın sol kenarından (sol girinti dahil değil) ölçülür.
    const originX = el.getBoundingClientRect().left + (parseFloat(cs.borderLeftWidth) || 0) * scale;
    const text = node.textBetween(0, node.content.size, undefined, "￼");
    const starts: number[] = [];
    for (let i = 0; i < text.length; i++) if (text[i] === "\t") starts.push(i);
    starts.forEach((off, k) => {
      const p = pos + 1 + off;
      let x = 0;
      try {
        x = (view.coordsAtPos(p, 1).left - originX) / scale;
      } catch {
        return;
      }
      const xPt = x / PT;
      const stop = tabs.find((t) => t.pos > xPt + 0.5 && t.align !== "bar");
      const nextOff = k + 1 < starts.length ? starts[k + 1] : text.length;
      let width: number;
      let leader: TabLeader = "none";
      if (stop) {
        const stopX = stop.pos * PT;
        leader = stop.leader;
        if (stop.align === "left") width = stopX - x;
        else {
          let seg = segmentWidth(view, p + 1, pos + 1 + nextOff, false) / scale;
          if (stop.align === "decimal") {
            const rest = text.slice(off + 1, nextOff);
            const d = rest.search(/[,.]/);
            if (d >= 0) seg = segmentWidth(view, p + 1, p + 1 + d, true) / scale;
          }
          width = stop.align === "center" ? stopX - x - seg / 2 : stopX - x - seg;
        }
      } else {
        const def = DEFAULT_TAB * PT;
        width = (Math.floor(x / def) + 1) * def - x;
      }
      out.push({ pos: p, width: Math.max(2, width), leader });
    });
    return false;
  });
  return out;
}

function sameWidths(a: TabWidth[], b: TabWidth[]) {
  return a.length === b.length && a.every((w, i) => w.pos === b[i].pos && Math.abs(w.width - b[i].width) < 0.5 && w.leader === b[i].leader);
}

function tabLayoutPlugin() {
  return new Plugin<{ widths: TabWidth[]; deco: DecorationSet }>({
    key: tabKey,
    state: {
      init: () => ({ widths: [], deco: DecorationSet.empty }),
      apply(tr, value, _o, state) {
        const meta = tr.getMeta(tabKey) as TabWidth[] | undefined;
        if (meta) return { widths: meta, deco: buildDeco(state.doc, meta) };
        if (!tr.docChanged) return value;
        const widths = value.widths.map((w) => ({ ...w, pos: tr.mapping.map(w.pos) }));
        return { widths, deco: value.deco.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => tabKey.getState(state)?.deco,
    },
    view(view) {
      let raf = 0;
      let rounds = 0;
      const run = () => {
        raf = 0;
        if (view.isDestroyed) return;
        const host = view.dom.closest(".bl-canvas") as HTMLElement | null;
        const scale = host ? parseFloat(getComputedStyle(host).zoom) || 1 : 1;
        const next = measure(view, scale);
        const cur = tabKey.getState(view.state)?.widths ?? [];
        if (!sameWidths(next, cur) && rounds++ < 4) {
          view.dispatch(view.state.tr.setMeta(tabKey, next).setMeta("addToHistory", false));
          raf = requestAnimationFrame(run);
        }
      };
      const schedule = () => {
        if (!raf) raf = requestAnimationFrame(run);
      };
      schedule();
      return {
        update(v, prev) {
          if (v.state.doc !== prev.doc) {
            rounds = 0;
            schedule();
          }
        },
        destroy() {
          cancelAnimationFrame(raf);
        },
      };
    },
  });
}
