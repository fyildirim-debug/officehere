// Grafik (Ekle › Grafik): sütun, çubuk, çizgi ve pasta grafikleri SVG olarak çizilir. Satır içi düğümdür
// (Word'de grafik varsayılan olarak metinle aynı hizadadır). Çift tıklama verileri düzenleme kutusunu açar.
// DOCX'te gömülü grafik parçası (c:chartSpace, değer önbellekli) olarak yazılır ve okunur.
import { Node, mergeAttributes } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import { closeHistory } from "@tiptap/pm/history";

export type ChartType = "column" | "bar" | "line" | "pie";

export interface ChartSeries {
  name: string;
  values: number[];
  color: string;
}

export interface ChartData {
  type: ChartType;
  title: string;
  categories: string[];
  series: ChartSeries[];
  width: number;
  height: number;
}

export const CHART_TYPE_LABEL: Record<ChartType, string> = { column: "Sütun", bar: "Çubuk", line: "Çizgi", pie: "Pasta" };

/** Varsayılan veri (yeni grafik). */
export function defaultChart(type: ChartType, accents: string[]): ChartData {
  return {
    type,
    title: "Grafik Başlığı",
    categories: ["Kategori 1", "Kategori 2", "Kategori 3", "Kategori 4"],
    series: [
      { name: "Seri 1", values: [4.3, 2.5, 3.5, 4.5], color: accents[0] ?? "#156082" },
      { name: "Seri 2", values: [2.4, 4.4, 1.8, 2.8], color: accents[1] ?? "#E97132" },
      { name: "Seri 3", values: [2, 2, 3, 5], color: accents[2] ?? "#196B24" },
    ],
    width: 576,
    height: 336,
  };
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const fmt = (v: number) => (Math.abs(v) >= 1000 ? v.toLocaleString("tr-TR", { maximumFractionDigits: 0 }) : v.toLocaleString("tr-TR", { maximumFractionDigits: 2 }));

/** Eksen için "yuvarlak" adım ve sınırlar. */
export function niceScale(min: number, max: number, ticks = 5): { lo: number; hi: number; step: number } {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { lo: 0, hi: 1, step: 0.2 };
  const lo0 = Math.min(0, min);
  const hi0 = Math.max(0, max);
  const span = hi0 - lo0 || 1;
  const raw = span / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  return { lo: Math.floor(lo0 / step) * step, hi: Math.ceil(hi0 / step) * step, step };
}

/** Grafiğin SVG'si (genişlik × yükseklik px). `mini`: küçük önizleme (yazısız). */
export function chartSvg(c: ChartData, mini = false): string {
  if (mini) return miniSvg(c);
  const W = Math.max(120, c.width);
  const H = Math.max(90, c.height);
  const font = "font-family:Aptos,Calibri,Arial,sans-serif";
  const parts: string[] = [`<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="#fff" stroke="#d9d9d9"/>`];
  const titleH = c.title ? 30 : 8;
  if (c.title) parts.push(`<text x="${W / 2}" y="22" text-anchor="middle" style="${font};font-size:15px" fill="#595959">${esc(c.title)}</text>`);
  // Gösterge (altta)
  const legendItems = c.type === "pie" ? c.categories.map((n, i) => ({ name: n, color: pieColor(c, i) })) : c.series.map((s) => ({ name: s.name, color: s.color }));
  const legendH = 24;
  const lw = legendItems.reduce((sum, it) => sum + 22 + it.name.length * 6.5, 0);
  let lx = Math.max(8, (W - lw) / 2);
  for (const it of legendItems) {
    parts.push(`<rect x="${lx}" y="${H - 17}" width="9" height="9" fill="${it.color}"/><text x="${lx + 13}" y="${H - 9}" style="${font};font-size:11px" fill="#595959">${esc(it.name)}</text>`);
    lx += 22 + it.name.length * 6.5;
  }
  const top = titleH + 6;
  const bottom = H - legendH - 8;
  if (c.type === "pie") {
    const s = c.series[0];
    const vals = (s?.values ?? []).map((v) => Math.max(0, v || 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const r = Math.max(10, Math.min(W - 40, bottom - top) / 2);
    const cx = W / 2;
    const cy = (top + bottom) / 2;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const p = (a: number) => `${(cx + r * Math.cos(a)).toFixed(2)},${(cy + r * Math.sin(a)).toFixed(2)}`;
      parts.push(v >= total ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${pieColor(c, i)}" stroke="#fff"/>` : `<path d="M${cx},${cy} L${p(a0)} A${r},${r} 0 ${large} 1 ${p(a1)} Z" fill="${pieColor(c, i)}" stroke="#fff" stroke-width="1.5"/>`);
      a0 = a1;
    });
    return wrapSvg(W, H, parts);
  }
  const all = c.series.flatMap((s) => s.values.filter((v) => Number.isFinite(v)));
  const { lo, hi, step } = niceScale(Math.min(...all, 0), Math.max(...all, 0));
  const horizontal = c.type === "bar";
  const left = horizontal ? 8 + Math.min(110, Math.max(...c.categories.map((x) => x.length), 4) * 6.5) : 8 + Math.max(...[lo, hi].map((v) => fmt(v).length)) * 7 + 6;
  const right = W - 14;
  const plotB = horizontal ? bottom - 16 : bottom - 18;
  const n = Math.max(1, c.categories.length);
  const valueAt = (v: number) => (horizontal ? left + ((v - lo) / (hi - lo || 1)) * (right - left) : plotB - ((v - lo) / (hi - lo || 1)) * (plotB - top));
  // Kılavuz çizgileri ve değer ekseni etiketleri
  for (let v = lo; v <= hi + step / 2; v += step) {
    const p = valueAt(v);
    if (horizontal) parts.push(`<line x1="${p}" y1="${top}" x2="${p}" y2="${plotB}" stroke="#d9d9d9"/><text x="${p}" y="${plotB + 13}" text-anchor="middle" style="${font};font-size:10px" fill="#595959">${fmt(+v.toFixed(10))}</text>`);
    else parts.push(`<line x1="${left}" y1="${p}" x2="${right}" y2="${p}" stroke="#d9d9d9"/><text x="${left - 5}" y="${p + 3.5}" text-anchor="end" style="${font};font-size:10px" fill="#595959">${fmt(+v.toFixed(10))}</text>`);
  }
  const band = horizontal ? (plotB - top) / n : (right - left) / n;
  c.categories.forEach((cat, i) => {
    const mid = (horizontal ? top : left) + band * (i + 0.5);
    if (horizontal) parts.push(`<text x="${left - 5}" y="${mid + 3.5}" text-anchor="end" style="${font};font-size:10px" fill="#595959">${esc(cat)}</text>`);
    else parts.push(`<text x="${mid}" y="${plotB + 14}" text-anchor="middle" style="${font};font-size:10px" fill="#595959">${esc(cat)}</text>`);
  });
  const zero = valueAt(Math.max(lo, Math.min(hi, 0)));
  if (c.type === "line") {
    c.series.forEach((s) => {
      const pts = s.values.map((v, i) => `${(left + band * (i + 0.5)).toFixed(2)},${valueAt(v || 0).toFixed(2)}`);
      parts.push(`<polyline points="${pts.join(" ")}" fill="none" stroke="${s.color}" stroke-width="2.25" stroke-linejoin="round"/>`);
      pts.forEach((p) => {
        const [x, y] = p.split(",");
        parts.push(`<circle cx="${x}" cy="${y}" r="3" fill="${s.color}"/>`);
      });
    });
  } else {
    const k = Math.max(1, c.series.length);
    const inner = band * 0.62;
    const w = inner / k;
    c.series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (i >= n) return;
        const start = (horizontal ? top : left) + band * i + (band - inner) / 2 + w * si;
        const p = valueAt(v || 0);
        if (horizontal) parts.push(`<rect x="${Math.min(p, zero).toFixed(2)}" y="${start.toFixed(2)}" width="${Math.abs(p - zero).toFixed(2)}" height="${(w * 0.92).toFixed(2)}" fill="${s.color}"/>`);
        else parts.push(`<rect x="${start.toFixed(2)}" y="${Math.min(p, zero).toFixed(2)}" width="${(w * 0.92).toFixed(2)}" height="${Math.abs(p - zero).toFixed(2)}" fill="${s.color}"/>`);
      });
    });
  }
  parts.push(horizontal ? `<line x1="${zero}" y1="${top}" x2="${zero}" y2="${plotB}" stroke="#bfbfbf"/>` : `<line x1="${left}" y1="${zero}" x2="${right}" y2="${zero}" stroke="#bfbfbf"/>`);
  return wrapSvg(W, H, parts);
}

/** Pasta dilimleri (tema vurgu renkleri sırasıyla; ilk dilim serinin rengi). */
/** Tür seçimi için yazısız küçük çizim. */
function miniSvg(c: ChartData): string {
  const W = c.width;
  const H = c.height;
  const parts: string[] = [`<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="#fff" stroke="#d9d9d9"/>`];
  const all = c.series.flatMap((s) => s.values);
  const max = Math.max(...all, 1);
  const n = c.categories.length;
  const k = c.series.length;
  if (c.type === "pie") {
    const vals = c.series[0]?.values ?? [];
    const total = vals.reduce((a, b) => a + Math.max(0, b), 0) || 1;
    const r = Math.min(W, H) / 2 - 5;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (Math.max(0, v) / total) * Math.PI * 2;
      const p = (a: number) => `${(W / 2 + r * Math.cos(a)).toFixed(1)},${(H / 2 + r * Math.sin(a)).toFixed(1)}`;
      parts.push(`<path d="M${W / 2},${H / 2} L${p(a0)} A${r},${r} 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${p(a1)} Z" fill="${pieColor(c, i)}" stroke="#fff"/>`);
      a0 = a1;
    });
  } else if (c.type === "line") {
    c.series.forEach((s) => parts.push(`<polyline points="${s.values.map((v, i) => `${(6 + ((W - 12) * (i + 0.5)) / n).toFixed(1)},${(H - 6 - ((H - 12) * v) / max).toFixed(1)}`).join(" ")}" fill="none" stroke="${s.color}" stroke-width="2"/>`));
  } else {
    const band = (c.type === "bar" ? H - 12 : W - 12) / n;
    const w = (band * 0.7) / k;
    c.series.forEach((s, si) =>
      s.values.forEach((v, i) => {
        const len = ((c.type === "bar" ? W - 12 : H - 12) * v) / max;
        const at = 6 + band * i + band * 0.15 + w * si;
        parts.push(c.type === "bar" ? `<rect x="6" y="${at.toFixed(1)}" width="${len.toFixed(1)}" height="${(w * 0.9).toFixed(1)}" fill="${s.color}"/>` : `<rect x="${at.toFixed(1)}" y="${(H - 6 - len).toFixed(1)}" width="${(w * 0.9).toFixed(1)}" height="${len.toFixed(1)}" fill="${s.color}"/>`);
      }),
    );
  }
  return wrapSvg(W, H, parts);
}

const PIE_COLORS = ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E", "#7F7F7F", "#FFC000"];
export function pieColor(c: Pick<ChartData, "series">, i: number): string {
  if (i === 0 && c.series[0]?.color) return c.series[0].color;
  return PIE_COLORS[i % PIE_COLORS.length];
}

function wrapSvg(W: number, H: number, parts: string[]): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img">${parts.join("")}</svg>`;
}

export const CHART_EVENT = "officehere:edit-chart";

const RESIZE_HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
/** Grafik boyut sınırları (px); en küçük değerler chartSvg'nin çizebildiği alt sınırdır. */
const MIN_W = 120;
const MIN_H = 90;
const MAX_H = 1400;

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    chart: {
      /** Grafik ekler; seçili grafik varsa verilerini değiştirir. */
      insertChart: (data: ChartData) => ReturnType;
    };
  }
}

export const Chart = Node.create({
  name: "chart",
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return {
      data: {
        default: null,
        parseHTML: (el) => {
          try {
            return JSON.parse(el.getAttribute("data-chart") ?? "null") as ChartData | null;
          } catch {
            return null;
          }
        },
        renderHTML: (a) => (a.data ? { "data-chart": JSON.stringify(a.data) } : {}),
      },
    };
  },
  parseHTML() {
    return [{ tag: "span[data-chart]" }];
  },
  renderHTML({ node, HTMLAttributes }) {
    const attrs = mergeAttributes(HTMLAttributes, { class: "bl-chart" });
    if (typeof document === "undefined" || !node.attrs.data) return ["span", attrs];
    const span = document.createElement("span");
    for (const [k, v] of Object.entries(attrs)) span.setAttribute(k, String(v));
    span.innerHTML = chartSvg(node.attrs.data as ChartData);
    return span;
  },
  renderText({ node }) {
    return (node.attrs.data as ChartData | null)?.title ?? "";
  },
  addNodeView() {
    return ({ node, getPos, editor }) => {
      const dom = document.createElement("span");
      dom.className = "bl-chart";
      dom.contentEditable = "false";
      const canvas = document.createElement("span");
      canvas.className = "bl-chart-svg";
      dom.appendChild(canvas);
      // Seçiliyken görünen 8 boyutlandırma tutamacı (köşeler ve kenar ortaları).
      for (const h of RESIZE_HANDLES) {
        const el = document.createElement("span");
        el.className = `bl-chart-h is-${h}`;
        el.dataset.handle = h;
        dom.appendChild(el);
      }
      let cur = node;
      const draw = (data = cur.attrs.data as ChartData | null) => {
        canvas.innerHTML = data ? chartSvg(data) : "";
      };
      draw();
      dom.addEventListener("dblclick", (e) => {
        if ((e.target as HTMLElement).closest?.(".bl-chart-h")) return;
        e.preventDefault();
        const pos = typeof getPos === "function" ? getPos() : undefined;
        window.dispatchEvent(new CustomEvent(CHART_EVENT, { detail: { pos, editor } }));
      });
      dom.addEventListener("pointerdown", (e) => {
        const handle = (e.target as HTMLElement).closest?.(".bl-chart-h") as HTMLElement | null;
        const data = cur.attrs.data as ChartData | null;
        if (!handle || !data || e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        const dir = handle.dataset.handle ?? "se";
        // Yakınlaştırmada ekrandaki piksel ile grafik pikseli farklıdır.
        const scale = canvas.getBoundingClientRect().width / data.width || 1;
        const area = dom.closest(".ProseMirror") as HTMLElement | null;
        const maxW = Math.max(MIN_W, area?.clientWidth ?? 1600);
        const { clientX: sx, clientY: sy } = e;
        const w0 = data.width;
        const h0 = data.height;
        let w = w0;
        let h = h0;
        handle.setPointerCapture(e.pointerId);
        dom.classList.add("is-resizing");
        const move = (ev: PointerEvent) => {
          const dx = (ev.clientX - sx) / scale;
          const dy = (ev.clientY - sy) / scale;
          w = w0 + (dir.includes("e") ? dx : dir.includes("w") ? -dx : 0);
          h = h0 + (dir.includes("s") ? dy : dir.includes("n") ? -dy : 0);
          // Köşe tutamacında Shift en-boy oranını korur.
          if (ev.shiftKey && dir.length === 2) {
            const r = Math.max(w / w0, h / h0);
            w = w0 * r;
            h = h0 * r;
          }
          w = Math.round(Math.min(maxW, Math.max(MIN_W, w)));
          h = Math.round(Math.min(MAX_H, Math.max(MIN_H, h)));
          draw({ ...data, width: w, height: h });
        };
        const up = () => {
          handle.removeEventListener("pointermove", move);
          handle.removeEventListener("pointerup", up);
          handle.removeEventListener("pointercancel", up);
          dom.classList.remove("is-resizing");
          const pos = typeof getPos === "function" ? getPos() : undefined;
          if (pos == null || (w === w0 && h === h0)) return draw();
          const tr = editor.state.tr.setNodeMarkup(pos, undefined, { ...cur.attrs, data: { ...data, width: w, height: h } });
          tr.setSelection(NodeSelection.create(tr.doc, pos));
          // Her boyutlandırma ayrı bir geri alma adımı olsun (önceki yazma/ekleme ile birleşmesin).
          editor.view.dispatch(closeHistory(tr));
        };
        handle.addEventListener("pointermove", move);
        handle.addEventListener("pointerup", up);
        handle.addEventListener("pointercancel", up);
      });
      return {
        dom,
        update: (n) => {
          if (n.type.name !== "chart") return false;
          cur = n;
          draw();
          return true;
        },
        selectNode: () => dom.classList.add("is-selected"),
        deselectNode: () => dom.classList.remove("is-selected"),
        // Tutamaç olaylarını düzenleyici işlemesin (seçim/sürükleme başlatmasın).
        stopEvent: (e: Event) => !!(e.target as HTMLElement | null)?.closest?.(".bl-chart-h"),
        ignoreMutation: () => true,
      };
    };
  },
  addCommands() {
    return {
      insertChart:
        (data) =>
        ({ state, dispatch }) => {
          const sel = state.selection;
          const type = state.schema.nodes.chart;
          if (dispatch) {
            if (sel instanceof NodeSelection && sel.node.type === type) {
              const tr = state.tr.setNodeMarkup(sel.from, undefined, { data });
              dispatch(tr.setSelection(NodeSelection.create(tr.doc, sel.from)));
            } else {
              const tr = state.tr.replaceSelectionWith(type.create({ data }), false);
              const at = tr.selection.from - 1;
              if (at >= 0 && tr.doc.nodeAt(at)?.type === type) tr.setSelection(NodeSelection.create(tr.doc, at));
              dispatch(tr.scrollIntoView());
            }
          }
          return true;
        },
    };
  },
});
