// Yazdırma Düzeni sayfalaması: blokların doğal konumları ölçülür, sayfa sınırını aşan
// bloğun (ya da satırın / tablo satırının) önüne sayfa boşluğu kadar yükseklikte bir
// dekorasyon parçacığı konur. İçerik tek bir akışta kalır; sayfa zeminleri arkada çizilir.
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

export interface PageGeometry {
  enabled: boolean;
  /** Sayfa yüksekliği (px). */
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  /** Sayfalar arası boşluk (px). */
  gap: number;
  /** Sütun sayısı; > 1 ise bloklar sütunlara dönüştürülerek yerleştirilir. */
  columns: number;
  columnWidth: number;
  columnGap: number;
}

export interface LayoutInfo {
  pageCount: number;
  /** Her sayfanın ilk içerik konumu. */
  pageStarts: number[];
  /** Her sayfanın altında gösterilecek dipnot kimlikleri. */
  footnotes: string[][];
}

/** Dipnot alanı ayıracı (çizgi + boşluk) yüksekliği (px). */
export const FOOTNOTE_SEP = 14;

export interface PageBreakMark {
  pos: number;
  kind: "block" | "row" | "line";
  h: number;
  /** Sütun düzeninde bloğa uygulanacak öteleme. */
  dx?: number;
  dy?: number;
}

interface PagState {
  breaks: PageBreakMark[];
  shifts: Shift[];
  deco: DecorationSet;
  remeasure: number;
}

interface Shift {
  pos: number;
  dx: number;
  dy: number;
}

export const paginationKey = new PluginKey<PagState>("belgePagination");

/** React tarafıyla paylaşılan denetleyici: geometri ve düzen bilgisi. */
export class PaginationController {
  geometry: PageGeometry = { enabled: false, pageHeight: 1123, marginTop: 94, marginBottom: 94, gap: 16, columns: 1, columnWidth: 600, columnGap: 47 };
  info: LayoutInfo = { pageCount: 1, pageStarts: [0], footnotes: [[]] };
  view: EditorView | null = null;
  private listeners = new Set<(i: LayoutInfo) => void>();
  measureNow: (() => void) | null = null;

  setGeometry(g: Partial<PageGeometry>) {
    const next = { ...this.geometry, ...g };
    if (JSON.stringify(next) === JSON.stringify(this.geometry)) return;
    this.geometry = next;
    this.request();
  }

  request() {
    const v = this.view;
    if (v && !v.isDestroyed) v.dispatch(v.state.tr.setMeta(paginationKey, { remeasure: true }).setMeta("addToHistory", false));
  }

  subscribe(fn: (i: LayoutInfo) => void) {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  publish(info: LayoutInfo) {
    if (info.pageCount === this.info.pageCount && info.pageStarts.join() === this.info.pageStarts.join() && JSON.stringify(info.footnotes) === JSON.stringify(this.info.footnotes)) return;
    this.info = info;
    this.listeners.forEach((l) => l(info));
  }

  /** Konumun bulunduğu sayfa (0 tabanlı). */
  pageOf(pos: number): number {
    const s = this.info.pageStarts;
    let lo = 0;
    let hi = s.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (s[mid] <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

function makeWidget(b: PageBreakMark) {
  if (b.kind === "row") {
    return () => {
      const tr = document.createElement("tr");
      tr.className = "bl-pg bl-pg-row";
      const td = document.createElement("td");
      td.colSpan = 99;
      td.style.height = `${b.h}px`;
      tr.appendChild(td);
      return tr;
    };
  }
  if (b.kind === "line") {
    return () => {
      const s = document.createElement("span");
      s.className = "bl-pg bl-pg-line";
      s.style.height = `${b.h}px`;
      return s;
    };
  }
  return () => {
    const d = document.createElement("div");
    d.className = "bl-pg";
    d.style.height = `${b.h}px`;
    return d;
  };
}

function buildDeco(doc: PMNode, breaks: PageBreakMark[], shifts: Shift[]): DecorationSet {
  const decos: Decoration[] = [];
  for (const b of breaks) {
    if (b.pos <= 0 || b.pos > doc.content.size) continue;
    decos.push(Decoration.widget(b.pos, makeWidget(b), { side: -1, key: `pg:${b.kind}:${b.pos}:${b.h.toFixed(1)}`, ignoreSelection: true }));
  }
  for (const s of shifts) {
    const node = doc.nodeAt(s.pos);
    if (!node) continue;
    decos.push(Decoration.node(s.pos, s.pos + node.nodeSize, { style: `transform:translate(${s.dx.toFixed(2)}px,${s.dy.toFixed(2)}px)` }));
  }
  return DecorationSet.create(doc, decos);
}

interface Unit {
  pos: number;
  node: PMNode;
  el: HTMLElement;
  breakPos: number;
  kind: "block" | "row";
  top: number;
  bottom: number;
  keepNext: boolean;
  pageBreak: boolean;
  breakBefore: boolean;
  /** Sütun düzeninde ötelenecek düğüm konumu. */
  shiftPos: number;
  /** Birimdeki dipnot kimlikleri ve metinleri. */
  notes: { id: string; text: string }[];
}

function footnotesIn(node: PMNode): { id: string; text: string }[] {
  const out: { id: string; text: string }[] = [];
  node.descendants((n) => {
    if (n.type.name === "note" && n.attrs.kind === "footnote") out.push({ id: n.attrs.id as string, text: (n.attrs.text as string) ?? "" });
    return true;
  });
  return out;
}

const LIST_TYPES = new Set(["bulletList", "orderedList"]);

function collectUnits(view: EditorView): Omit<Unit, "top" | "bottom">[] {
  const out: Omit<Unit, "top" | "bottom">[] = [];
  const push = (node: PMNode, pos: number, breakPos: number, kind: "block" | "row", shiftPos: number) => {
    const el = view.nodeDOM(pos);
    if (!(el instanceof HTMLElement)) return;
    out.push({
      pos,
      node,
      el,
      breakPos,
      kind,
      keepNext: !!node.attrs.keepNext || node.type.name === "heading",
      pageBreak: node.type.name === "pageBreak",
      breakBefore: !!node.attrs.pageBreakBefore,
      shiftPos,
      notes: footnotesIn(node),
    });
  };
  const visit = (node: PMNode, pos: number, breakPos: number, shiftPos: number) => {
    if (LIST_TYPES.has(node.type.name)) {
      node.forEach((li, off, idx) => {
        const liPos = pos + 1 + off;
        li.forEach((c, off2, idx2) => {
          const cPos = liPos + 1 + off2;
          const bp = idx2 === 0 ? (idx === 0 ? breakPos : liPos) : cPos;
          visit(c, cPos, bp, idx2 === 0 ? liPos : shiftPos);
        });
      });
      return;
    }
    if (node.type.name === "table") {
      node.forEach((row, off, idx) => {
        const rPos = pos + 1 + off;
        push(row, rPos, idx === 0 ? breakPos : rPos, idx === 0 ? "block" : "row", pos);
      });
      return;
    }
    push(node, pos, breakPos, "block", shiftPos);
  };
  view.state.doc.forEach((child, offset) => visit(child, offset, offset, offset));
  return out;
}

/** Sütun düzeni için birimler: üst düzey bloklar, listelerde liste öğeleri. */
function collectColumnUnits(view: EditorView): Omit<Unit, "top" | "bottom">[] {
  const out: Omit<Unit, "top" | "bottom">[] = [];
  const push = (node: PMNode, pos: number) => {
    const el = view.nodeDOM(pos);
    if (!(el instanceof HTMLElement)) return;
    const first = node.isTextblock ? node : node.firstChild;
    out.push({
      pos,
      node,
      el,
      breakPos: pos,
      kind: "block",
      keepNext: node.type.name === "heading" || !!first?.attrs.keepNext,
      pageBreak: node.type.name === "pageBreak",
      breakBefore: !!first?.attrs.pageBreakBefore,
      shiftPos: pos,
      notes: footnotesIn(node),
    });
  };
  view.state.doc.forEach((child, offset) => {
    if (LIST_TYPES.has(child.type.name)) child.forEach((li, off) => push(li, offset + 1 + off));
    else push(child, offset);
  });
  return out;
}

function sameBreaks(a: PageBreakMark[], b: PageBreakMark[]) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].pos !== b[i].pos || a[i].kind !== b[i].kind || Math.abs(a[i].h - b[i].h) > 0.5) return false;
  }
  return true;
}

function sameShifts(a: Shift[], b: Shift[]) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].pos !== b[i].pos || Math.abs(a[i].dx - b[i].dx) > 0.5 || Math.abs(a[i].dy - b[i].dy) > 0.5) return false;
  }
  return true;
}

class PaginationView {
  private raf = 0;
  private timer = 0;
  private ro: ResizeObserver;
  private lastWidth = 0;
  private slow = false;

  constructor(
    private view: EditorView,
    private ctrl: PaginationController,
  ) {
    ctrl.view = view;
    ctrl.measureNow = () => this.measure();
    this.ro = new ResizeObserver(() => {
      const w = view.dom.clientWidth;
      if (Math.abs(w - this.lastWidth) > 0.5) {
        this.lastWidth = w;
        this.schedule();
      }
    });
    this.ro.observe(view.dom);
    view.dom.addEventListener("load", this.onLoad, true);
    document.fonts?.ready.then(() => this.schedule()).catch(() => {});
    document.fonts?.addEventListener?.("loadingdone", this.onLoad);
    this.schedule();
  }

  private onLoad = () => this.schedule();
  private measurer: HTMLElement | null = null;
  private noteCache = new Map<string, number>();
  private noteWidth = 0;

  /** Dipnot satırlarının yüksekliğini sayfa altındakiyle aynı biçimde ölçer. */
  private noteMeasurer(dom: HTMLElement) {
    const cs = getComputedStyle(dom);
    const width = dom.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0);
    if (Math.abs(width - this.noteWidth) > 0.5) {
      this.noteCache.clear();
      this.noteWidth = width;
    }
    return (text: string) => {
      const hit = this.noteCache.get(text);
      if (hit !== undefined) return hit;
      if (!this.measurer) {
        this.measurer = document.createElement("div");
        this.measurer.className = "bl-footnotes bl-doc is-measure";
        this.measurer.setAttribute("aria-hidden", "true");
        (dom.parentElement ?? document.body).appendChild(this.measurer);
      }
      this.measurer.style.width = `${width}px`;
      this.measurer.innerHTML = "";
      const row = document.createElement("div");
      row.className = "bl-fn-row";
      const sup = document.createElement("sup");
      sup.textContent = "99";
      const span = document.createElement("div");
      span.className = "bl-fn-text";
      span.textContent = text || "\u200b";
      row.append(sup, span);
      this.measurer.appendChild(row);
      const h = row.getBoundingClientRect().height / (dom.offsetWidth ? dom.getBoundingClientRect().width / dom.offsetWidth : 1);
      this.noteCache.set(text, h);
      return h;
    };
  }

  update(view: EditorView, prev: EditorState) {
    this.view = view;
    // Yalnız belge değiştiğinde ya da yeniden ölçüm istendiğinde ölç (kendi ara parça güncellememiz hariç).
    if (view.state.doc !== prev.doc || paginationKey.getState(view.state)?.remeasure !== paginationKey.getState(prev)?.remeasure) this.schedule();
  }

  schedule() {
    if (this.raf || this.timer) return;
    const run = () => {
      this.raf = 0;
      this.timer = 0;
      if (this.view.isDestroyed) return;
      if (this.view.composing) {
        this.timer = window.setTimeout(run, 300);
        return;
      }
      this.measure();
    };
    if (this.slow) this.timer = window.setTimeout(run, 120);
    else this.raf = requestAnimationFrame(run);
  }

  destroy() {
    cancelAnimationFrame(this.raf);
    clearTimeout(this.timer);
    this.ro.disconnect();
    this.measurer?.remove();
    this.view.dom.removeEventListener("load", this.onLoad, true);
    document.fonts?.removeEventListener?.("loadingdone", this.onLoad);
    if (this.ctrl.view === this.view) {
      this.ctrl.view = null;
      this.ctrl.measureNow = null;
    }
  }

  measure() {
    const view = this.view;
    const g = this.ctrl.geometry;
    const old = paginationKey.getState(view.state);
    if (!old) return;
    if (!g.enabled) {
      if (old.breaks.length || old.shifts.length) this.apply([], []);
      this.ctrl.publish({ pageCount: 1, pageStarts: [0], footnotes: [[]] });
      return;
    }
    const t0 = performance.now();
    const dom = view.dom as HTMLElement;
    dom.classList.add("bl-measuring");
    let breaks: PageBreakMark[] = [];
    let shifts: Shift[] = [];
    let info: LayoutInfo = { pageCount: 1, pageStarts: [0], footnotes: [[]] };
    try {
      const r = dom.getBoundingClientRect();
      const scale = dom.offsetWidth ? r.width / dom.offsetWidth : 1;
      const originY = r.top + (parseFloat(getComputedStyle(dom).paddingTop) || 0) * scale;
      const Y = (sy: number) => (sy - originY) / scale;
      const raw = g.columns > 1 ? collectColumnUnits(view) : collectUnits(view);
      const units: Unit[] = raw.map((u) => {
        const rr = u.el.getBoundingClientRect();
        return { ...u, top: Y(rr.top), bottom: Y(rr.bottom) };
      });
      const noteH = this.noteMeasurer(dom);
      const res = g.columns > 1 ? layoutColumns(units, g, Y) : layoutSingle(view, units, g, Y, originY, scale, noteH);
      breaks = res.breaks;
      shifts = res.shifts;
      info = res.info;
    } finally {
      dom.classList.remove("bl-measuring");
    }
    this.slow = performance.now() - t0 > 12;
    if (!sameBreaks(breaks, old.breaks) || !sameShifts(shifts, old.shifts)) this.apply(breaks, shifts);
    this.ctrl.publish(info);
  }

  private apply(breaks: PageBreakMark[], shifts: Shift[]) {
    const tr = this.view.state.tr.setMeta(paginationKey, { breaks, shifts }).setMeta("addToHistory", false);
    this.view.dispatch(tr);
  }
}

const marginTopOf = (el: HTMLElement) => parseFloat(getComputedStyle(el).marginTop) || 0;

/** Tek sütun: blok, tablo satırı ve paragraf satırı düzeyinde bölme; dipnotlara sayfa altında yer ayrılır. */
function layoutSingle(view: EditorView, units: Unit[], g: PageGeometry, Y: (n: number) => number, originY: number, scale: number, noteH: (text: string) => number) {
  const H = g.pageHeight - g.marginTop - g.marginBottom;
  const stride = g.pageHeight + g.gap;
  const breaks: PageBreakMark[] = [];
  const pageStarts: number[] = [0];
  const footnotes: string[][] = [[]];
  const footUsed: number[] = [0];
  const unitPage: number[] = [];
  const sBefore: number[] = [];
  let S = 0;
  let forceNext = false;

  const outerTop = (u: Unit) => {
    const el = view.nodeDOM(u.breakPos);
    if (el instanceof HTMLElement && el !== u.el) return Y(el.getBoundingClientRect().top) - marginTopOf(el);
    return u.top - marginTopOf(u.el);
  };
  const setStart = (page: number, pos: number) => {
    for (let p = pageStarts.length; p < page; p++) pageStarts[p] = pageStarts[p - 1];
    pageStarts[page] = pos;
  };
  const notesH = (u: Unit) => u.notes.reduce((sum, n) => sum + noteH(n.text), 0);
  const record = (k: number, page: number) => {
    const u = units[k];
    unitPage[k] = page;
    if (!u.notes.length) return;
    while (footnotes.length <= page) {
      footnotes.push([]);
      footUsed.push(0);
    }
    footUsed[page] += notesH(u) + (footnotes[page].length ? 0 : FOOTNOTE_SEP);
    for (const n of u.notes) footnotes[page].push(n.id);
  };
  const unrecord = (k: number) => {
    const u = units[k];
    const page = unitPage[k];
    if (page === undefined || !u.notes.length || !footnotes[page]) return;
    footnotes[page] = footnotes[page].filter((id) => !u.notes.some((n) => n.id === id));
    footUsed[page] = footnotes[page].length ? Math.max(0, footUsed[page] - notesH(u)) : 0;
  };

  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    sBefore[i] = S;
    const top = u.top + S;
    const bottom = u.bottom + S;
    const page = Math.max(0, Math.floor((top + 0.5) / stride));
    const pageTop = page * stride;
    const ownNotes = u.notes.length ? notesH(u) + ((footnotes[page]?.length ?? 0) ? 0 : FOOTNOTE_SEP) : 0;
    const contentEnd = pageTop + H - (footUsed[page] ?? 0) - ownNotes;
    // Blok sayfanın alt kenar boşluğunda başlıyorsa sonraki sayfaya taşınmalı.
    const inMargin = top > contentEnd - 0.5;
    const oTop = () => outerTop(u) + S;
    const atTop = () => oTop() <= pageTop + 1;
    const force = forceNext || u.breakBefore;
    forceNext = u.pageBreak;

    if (!(force || inMargin || bottom > contentEnd + 0.5) || atTop()) {
      record(i, page);
      continue;
    }

    // Paragrafı satır düzeyinde bölmeyi dene (dipnot içermeyenlerde).
    if (!force && !inMargin && !u.notes.length && u.kind === "block" && u.node.isTextblock && !u.node.attrs.keepLines && u.node.content.size > 0) {
      const split = splitParagraph(view, u, S, page, g, Y, originY, scale, footUsed[page] ?? 0);
      if (split) {
        for (const b of split.breaks) breaks.push(b);
        split.breaks.forEach((b, k) => setStart(page + 1 + k, b.pos));
        S += split.added;
        record(i, page);
        continue;
      }
    }

    // keepNext zinciri: başlık bir sonraki paragrafla birlikte kalsın.
    let target = i;
    if (!force) {
      let j = i - 1;
      while (j >= 0 && units[j].keepNext && !units[j].pageBreak) {
        const jt = units[j].top + sBefore[j];
        if (jt < pageTop) break;
        j--;
      }
      const cand = j + 1;
      if (cand < i) {
        const ct = outerTop(units[cand]) + sBefore[cand];
        if (ct > pageTop + 1) target = cand;
      }
    }
    for (let k = target; k < i; k++) unrecord(k);
    S = sBefore[target];
    const t = units[target];
    const h = (page + 1) * stride - (outerTop(t) + S);
    if (h <= 0) {
      record(i, page);
      continue;
    }
    // Aynı konumda bir önceki ara parça varsa birleştir.
    const existing = breaks.find((b) => b.pos === t.breakPos);
    if (existing) existing.h += h;
    else breaks.push({ pos: t.breakPos, kind: t.kind === "row" ? "row" : "block", h });
    S += h;
    for (let k = target; k <= i; k++) {
      sBefore[k] = S;
      record(k, page + 1);
    }
    setStart(page + 1, t.pos);
  }

  breaks.sort((a, b) => a.pos - b.pos);
  const last = units[units.length - 1];
  const lastBottom = last ? last.bottom + S : 0;
  const pageCount = Math.max(pageStarts.length, footnotes.length, Math.floor(Math.max(0, lastBottom - 1) / stride) + 1);
  for (let p = pageStarts.length; p < pageCount; p++) pageStarts[p] = pageStarts[p - 1];
  for (let p = footnotes.length; p < pageCount; p++) footnotes[p] = [];
  return { breaks, shifts: [] as Shift[], info: { pageCount, pageStarts, footnotes } };
}

/** Uzun paragrafı satır sınırından böler (dul/öksüz satır denetimiyle). */
function splitParagraph(view: EditorView, u: Unit, S0: number, page0: number, g: PageGeometry, Y: (n: number) => number, originY: number, scale: number, reserved = 0) {
  const H = g.pageHeight - g.marginTop - g.marginBottom;
  const stride = g.pageHeight + g.gap;
  const start = u.pos + 1;
  const end = u.pos + u.node.nodeSize - 1;
  if (end - start < 2) return null;
  const coords = (p: number, side = 1) => {
    try {
      return view.coordsAtPos(p, side);
    } catch {
      return null;
    }
  };
  const first = coords(start);
  if (!first) return null;
  const cs = getComputedStyle(u.el);
  const lhRaw = parseFloat(cs.lineHeight);
  const charH = Math.max(1, first.bottom - first.top);
  const lineH = Number.isFinite(lhRaw) ? lhRaw * scale : charH * 1.2;
  const halfLead = Math.max(0, (lineH - charH) / 2);
  /** bottom > y olan ilk konum. */
  const firstPosBelow = (lo: number, y: number) => {
    let hi = end;
    let ans = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const c = coords(mid);
      if (c && c.bottom > y) {
        ans = mid;
        hi = mid - 1;
      } else lo = mid + 1;
    }
    return ans;
  };
  const paraBottomScreen = originY + u.bottom * scale;
  const breaks: PageBreakMark[] = [];
  let S = S0;
  let page = page0;
  let from = start;
  let lineTopFirst = first.top;
  for (let guard = 0; guard < 200; guard++) {
    const contentEnd = page * stride + H - (page === page0 ? reserved : 0);
    if (u.bottom + S <= contentEnd + 0.5) break;
    const threshold = originY + (contentEnd - S) * scale;
    const p = firstPosBelow(from, threshold);
    if (p < 0) break;
    const pc = coords(p);
    if (!pc) break;
    let T = pc.top;
    let q = firstPosBelow(from, T + 1);
    if (q < 0) break;
    // Tek satır sonraki sayfaya kalıyorsa bir satır geri al.
    if (paraBottomScreen - T < lineH * 1.5 && q > from) {
      const prev = coords(q - 1, -1);
      if (prev) {
        const q2 = firstPosBelow(from, prev.top + 1);
        if (q2 >= from && q2 < q) {
          q = q2;
          T = coords(q)?.top ?? T;
        }
      }
    }
    // Sayfada tek satır kalacaksa bölme.
    if (T - lineTopFirst < lineH * 1.5 || q <= from) {
      if (breaks.length === 0) return null;
      break;
    }
    const lineTop = Y(T - halfLead) + S;
    const h = (page + 1) * stride - lineTop;
    if (h <= 0) break;
    breaks.push({ pos: q, kind: "line", h });
    S += h;
    page += 1;
    from = q;
    lineTopFirst = T;
  }
  if (!breaks.length) return null;
  return { breaks, added: S - S0 };
}

/**
 * Çok sütunlu düzen: her sütun sanal bir sayfa gibi doldurulur, bloklar
 * transform ile sayfadaki sütun konumuna taşınır (paragraf bölünmez).
 */
function layoutColumns(units: Unit[], g: PageGeometry, _Y: (n: number) => number) {
  const H = g.pageHeight - g.marginTop - g.marginBottom;
  const stride = g.pageHeight + g.gap;
  const n = g.columns;
  const shifts: Shift[] = [];
  const pageStarts: number[] = [0];
  let col = 0;
  let base = 0; // geçerli sütunun doğal akıştaki başlangıcı
  let forceNext = false;
  for (const u of units) {
    const top = u.top - base;
    const bottom = u.bottom - base;
    const force = forceNext || u.breakBefore;
    forceNext = u.pageBreak;
    if ((force || bottom > H + 0.5) && top > 1) {
      col = force ? (Math.floor(col / n) + 1) * n : col + 1;
      base = u.top;
      if (col % n === 0) pageStarts[col / n] = u.pos;
    }
    const pageIdx = Math.floor(col / n);
    const dx = (col % n) * (g.columnWidth + g.columnGap);
    const dy = pageIdx * stride - base;
    if (Math.abs(dx) > 0.01 || Math.abs(dy) > 0.01) shifts.push({ pos: u.shiftPos, dx, dy });
  }
  const pageCount = Math.floor(col / n) + 1;
  for (let p = 1; p < pageCount; p++) if (pageStarts[p] === undefined) pageStarts[p] = pageStarts[p - 1];
  return { breaks: [] as PageBreakMark[], shifts, info: { pageCount, pageStarts, footnotes: Array.from({ length: pageCount }, () => [] as string[]) } };
}

export const Pagination = Extension.create<{ controller: PaginationController }>({
  name: "belgePagination",
  addOptions() {
    return { controller: new PaginationController() };
  },
  addProseMirrorPlugins() {
    const ctrl = this.options.controller;
    return [
      new Plugin<PagState>({
        key: paginationKey as never,
        state: {
          init: () => ({ breaks: [], shifts: [], deco: DecorationSet.empty, remeasure: 0 }),
          apply(tr, value, _old, newState) {
            const meta = tr.getMeta(paginationKey) as { breaks?: PageBreakMark[]; shifts?: Shift[]; remeasure?: boolean } | undefined;
            if (meta?.breaks) {
              const shifts = meta.shifts ?? [];
              return { breaks: meta.breaks, shifts, deco: buildDeco(newState.doc, meta.breaks, shifts), remeasure: value.remeasure };
            }
            let next = value;
            if (tr.docChanged) {
              const breaks = value.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1) }));
              const shifts = value.shifts.map((s) => ({ ...s, pos: tr.mapping.map(s.pos, 1) }));
              next = { ...value, breaks, shifts, deco: value.deco.map(tr.mapping, tr.doc) };
            }
            if (meta?.remeasure) next = { ...next, remeasure: next.remeasure + 1 };
            return next;
          },
        },
        props: {
          decorations(state) {
            return paginationKey.getState(state)?.deco;
          },
        },
        view: (view) => new PaginationView(view, ctrl),
      }),
    ];
  },
});
