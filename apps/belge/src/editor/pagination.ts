// Yazdırma Düzeni sayfalaması: blokların doğal konumları ölçülür, sayfa sınırını aşan
// bloğun (ya da satırın / tablo satırının) önüne sayfa boşluğu kadar yükseklikte bir
// dekorasyon parçacığı konur. İçerik tek bir akışta kalır; sayfa zeminleri arkada çizilir.
// Her sayfa kendi bölümünün boyut ve kenar boşluklarını kullanır.
import { headerCloneWidget } from "./tableTools";
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";
import { defaultPage, type PageSetup } from "../model/document";
import { bodyWidth, geomOf, sectionGeoms, sectionLayoutKey, type SecGeom } from "./sections";
import { planTileRows, tileColumns } from "./tileRows";
import { collectSectionChildren, collectUnits, layoutFlow, lineSplitAt, type ColumnCtx, type ColumnResult, type LayoutInfo, type PageBreakMark, type PageInfo, type SectionStyle, type Unit } from "./paginate";

export type { LayoutInfo, PageInfo, PageBreakMark } from "./paginate";
export { FOOTNOTE_SEP } from "./paginate";

export interface PageGeometry {
  enabled: boolean;
  /** Sayfalar arası boşluk (px). */
  gap: number;
  /** Belge varsayılan sayfa yapısı (kendi yapısı olmayan bölümler için). */
  defaultPage: PageSetup;
  /** Birden Çok Sayfa: satır başına sayfa (0 = kapalı). */
  tiles: number;
  /** Ekran ölçeği (CSS zoom). Ölçeği öğe genişliğinden çıkarmak tamsayı yuvarlaması yüzünden uzun belgede kaymaya yol açar. */
  scale: number;
}

interface PagState {
  breaks: PageBreakMark[];
  styles: SectionStyle[];
  deco: DecorationSet;
  remeasure: number;
}

export const paginationKey = new PluginKey<PagState>("belgePagination");

/** Tek sayfalık varsayılan düzen bilgisi (Web Düzeni ya da ölçüm öncesi). */
export function singlePageInfo(page: PageSetup): LayoutInfo {
  const g = geomOf(page);
  return {
    pageCount: 1,
    pageStarts: [0],
    footnotes: [[]],
    bodyW: g.w,
    pages: [{ top: 0, left: 0, w: g.w, h: g.h, mt: g.mt, mb: g.mb, ml: g.ml, mr: g.mr, headerY: g.headerY, footerY: g.footerY, section: 0, first: true, titlePage: g.titlePage }],
  };
}

const pagesKey = (pages: PageInfo[]) => pages.map((p) => `${p.top.toFixed(1)}:${p.w.toFixed(1)}:${p.h.toFixed(1)}:${p.mt.toFixed(1)}:${p.mb.toFixed(1)}:${p.ml.toFixed(1)}:${p.mr.toFixed(1)}:${p.headerY}:${p.footerY}:${p.section}:${p.first ? 1 : 0}:${p.blank ? 1 : 0}:${p.titlePage ? 1 : 0}:${p.left.toFixed(1)}`).join("|");

/** React tarafıyla paylaşılan denetleyici: geometri ve düzen bilgisi. */
export class PaginationController {
  geometry: PageGeometry = { enabled: false, gap: 16, defaultPage: defaultPage(), tiles: 0, scale: 1 };
  info: LayoutInfo = singlePageInfo(defaultPage());
  view: EditorView | null = null;
  private listeners = new Set<(i: LayoutInfo) => void>();
  measureNow: (() => void) | null = null;
  /** Son Yazdırma Düzeni sayfalaması (Taslak/Anahat görünümünde durum çubuğu için; düzenlemelerle eşlenir). */
  lastPaged: { starts: number[]; count: number } | null = null;

  /** Taslak/Anahat görünümünde konumun sayfası ve sayfa sayısı (son sayfalamaya göre). */
  pagedPosition(pos: number): { page: number; count: number } {
    const lp = this.lastPaged;
    if (!lp) return { page: 0, count: 1 };
    let page = 0;
    for (let i = 0; i < lp.starts.length; i++) if (lp.starts[i] <= pos) page = i;
    return { page, count: lp.count };
  }

  setGeometry(g: Partial<PageGeometry>) {
    const next = { ...this.geometry, ...g };
    if (JSON.stringify(next) === JSON.stringify(this.geometry)) return;
    this.geometry = next;
    const v = this.view;
    if (v && !v.isDestroyed) syncDomSelection(v);
    if (v && !v.isDestroyed) v.dispatch(v.state.tr.setMeta(sectionLayoutKey, true).setMeta(paginationKey, { remeasure: true }).setMeta("addToHistory", false));
  }

  request() {
    const v = this.view;
    if (v && !v.isDestroyed) syncDomSelection(v);
    if (v && !v.isDestroyed) v.dispatch(v.state.tr.setMeta(paginationKey, { remeasure: true }).setMeta("addToHistory", false));
  }

  subscribe(fn: (i: LayoutInfo) => void) {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  publish(info: LayoutInfo) {
    const sameNotes = JSON.stringify(info.footnotes) === JSON.stringify(this.info.footnotes);
    const samePages = info.bodyW === this.info.bodyW && pagesKey(info.pages) === pagesKey(this.info.pages);
    const sameTiles = info.tileOk === this.info.tileOk && info.tileCols === this.info.tileCols && info.tileTarget === this.info.tileTarget && (info.tileRows ?? []).join() === (this.info.tileRows ?? []).join();
    if (info.pageCount === this.info.pageCount && info.pageStarts.join() === this.info.pageStarts.join() && sameNotes && samePages && sameTiles) return;
    // Değişmeyen parçaların başvurusu korunur (abone bileşenler gereksiz yere yeniden çizilmesin).
    if (sameNotes) info.footnotes = this.info.footnotes;
    if (samePages) info.pages = this.info.pages;
    if (sameTiles) info.tileRows = this.info.tileRows;
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

  /** Sayfa kapsayıcısındaki y konumunun düştüğü sayfa. */
  pageAtY(y: number): number {
    const pages = this.info.pages;
    for (let i = pages.length - 1; i >= 0; i--) if (y >= pages[i].top - this.geometry.gap / 2) return i;
    return 0;
  }
}

/** Erken biten satırın boş sayfa yuvalarını dolduran parça (tek akış ölçümünde gizlenir). */
function tileExtra(tag: "div" | "span", b: PageBreakMark): HTMLElement[] {
  if (!b.tileFill) return [];
  const x = Object.assign(document.createElement(tag), { className: "bl-pg-fill bl-pg-extra" });
  x.style.height = `${b.tileFill}px`;
  return [x];
}

function makeWidget(b: PageBreakMark) {
  if (b.kind === "rowhead") return headerCloneWidget;
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
  if (b.kind === "line" || b.kind === "spanline") {
    return () => {
      const s = document.createElement("span");
      s.className = b.kind === "line" ? "bl-pg bl-pg-line" : "bl-pg bl-pg-line bl-pg-span";
      if (b.tileRow) {
        // Boşluk önce, satır başı ara parçası sonda (sütun kümesi boşlukla birlikte biter).
        const fill = Object.assign(document.createElement("span"), { className: "bl-pg-fill" });
        fill.style.height = `${b.h}px`;
        s.append(fill, ...tileExtra("span", b), Object.assign(document.createElement("span"), { className: "bl-pg-tilerow" }));
      } else s.style.height = `${b.h}px`;
      return s;
    };
  }
  return () => {
    const d = document.createElement("div");
    d.className = b.kind === "span" ? "bl-pg bl-pg-span" : "bl-pg";
    if (b.tileRow) {
      const fill = Object.assign(document.createElement("div"), { className: "bl-pg-fill" });
      fill.style.height = `${b.h}px`;
      d.append(fill, ...tileExtra("div", b), Object.assign(document.createElement("div"), { className: "bl-pg-tilerow" }));
    } else d.style.height = `${b.h}px`;
    return d;
  };
}

function buildDeco(doc: PMNode, breaks: PageBreakMark[], styles: SectionStyle[]): DecorationSet {
  const decos: Decoration[] = [];
  for (const b of breaks) {
    if (b.pos <= 0 || b.pos > doc.content.size) continue;
    const head = b.kind === "rowhead" ? `:${doc.resolve(b.pos).parent.firstChild?.textContent ?? ""}:${JSON.stringify(doc.resolve(b.pos).parent.attrs)}` : "";
    decos.push(Decoration.widget(b.pos, makeWidget(b), { side: -1, key: `pg:${b.kind}:${b.pos}:${b.h.toFixed(2)}${b.tileRow ? `:t${(b.tileFill ?? 0).toFixed(1)}` : ""}${head}`, ignoreSelection: true }));
  }
  for (const s of styles) {
    const node = doc.nodeAt(s.pos);
    if (!node || node.type.name !== "section") continue;
    decos.push(Decoration.node(s.pos, s.pos + node.nodeSize, { style: `height:${s.height.toFixed(2)}px;column-fill:auto` }));
  }
  return DecorationSet.create(doc, decos);
}

function sameBreaks(a: PageBreakMark[], b: PageBreakMark[]) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const tol = a[i].kind === "span" || a[i].kind === "spanline" ? 0.05 : 0.5;
    if (a[i].pos !== b[i].pos || a[i].kind !== b[i].kind || Math.abs(a[i].h - b[i].h) > tol || !!a[i].tileRow !== !!b[i].tileRow || Math.abs((a[i].tileFill ?? 0) - (b[i].tileFill ?? 0)) > 0.5) return false;
  }
  return true;
}

function sameStyles(a: SectionStyle[], b: SectionStyle[]) {
  return a.length === b.length && a.every((s, i) => s.pos === b[i].pos && Math.abs(s.height - b[i].height) < 0.5);
}

class PaginationView {
  private raf = 0;
  private timer = 0;
  private ro: ResizeObserver;
  private lastWidth = 0;
  private slow = false;
  /** Çok sütunlu sayfa bölmelerinin ölçüm sonrası düzeltmeleri (ara parça konumu → px). */
  private spanFix = new Map<number, number>();
  private tailFix = new Map<number, number>();
  private refineRounds = 0;
  private spanTargets: { pos: number; target: number; page: number; sec: number }[] = [];
  private tailPredict: { pos: number; bottom: number; fill: boolean; sec: number }[] = [];

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
  private noteMeasurer(width: number) {
    if (Math.abs(width - this.noteWidth) > 0.5) {
      this.noteCache.clear();
      this.noteWidth = width;
    }
    return (key: string) => {
      const hit = this.noteCache.get(key);
      if (hit !== undefined) return hit;
      if (!this.measurer) {
        this.measurer = document.createElement("div");
        this.measurer.className = "bl-footnotes bl-doc is-measure";
        this.measurer.setAttribute("aria-hidden", "true");
        (this.view.dom.parentElement ?? document.body).appendChild(this.measurer);
      }
      this.measurer.style.width = `${width}px`;
      this.measurer.innerHTML = "";
      const row = document.createElement("div");
      row.className = "bl-fn-row";
      const sup = document.createElement("sup");
      sup.textContent = "99";
      const span = document.createElement("div");
      span.className = "bl-fn-text";
      span.innerHTML = noteHtml(key) || "​";
      row.append(sup, span);
      this.measurer.appendChild(row);
      const h = row.getBoundingClientRect().height / (this.ctrl.geometry.scale || 1);
      this.noteCache.set(key, h);
      return h;
    };
  }

  update(view: EditorView, prev: EditorState) {
    this.view = view;
    if (view.state.doc !== prev.doc) this.refineRounds = 0;
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
      if (old.breaks.length || old.styles.length) this.apply([], []);
      const geoms = sectionGeoms(view.state.doc, g.defaultPage);
      this.ctrl.publish({ ...singlePageInfo(geoms[0].page), bodyW: bodyWidth(geoms) });
      return;
    }
    const t0 = performance.now();
    const dom = view.dom as HTMLElement;
    const body = dom.closest(".bl-body");
    dom.classList.add("bl-measuring");
    body?.classList.add("bl-measuring-body");
    let breaks: PageBreakMark[] = [];
    let styles: SectionStyle[] = [];
    let info: LayoutInfo | null = null;
    let slot = 0;
    this.spanTargets = [];
    this.tailPredict = [];
    try {
      const geoms = sectionGeoms(view.state.doc, g.defaultPage);
      const bodyW = bodyWidth(geoms);
      const host = (dom.closest(".bl-pages") as HTMLElement | null) ?? dom;
      const r = host.getBoundingClientRect();
      const scale = g.scale || 1;
      const originY = r.top;
      const Y = (sy: number) => (sy - originY) / scale;
      const screenY = (y: number) => originY + y * scale;
      const units: Unit[] = collectUnits(view, geoms).map((u) => {
        const rr = u.el.getBoundingClientRect();
        return { ...u, top: Y(rr.top), bottom: Y(rr.bottom) };
      });
      const noteH = this.noteMeasurer(geoms[0].contentW);
      const slotH = g.tiles > 0 ? Math.max(...geoms.map((x) => x.h)) : 0;
      slot = slotH;
      const res = layoutFlow({ view, units, geoms, bodyW, gap: g.gap, Y, screenY, scale, noteH, slotH, columns: (u, c) => this.columns(u, c, geoms, Y, screenY) });
      breaks = res.breaks;
      styles = res.styles;
      info = res.info;
    } finally {
      dom.classList.remove("bl-measuring");
      body?.classList.remove("bl-measuring-body");
    }
    if (info) this.markTileRows(breaks, info, g.tiles, slot + g.gap);
    if (info) this.ctrl.lastPaged = { starts: info.pageStarts.slice(), count: info.pageCount };
    this.slow = performance.now() - t0 > 12;
    if (!sameBreaks(breaks, old.breaks) || !sameStyles(styles, old.styles)) this.apply(breaks, styles);
    if (info) this.ctrl.publish(info);
    if ((this.spanTargets.length || this.tailPredict.some((t) => !t.fill)) && this.refineRounds < 3) requestAnimationFrame(() => this.refine());
  }

  /**
   * Birden Çok Sayfa: sayfalar satırlara dağıtılır; her satırın ilk sayfasını başlatan kesme satır başı
   * olarak işaretlenir. Satır başı olamayan sınırlarda (tablo satırı, çok sütunlu bölüm içi, boş sayfa)
   * satır erken biter ya da uzar; kısa satırların boş yuvaları dolguyla tamamlanır.
   */
  private markTileRows(breaks: PageBreakMark[], info: LayoutInfo, tiles: number, stride: number) {
    info.tileOk = true;
    if (tiles < 2) return;
    const breakAt = (k: number) => {
      if (info.pages[k]?.blank || info.pages[k - 1]?.blank) return null;
      const pos = info.pageStarts[k];
      return breaks.find((x) => x.pos === pos && (x.kind === "block" || x.kind === "line")) ?? null;
    };
    const starts = planTileRows(info.pageCount, tiles, (k) => !!breakAt(k));
    const cols = tileColumns(starts, info.pageCount);
    info.tileRows = starts;
    info.tileCols = cols;
    info.tileTarget = tiles;
    for (let r = 1; r < starts.length; r++) {
      const b = breakAt(starts[r]);
      if (!b) continue;
      b.tileRow = true;
      const len = starts[r] - starts[r - 1];
      if (len < cols) b.tileFill = (cols - len) * stride;
    }
  }

  /**
   * Çok sütunlu bölüm: tek sayfaya sığıyorsa CSS sütunları dengeler (sayfa sonuyla bitiyorsa
   * sıralı doldurur); sığmıyorsa dar (tek sütun) ölçümle sayfa sınırları bulunur ve bu
   * noktalara tüm sütunları kesen ara parçalar konur.
   */
  private columns(u: Unit, c: ColumnCtx, geoms: SecGeom[], Y: (n: number) => number, screenY: (n: number) => number): ColumnResult {
    const g = geoms[u.sec];
    const shift = c.top - u.top;
    const realBottom = u.bottom + shift;
    let p = c.page;
    const breaks: PageBreakMark[] = [];
    const starts: { page: number; pos: number }[] = [];
    if (realBottom <= c.contentEnd(p) + 0.5) {
      const end = c.contentEnd(p);
      if (c.endsPage) {
        this.tailPredict.push({ pos: u.pos, bottom: end, fill: true, sec: u.pos });
        return { breaks, bottom: end, page: p, style: { pos: u.pos, height: Math.max(0, end - c.top) }, starts };
      }
      return { breaks, bottom: realBottom, page: p, starts };
    }
    // Dar ölçüm: bölüm geçici olarak tek ve sütun genişliğinde bir akışa döner.
    const el = u.el;
    el.classList.add("bl-narrow");
    try {
      const kids = collectSectionChildren(this.view, u.node, u.pos).map((k) => {
        const rr = k.el.getBoundingClientRect();
        const cs = getComputedStyle(k.el);
        return { ...k, top: Y(rr.top), bottom: Y(rr.bottom), mt: parseFloat(cs.marginTop) || 0, mb: parseFloat(cs.marginBottom) || 0, sec: u.sec };
      });
      const n0 = Y(el.getBoundingClientRect().top);
      const cols = g.cols;
      const slack = 4;
      let setTop = c.top;
      let A = 0;
      let cap = (c.contentEnd(p) - setTop) * cols - slack;
      let forceNext = false;
      for (const k of kids) {
        const kt = k.top - n0 - k.mt;
        const kb = k.bottom - n0 + k.mb;
        const forced = forceNext || k.breakBefore;
        forceNext = k.pageBreak;
        let first = true;
        for (let guard = 0; guard < 200 && ((forced && first && kt - A > 0.5) || kb - A > cap + 0.5); guard++) {
          let at = -1;
          let atY = 0;
          let kind: PageBreakMark["kind"] = "span";
          if (!(forced && first) && k.node.isTextblock && !k.node.attrs.keepLines && k.node.content.size > 0) {
            const from = breaks.length && breaks[breaks.length - 1].pos > k.pos ? breaks[breaks.length - 1].pos : k.pos + 1;
            const split = lineSplitAt(this.view, k, from, screenY(n0 + A + cap));
            if (split && split.pos > from) {
              const lineTop = Y(split.top) - n0;
              if (lineTop - A > 1) {
                at = split.pos;
                atY = lineTop;
                kind = "spanline";
              }
            }
          }
          if (at < 0) {
            if (kt - A <= 0.5 || (breaks.length && breaks[breaks.length - 1].pos >= k.pos)) break; // Tek başına sığmayan blok olduğu gibi kalır.
            at = k.pos;
            atY = kt;
          }
          const used = (atY - A) / cols;
          const np = c.nextPage();
          const target = c.pages[np].top + c.pages[np].mt;
          const est = target - (setTop + used) + (this.spanFix.get(at) ?? 0);
          breaks.push({ pos: at, kind, h: Math.max(1, est) });
          this.spanTargets.push({ pos: at, target, page: np, sec: u.pos });
          starts.push({ page: np, pos: at });
          p = np;
          setTop = target;
          A = atY;
          cap = (c.contentEnd(p) - setTop) * cols - slack;
          first = false;
        }
      }
      const last = kids[kids.length - 1];
      const R = last ? last.bottom - n0 + last.mb - A : 0;
      if (c.endsPage) {
        const end = c.contentEnd(p);
        this.tailPredict.push({ pos: u.pos, bottom: end, fill: true, sec: u.pos });
        return { breaks, bottom: end, page: p, style: { pos: u.pos, height: Math.max(0, end - c.top) }, starts };
      }
      const bottom = setTop + R / cols + (this.tailFix.get(u.pos) ?? 0);
      this.tailPredict.push({ pos: u.pos, bottom, fill: false, sec: u.pos });
      return { breaks, bottom, page: p, starts };
    } finally {
      el.classList.remove("bl-narrow");
    }
  }

  /** Ara parçaların gerçek konumlarına göre düzeltme; gerekirse yeniden ölçer. */
  private refine() {
    const view = this.view;
    if (view.isDestroyed || !this.ctrl.geometry.enabled) return;
    const dom = view.dom as HTMLElement;
    // Yan yana dizimde konumlar satır/sütuna göredir: düzeltme tek akışta (sütunsuz) ölçülür.
    const body = dom.closest(".bl-body");
    const tiled = !!body?.classList.contains("is-tiled");
    if (tiled) body!.classList.add("bl-measuring-body");
    let changed = false;
    try {
      changed = this.refineErrors(view, dom);
    } finally {
      if (tiled) body!.classList.remove("bl-measuring-body");
    }
    if (changed) {
      this.refineRounds++;
      this.measure();
    }
  }

  private refineErrors(view: EditorView, dom: HTMLElement): boolean {
    const host = (dom.closest(".bl-pages") as HTMLElement | null) ?? dom;
    const scale = this.ctrl.geometry.scale || 1;
    const originY = host.getBoundingClientRect().top;
    const Y = (sy: number) => (sy - originY) / scale;
    let changed = false;
    // Zincirdeki her ara parçanın yalnız kendi hatası düzeltilir (öncekinin düzeltmesi sonrakini zaten kaydırır).
    const upstream = new Map<number, number>();
    for (const s of this.spanTargets) {
      const el = findWidget(view, s.pos);
      if (!el) continue;
      const err = s.target - Y(el.getBoundingClientRect().bottom);
      const local = err - (upstream.get(s.sec) ?? 0);
      upstream.set(s.sec, err);
      if (Math.abs(err) > 0.25 && Math.abs(local) > 0.05) {
        this.spanFix.set(s.pos, (this.spanFix.get(s.pos) ?? 0) + local);
        changed = true;
      }
    }
    for (const t of this.tailPredict) {
      if (t.fill) continue;
      const el = view.nodeDOM(t.pos);
      if (!(el instanceof HTMLElement)) continue;
      const err = Y(el.getBoundingClientRect().bottom) - t.bottom;
      const local = err + (upstream.get(t.sec) ?? 0);
      if (Math.abs(local) > 0.25) {
        this.tailFix.set(t.pos, (this.tailFix.get(t.pos) ?? 0) + local);
        changed = true;
      }
    }
    return changed;
  }

  private apply(breaks: PageBreakMark[], styles: SectionStyle[]) {
    syncDomSelection(this.view);
    const tr = this.view.state.tr.setMeta(paginationKey, { breaks, styles }).setMeta("addToHistory", false);
    this.view.dispatch(tr);
  }
}

/**
 * Zamanlayıcıyla gelen (kullanıcı kaynaklı olmayan) bir işlemden önce bekleyen yerel seçim değişikliği
 * okunur: tarayıcı seçimi değiştirmiş ama "selectionchange" henüz işlenmemişse, güncelleme seçimi eski
 * konuma geri yazmasın (ör. Home'dan hemen sonra Ctrl+Shift+Sağ). ProseMirror'ın iç gözlemcisi kullanılır.
 */
export function syncDomSelection(view: EditorView) {
  try {
    // forceFlush: "selectionchange" gelmiş, gecikmeli okuma bekliyor; flush: olay henüz gelmemiş.
    const obs = (view as unknown as { domObserver?: { flush(): void; forceFlush(): void } }).domObserver;
    obs?.forceFlush();
    obs?.flush();
  } catch {
    /* gözlemci yoksa yoksay */
  }
}

/** Konumdaki, tüm sütunları kesen sayfa ara parçası. */
function findWidget(view: EditorView, pos: number): HTMLElement | null {
  for (const el of Array.from(view.dom.querySelectorAll<HTMLElement>(".bl-pg-span"))) {
    try {
      if (view.posAtDOM(el, 0) === pos) return el;
    } catch {
      // yok say
    }
  }
  return null;
}

/** Dipnot anahtarından ölçüm HTML'i (zengin içerik JSON'u ya da düz metin). */
function noteHtml(key: string): string {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  if (key.startsWith("[") || key.startsWith("{")) {
    try {
      const json = JSON.parse(key) as { content?: unknown[] } | unknown[];
      const blocks = Array.isArray(json) ? json : (json.content ?? []);
      const text = (n: unknown): string => {
        const o = n as { text?: string; content?: unknown[]; type?: string };
        if (o.type === "text") return esc(o.text ?? "");
        if (o.type === "hardBreak") return "<br>";
        return (o.content ?? []).map(text).join("");
      };
      return blocks.map((b) => `<p style="margin:0">${text(b) || "​"}</p>`).join("");
    } catch {
      return esc(key);
    }
  }
  return esc(key);
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
          init: () => ({ breaks: [], styles: [], deco: DecorationSet.empty, remeasure: 0 }),
          apply(tr, value, _old, newState) {
            const meta = tr.getMeta(paginationKey) as { breaks?: PageBreakMark[]; styles?: SectionStyle[]; remeasure?: boolean } | undefined;
            if (meta?.breaks) {
              const styles = meta.styles ?? [];
              return { breaks: meta.breaks, styles, deco: buildDeco(newState.doc, meta.breaks, styles), remeasure: value.remeasure };
            }
            let next = value;
            // Sayfasız görünümde son sayfalamanın sayfa başları düzenlemeyle kayar.
            if (tr.docChanged && !ctrl.geometry.enabled && ctrl.lastPaged) ctrl.lastPaged = { ...ctrl.lastPaged, starts: ctrl.lastPaged.starts.map((p) => tr.mapping.map(p, -1)) };
            if (tr.docChanged) {
              const breaks = value.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1) }));
              const styles = value.styles.map((s) => ({ ...s, pos: tr.mapping.map(s.pos, 1) }));
              let deco = value.deco.map(tr.mapping, tr.doc);
              // Kapsayıcı düğümün özniteliği değişince (ör. bölüm) sınırdaki parçacıklar düşebilir: yeniden kurulur.
              if (deco.find().length !== value.deco.find().length) deco = buildDeco(tr.doc, breaks, styles);
              next = { ...value, breaks, styles, deco };
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
