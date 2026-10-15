// Yazdırma Düzeni sayfalaması: blokların doğal konumları ölçülür, sayfa sınırını aşan
// bloğun (ya da satırın / tablo satırının) önüne sayfa boşluğu kadar yükseklikte bir
// dekorasyon parçacığı konur. İçerik tek bir akışta kalır; sayfa zeminleri arkada çizilir.
// Her sayfa kendi bölümünün boyut ve kenar boşluklarını kullanır.
import { headerCloneWidget, repeatHeaderHeight } from "./tableTools";
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, EditorView } from "@tiptap/pm/view";
import { defaultPage, type PageSetup } from "../model/document";
import { bodyWidth, geomOf, sectionGeoms, sectionLayoutKey, type SecGeom } from "./sections";
import { planTileRows, tileColumns } from "./tileRows";
import { alignUnits, flowEnv, footnotesIn, layoutFlow, lineSplitAt, type ColumnCtx, type ColumnResult, type FlowCache, type FlowMeasure, type FlowResult, type FlowUnit, type LayoutInfo, type PageBreakMark, type PageInfo, type Rect, type SectionStyle } from "./paginate";

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

/** Ondalık bir basamağa yuvarlanınca aynı mı (her tuşta dize kurmadan). */
const near1 = (a: number, b: number) => a === b || a.toFixed(1) === b.toFixed(1);

function samePageList(a: PageInfo[], b: PageInfo[]) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const p = a[i];
    const q = b[i];
    if (!near1(p.top, q.top) || !near1(p.w, q.w) || !near1(p.h, q.h) || !near1(p.mt, q.mt) || !near1(p.mb, q.mb) || !near1(p.ml, q.ml) || !near1(p.mr, q.mr) || !near1(p.left, q.left)) return false;
    if (p.headerY !== q.headerY || p.footerY !== q.footerY || p.section !== q.section || !!p.first !== !!q.first || !!p.blank !== !!q.blank || !!p.titlePage !== !!q.titlePage) return false;
  }
  return true;
}

const sameList = <T,>(a: readonly T[], b: readonly T[]) => a.length === b.length && a.every((x, i) => x === b[i]);

/** React tarafıyla paylaşılan denetleyici: geometri ve düzen bilgisi. */
export class PaginationController {
  geometry: PageGeometry = { enabled: false, gap: 16, defaultPage: defaultPage(), tiles: 0, scale: 1 };
  info: LayoutInfo = singlePageInfo(defaultPage());
  view: EditorView | null = null;
  private listeners = new Set<(i: LayoutInfo) => void>();
  measureNow: (() => void) | null = null;
  /** Son Yazdırma Düzeni sayfalaması (Taslak/Anahat görünümünde durum çubuğu için; düzenlemelerle eşlenir). */
  lastPaged: { starts: number[]; count: number } | null = null;
  /** Son ölçüm: artımlı mı, süresi ve yeniden hesaplanan birim aralığı (tanılama ve testler için). */
  lastRun: { incremental: boolean; ms: number; units: number; from: number; to: number } | null = null;

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
    const sameNotes = info.footnotes.length === this.info.footnotes.length && info.footnotes.every((f, i) => sameList(f, this.info.footnotes[i]));
    const samePages = info.bodyW === this.info.bodyW && samePageList(info.pages, this.info.pages);
    const sameTiles = info.tileOk === this.info.tileOk && info.tileCols === this.info.tileCols && info.tileTarget === this.info.tileTarget && sameList(info.tileRows ?? [], this.info.tileRows ?? []);
    if (info.pageCount === this.info.pageCount && sameList(info.pageStarts, this.info.pageStarts) && sameNotes && samePages && sameTiles) return;
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

/* ------------------------------------------------------------------ */
/* Ölçüm birimleri                                                      */
/* ------------------------------------------------------------------ */

const LIST_TYPES = new Set(["bulletList", "orderedList"]);

/** Belge ağacındaki ölçüm birimi adayı (DOM'a bakmadan). */
interface UnitSpec {
  node: PMNode;
  pos: number;
  breakPos: number;
  kind: FlowUnit["kind"];
  sec: number;
  ctx: unknown[];
  /** Tablo satırı: tablonun konumu ve satır sırası. */
  tablePos?: number;
  rowIdx?: number;
}

/** Birimler: tek sütunlu bölümlerde bloklar, liste paragrafları ve tablo satırları; çok sütunlu bölüm tek birimdir. */
function walkUnits(doc: PMNode, geoms: SecGeom[], out: (s: UnitSpec) => void) {
  const visit = (node: PMNode, pos: number, breakPos: number, sec: number, ctx: unknown[]) => {
    if (LIST_TYPES.has(node.type.name)) {
      const lctx = [...ctx, node.attrs];
      node.forEach((li, off, idx) => {
        const liPos = pos + 1 + off;
        const ictx = [...lctx, li.attrs];
        li.forEach((c, off2, idx2) => {
          const cPos = liPos + 1 + off2;
          const bp = idx2 === 0 ? (idx === 0 ? breakPos : liPos) : cPos;
          visit(c, cPos, bp, sec, ictx);
        });
      });
      return;
    }
    if (node.type.name === "table") {
      const tctx = [...ctx, node.attrs];
      node.forEach((row, off, idx) => {
        const rPos = pos + 1 + off;
        out({ node: row, pos: rPos, breakPos: idx === 0 ? breakPos : rPos, kind: idx === 0 ? "block" : "row", sec, ctx: tctx, tablePos: pos, rowIdx: idx });
      });
      return;
    }
    out({ node, pos, breakPos, kind: "block", sec, ctx });
  };
  doc.forEach((section, secOff, si) => {
    const g = geoms[si];
    if (g && g.cols > 1) {
      // Bölümden önceki boşluk bölümler arasına (sütunların dışına) konur.
      out({ node: section, pos: secOff, breakPos: secOff, kind: "section", sec: si, ctx: [] });
      return;
    }
    section.forEach((child, off) => visit(child, secOff + 1 + off, secOff + 1 + off, si, []));
  });
}

/** Düğümden düzen birimi (konumlar ölçülmeden). */
function unitOf(s: UnitSpec, secStart: boolean): FlowUnit {
  const n = s.node;
  return {
    key: n,
    ctx: s.ctx,
    pos: s.pos,
    breakPos: s.breakPos,
    end: s.pos + n.nodeSize,
    kind: s.kind,
    top: 0,
    bottom: 0,
    keepNext: s.kind !== "section" && (!!n.attrs.keepNext || n.type.name === "heading"),
    pageBreak: n.type.name === "pageBreak",
    colBreak: n.type.name === "columnBreak",
    breakBefore: s.kind !== "section" && !!n.attrs.pageBreakBefore,
    keepLines: !!n.attrs.keepLines,
    text: n.isTextblock && n.content.size > 0,
    notes: footnotesIn(n),
    sec: s.sec,
    secStart,
    hh: 0,
  };
}

const nodeEl = (view: EditorView, pos: number) => {
  const el = view.nodeDOM(pos);
  return el instanceof HTMLElement ? el : null;
};

/** Tablo satırı: yeni sayfada yinelenen başlık satırının yüksekliği (tablo başına bir kez ölçülür). */
function headerHeight(view: EditorView, s: UnitSpec, scale: number, heads: Map<number, number>): number {
  if (s.kind !== "row" || s.tablePos === undefined) return 0;
  let hh = heads.get(s.tablePos);
  if (hh === undefined) heads.set(s.tablePos, (hh = repeatHeaderHeight(view, s.pos, scale)));
  return hh;
}

type DomUnit = FlowUnit & { el: HTMLElement; node: PMNode };

/** Tam ölçümün birimleri (öğesi olmayan düğümler atlanır). */
function collectUnits(view: EditorView, geoms: SecGeom[], scale: number): DomUnit[] {
  const out: DomUnit[] = [];
  let sec = -1;
  const heads = new Map<number, number>();
  walkUnits(view.state.doc, geoms, (s) => {
    const el = nodeEl(view, s.pos);
    if (!el) return;
    const u = unitOf(s, s.sec !== sec) as DomUnit;
    sec = s.sec;
    u.el = el;
    u.node = s.node;
    u.hh = headerHeight(view, s, scale, heads);
    out.push(u);
  });
  return out;
}

/** Çok sütunlu bölümün dar (tek sütun) ölçümündeki birimleri. */
function collectSectionChildren(view: EditorView, section: PMNode, secOff: number) {
  const out: { pos: number; node: PMNode; el: HTMLElement; breakBefore: boolean; pageBreak: boolean }[] = [];
  section.forEach((child, off) => {
    const pos = secOff + 1 + off;
    const el = nodeEl(view, pos);
    if (!el) return;
    out.push({ pos, node: child, el, breakBefore: !!child.attrs.pageBreakBefore, pageBreak: child.type.name === "pageBreak" });
  });
  return out;
}

const marginTopOf = (el: HTMLElement) => parseFloat(getComputedStyle(el).marginTop) || 0;

const viewCoords = (view: EditorView, p: number, side: number): Rect | null => {
  try {
    return view.coordsAtPos(p, side);
  } catch {
    return null;
  }
};

interface BlockDesc {
  node?: unknown;
  contentDOM?: Node | null;
  domFromPos?: (pos: number, side: number) => unknown;
}

/**
 * Metin bloğu içindeki konumun ekran dikdörtgeni. `coordsAtPos` konumu kökten başlayıp kardeşleri tek tek
 * sayarak bulur (büyük belgede bölüm başına binlerce blok); blok içindeki konumlarda kök aramayı doğrudan
 * bloğun görünüm tanımına yönlendirmek aynı sonucu verir. Tanım beklenen biçimde değilse (ya da ölçümün
 * ilk karşılaştırmasında sonuç farklıysa) olağan `coordsAtPos` kullanılır.
 */
function blockCoords(view: EditorView, el: HTMLElement, key: object, start: number, check: { ok?: boolean }): (p: number, side: number) => Rect | null {
  const plain = (p: number, side: number) => viewCoords(view, p, side);
  const desc = (el as unknown as { pmViewDesc?: BlockDesc }).pmViewDesc;
  const find = desc?.domFromPos;
  if (check.ok === false || !desc || desc.node !== key || !desc.contentDOM || typeof find !== "function") return plain;
  const shim = { state: view.state, domObserver: { flush() {} }, docView: { domFromPos: (p: number, side: number) => find.call(desc, p - start, side) } } as unknown as EditorView;
  const local = (p: number, side: number) => {
    try {
      return EditorView.prototype.coordsAtPos.call(shim, p, side);
    } catch {
      return null;
    }
  };
  if (check.ok === undefined) {
    const a = local(start, 1);
    const b = plain(start, 1);
    check.ok = !!a && !!b && a.top === b.top && a.bottom === b.bottom;
    if (!check.ok) return plain;
  }
  return local;
}

/** Artımlı ölçüm uygulanamadı: tam ölçüme geçilir. */
class NeedFull extends Error {}

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
  /** Önceki sayfalamanın akış durumu (artımlı devam için). */
  private cache: FlowCache | null = null;
  /** Sonraki ölçüm tam olmalı (geometri, yazı tipi, resim, boyut değişti ya da istendi). */
  private needFull = true;
  /** Belge dışı DOM değişiklikleri: içeriği değişen birimler ve birime bağlanamayan değişiklik. */
  private mo: MutationObserver;
  private mutated = new Set<object>();
  private mutatedAll = false;
  /** Birime henüz bağlanamayan değişiklikler (ölçümde konumdan bulunur; `attr`: kapsayıcının özniteliği). */
  private pending: { el: Element; attr: boolean }[] = [];
  /** Birim öğesi → düğüm (DOM değişikliğini birime bağlamak için). */
  private unitEls = new WeakMap<Element, object>();
  /** Artımlı ölçüm için geçici gizlenen ara parçalar. */
  private hidden: [HTMLElement, string][] = [];

  constructor(
    private view: EditorView,
    private ctrl: PaginationController,
  ) {
    ctrl.view = view;
    ctrl.measureNow = () => {
      this.needFull = true;
      this.measure();
    };
    this.ro = new ResizeObserver(() => {
      const w = view.dom.clientWidth;
      if (Math.abs(w - this.lastWidth) > 0.5) {
        this.lastWidth = w;
        this.needFull = true;
        this.schedule();
      }
    });
    this.ro.observe(view.dom);
    this.mo = new MutationObserver((recs) => this.noteMutations(recs));
    this.mo.observe(view.dom, { subtree: true, childList: true, characterData: true, attributes: true, attributeOldValue: true });
    view.dom.addEventListener("load", this.onLoad, true);
    document.fonts?.ready.then(() => this.onLoad()).catch(() => {});
    document.fonts?.addEventListener?.("loadingdone", this.onLoad);
    this.schedule();
  }

  private onLoad = () => {
    this.needFull = true;
    this.schedule();
  };
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

  /**
   * Belge değişikliği dışındaki DOM değişiklikleri de (dipnot numarası, alan değeri, başka eklentilerin
   * dekorasyonları) izlenir: içinde oldukları birim değişmiş sayılır; birime bağlanamayan değişiklik tam
   * ölçüm ister.
   */
  private noteMutations(recs: MutationRecord[]) {
    const root = this.view.dom;
    for (const r of recs) {
      if (this.mutatedAll) return;
      const t = r.target instanceof Element ? r.target : r.target.parentElement;
      if (!t || !root.contains(t) || t.closest(".bl-pg")) continue;
      // Kök öğenin öznitelikleri (odak, imleç sınıfları) yerleşimi değiştirmez.
      if (t === root && r.type === "attributes") continue;
      // Aynı değerin yeniden yazılması (ör. tablo görünümü her güncellemede stilini yeniden kurar).
      if (r.type === "attributes" && r.attributeName && r.oldValue === t.getAttribute(r.attributeName)) continue;
      if (r.type === "childList") {
        const added = Array.from(r.addedNodes);
        // Yalnız sayfa ara parçaları eklenip çıkarıldı.
        if ([...added, ...Array.from(r.removedNodes)].every((x) => x instanceof Element && x.classList.contains("bl-pg"))) continue;
        for (const x of added) {
          if (!(x instanceof Element)) continue;
          // Başka bir eklentinin bloklar arasına koyduğu araç: birime bağlanamaz.
          if (x.classList.contains("ProseMirror-widget") && !x.classList.contains("bl-pg") && !this.unitOfEl(t)) {
            this.mutatedAll = true;
            return;
          }
          if (!x.classList.contains("bl-pg")) this.pending.push({ el: x, attr: false });
        }
      }
      const key = this.unitOfEl(t);
      if (key) this.mutated.add(key);
      else this.pending.push({ el: t, attr: r.type === "attributes" });
    }
  }

  private unitOfEl(t: Element): object | undefined {
    const root = this.view.dom;
    for (let e: Element | null = t; e && e !== root; e = e.parentElement) {
      const key = this.unitEls.get(e);
      if (key) return key;
    }
    return undefined;
  }

  update(view: EditorView, prev: EditorState) {
    this.view = view;
    if (view.state.doc !== prev.doc) this.refineRounds = 0;
    const remeasure = paginationKey.getState(view.state)?.remeasure !== paginationKey.getState(prev)?.remeasure;
    if (remeasure) this.needFull = true;
    // Yalnız belge değiştiğinde ya da yeniden ölçüm istendiğinde ölç (kendi ara parça güncellememiz hariç).
    if (view.state.doc !== prev.doc || remeasure) this.schedule();
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
    this.mo.disconnect();
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
      this.cache = null;
      this.needFull = true;
      return;
    }
    const t0 = performance.now();
    this.spanTargets = [];
    this.tailPredict = [];
    const geoms = sectionGeoms(view.state.doc, g.defaultPage);
    const bodyW = bodyWidth(geoms);
    const scale = g.scale || 1;
    const slotH = g.tiles > 0 ? Math.max(...geoms.map((x) => x.h)) : 0;
    const host = (view.dom.closest(".bl-pages") as HTMLElement | null) ?? (view.dom as HTMLElement);
    let res: FlowResult | null = null;
    // Henüz bildirilmemiş DOM değişiklikleri.
    this.noteMutations(this.mo.takeRecords());
    try {
      const cache = this.cache;
      // Belge değişikliğinde (yan yana dizim ve çok sütunlu bölüm yoksa) değişen yerden artımlı ölçülür;
      // olmazsa (beklenmeyen durumda da) tam ölçüme geçilir.
      if (cache && !this.needFull && !this.mutatedAll && g.tiles === 0 && geoms.every((x) => x.cols === 1) && cache.env === flowEnv(geoms, bodyW, g.gap, slotH, scale)) {
        try {
          res = this.measureIncremental(cache, geoms, bodyW, scale, host);
        } catch {
          this.restoreHidden();
          res = null;
        }
      }
      if (!res) res = this.measureFull(geoms, bodyW, scale, slotH, host);
      const { breaks, styles, info } = res;
      this.markTileRows(breaks, info, g.tiles, slotH + g.gap);
      this.ctrl.lastPaged = { starts: info.pageStarts.slice(), count: info.pageCount };
      const ms = performance.now() - t0;
      this.slow = ms > 12;
      this.ctrl.lastRun = { incremental: res.stats.incremental, ms, units: res.cache.units.length, from: res.stats.from, to: res.stats.to };
      if (!sameBreaks(breaks, old.breaks) || !sameStyles(styles, old.styles)) this.apply(breaks, styles);
      this.ctrl.publish(info);
      this.cache = res.cache;
      this.needFull = false;
    } finally {
      this.restoreHidden();
      // Kendi değişikliklerimiz (ölçüm sınıfları, ara parçalar) birim değişikliği sayılmaz.
      this.mo.takeRecords();
      this.mutated.clear();
      this.mutatedAll = false;
      this.pending = [];
    }
    if ((this.spanTargets.length || this.tailPredict.some((t) => !t.fill)) && this.refineRounds < 3) requestAnimationFrame(() => this.refine());
  }

  /** Tam ölçüm: tüm ara parçalar gizlenip bütün birimlerin doğal konumları ölçülür. */
  private measureFull(geoms: SecGeom[], bodyW: number, scale: number, slotH: number, host: HTMLElement): FlowResult {
    const view = this.view;
    const g = this.ctrl.geometry;
    const dom = view.dom as HTMLElement;
    const body = dom.closest(".bl-body");
    dom.classList.add("bl-measuring");
    body?.classList.add("bl-measuring-body");
    try {
      const originY = host.getBoundingClientRect().top;
      const Y = (sy: number) => (sy - originY) / scale;
      const screenY = (y: number) => originY + y * scale;
      const units = collectUnits(view, geoms, scale).map((u) => {
        const rr = u.el.getBoundingClientRect();
        u.top = Y(rr.top);
        u.bottom = Y(rr.bottom);
        return u;
      });
      for (const u of units) this.unitEls.set(u.el, u.key);
      const cf = new Map<number, (p: number, side: number) => Rect | null>();
      const check = {};
      const m: FlowMeasure = {
        outerTop: (i) => {
          const u = units[i];
          const el = view.nodeDOM(u.breakPos);
          const v = el instanceof HTMLElement && el !== u.el ? Y(el.getBoundingClientRect().top) - marginTopOf(el) : u.top - marginTopOf(u.el);
          u.ot = v - u.top;
          return v;
        },
        coords: (i, p, side) => {
          let f = cf.get(i);
          if (!f) cf.set(i, (f = blockCoords(view, units[i].el, units[i].key, units[i].pos + 1, check)));
          return f(p, side);
        },
        lineHeight: (i) => parseFloat(getComputedStyle(units[i].el).lineHeight),
        noteH: this.noteMeasurer(geoms[0].contentW),
        columns: (i, c) => this.columns(units[i], c, geoms, Y, screenY),
      };
      return layoutFlow({ units, geoms, bodyW, gap: g.gap, Y, screenY, scale, slotH, m });
    } finally {
      dom.classList.remove("bl-measuring");
      body?.classList.remove("bl-measuring-body");
    }
  }

  /**
   * Artımlı ölçüm: belge ağacı önceki birimlerle düğüm kimliğinden eşlenir. Değişmeyen birimlerin doğal
   * konumu önbellektekinin kaydırılmışıdır; kaydırma, değişen birimlerin ve sınırdaki birimlerin önceki
   * birime uzaklığından ölçülür (aradaki ara parçalar ölçüm süresince gizlenir, bütün belge gizlenmez).
   * Düzen birimi değişen yerin sayfasından başlar ve önceki hesapla örtüşünce durur.
   */
  private measureIncremental(cache: FlowCache, geoms: SecGeom[], bodyW: number, scale: number, host: HTMLElement): FlowResult {
    const view = this.view;
    const g = this.ctrl.geometry;
    const old = cache.units;
    // 1. Birimler ve eşleşme (DOM'a bakmadan)
    const specs: UnitSpec[] = [];
    const units: FlowUnit[] = [];
    let sec = -1;
    let last = -1;
    walkUnits(view.state.doc, geoms, (s) => {
      const j = cache.index.get(s.node);
      let u: FlowUnit;
      if (j !== undefined && j > last && !this.mutated.has(s.node)) {
        // Aynı düğüm: bayraklar ve dipnotlar aynı; konumlar sonra.
        const o = old[j];
        u = { ...o, ctx: s.ctx, pos: s.pos, breakPos: s.breakPos, end: s.pos + s.node.nodeSize, kind: s.kind, sec: s.sec, secStart: s.sec !== sec };
        last = j;
      } else {
        if (!nodeEl(view, s.pos)) return;
        u = unitOf(s, s.sec !== sec);
      }
      sec = s.sec;
      specs.push(s);
      units.push(u);
    });
    const n = units.length;
    const align = alignUnits(units, cache);
    const changed = new Uint8Array(n);
    for (let i = 0; i < n; i++) if (this.mutated.has(units[i].key)) changed[i] = 1;
    // Kaydı tutulmayan öğelerdeki değişiklikler: konumu içeren birim (bloklar arasıysa yapısal değişiklik).
    for (const { el, attr } of this.pending) {
      if (!el.isConnected) continue;
      let pos: number;
      try {
        pos = view.posAtDOM(el, 0);
      } catch {
        throw new NeedFull();
      }
      let lo = 0;
      let hi = n - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (units[mid].pos < pos) lo = mid;
        else hi = mid - 1;
      }
      if (n && units[lo].pos < pos && pos < units[lo].end) changed[lo] = 1;
      else if (attr) throw new NeedFull();
    }
    for (let i = 0; i < n; i++) if (changed[i]) align[i] = -1;
    // 2. Ölçülecek birimler: değişenler ve eşleşme zincirinin koptuğu yerler (ve öncekileri).
    const els: (HTMLElement | null)[] = new Array(n).fill(null);
    const elOf = (i: number) => {
      const el = els[i] ?? (els[i] = nodeEl(view, units[i].pos));
      if (!el) throw new NeedFull();
      return el;
    };
    const boundary = new Uint8Array(n);
    for (let i = 0; i < n; i++) {
      const dirty = align[i] < 0;
      if (!(dirty || (i === 0 ? align[0] !== 0 : align[i] !== align[i - 1] + 1))) continue;
      boundary[i] = 1;
      // Önceki birimle arasındaki ara parçalar ve değişen birimin içindekiler gizlenir.
      const u = units[i];
      const bp = u.breakPos === u.pos ? elOf(i) : nodeEl(view, u.breakPos);
      for (let w = bp?.previousElementSibling; w instanceof HTMLElement && w.classList.contains("ProseMirror-widget"); w = w.previousElementSibling) if (w.classList.contains("bl-pg")) this.hide(w);
      if (dirty) for (const w of Array.from(elOf(i).getElementsByClassName("bl-pg"))) if (w instanceof HTMLElement) this.hide(w);
    }
    // 3. Doğal konumlar (ilk ölçüm düzeni zorlar; gizleme ondan önce yapıldı)
    const originY = host.getBoundingClientRect().top;
    const Y = (sy: number) => (sy - originY) / scale;
    const screenY = (y: number) => originY + y * scale;
    let delta = 0;
    for (let i = 0; i < n; i++) {
      const u = units[i];
      const j = align[i];
      if (!boundary[i]) {
        u.top = old[j].top + delta;
        u.bottom = old[j].bottom + delta;
        continue;
      }
      const r = elOf(i).getBoundingClientRect();
      const top = i === 0 ? Y(r.top) : units[i - 1].bottom + (Y(r.top) - Y(elOf(i - 1).getBoundingClientRect().bottom));
      if (j < 0) {
        u.top = top;
        u.bottom = top + (Y(r.bottom) - Y(r.top));
        u.ot = undefined;
        this.unitEls.set(elOf(i), u.key);
        continue;
      }
      // Değişmeyen birim: önceki doğal yüksekliği gerçek yüksekliğinden (iç ara parçalar düşülerek) doğrulanır.
      let inner = 0;
      for (const w of Array.from(elOf(i).getElementsByClassName("bl-pg"))) inner += Y(w.getBoundingClientRect().bottom) - Y(w.getBoundingClientRect().top);
      if (Math.abs(Y(r.bottom) - Y(r.top) - inner - (old[j].bottom - old[j].top)) > 0.05) throw new NeedFull();
      delta = top - old[j].top;
      u.top = top;
      u.bottom = old[j].bottom + delta;
    }
    // 4. Yinelenen başlık satırı yüksekliği: ilk satırı ya da kendisi değişen tablolarda yeniden ölçülür.
    const heads = new Map<number, number>();
    const headIdx = new Map<number, number>();
    for (let i = 0; i < n; i++) {
      const s = specs[i];
      if (s.tablePos === undefined) continue;
      if (s.rowIdx === 0) headIdx.set(s.tablePos, i);
      if (s.kind !== "row") continue;
      const head = headIdx.get(s.tablePos);
      if (align[i] >= 0 && head !== undefined && align[head] >= 0) continue;
      units[i].hh = headerHeight(view, s, scale, heads);
    }
    // 5. Düzen: ölçüm arayüzü değişmeyen birimlerde gerçek konumdan ara parça etkisini düşer.
    const lines = (paginationKey.getState(view.state)?.breaks ?? []).filter((b) => b.kind === "line");
    const shift = new Map<number, { base: number; inner: { pos: number; h: number }[]; at: (p: number, side: number) => Rect | null }>();
    const check = {};
    const shiftOf = (i: number) => {
      let sh = shift.get(i);
      if (!sh) {
        const u = units[i];
        const el = elOf(i);
        const ws = Array.from(el.getElementsByClassName("bl-pg"));
        const inside = lines.filter((b) => b.pos > u.pos && b.pos < u.end);
        if (ws.length !== inside.length) throw new NeedFull();
        sh = { base: el.getBoundingClientRect().top - screenY(u.top), inner: inside.map((b, k) => ({ pos: b.pos, h: ws[k].getBoundingClientRect().height })), at: blockCoords(view, el, u.key, u.pos + 1, check) };
        shift.set(i, sh);
      }
      return sh;
    };
    const m: FlowMeasure = {
      outerTop: (i) => {
        const u = units[i];
        if (u.ot === undefined) {
          const el = elOf(i);
          const bp = u.breakPos === u.pos ? el : nodeEl(view, u.breakPos);
          u.ot = bp && bp !== el ? Y(bp.getBoundingClientRect().top) - marginTopOf(bp) - Y(el.getBoundingClientRect().top) : -marginTopOf(el);
        }
        return u.top + u.ot;
      },
      coords: (i, p, side) => {
        const sh = shiftOf(i);
        const r = sh.at(p, side);
        if (!r) return null;
        let d = sh.base;
        for (const w of sh.inner) if (w.pos < p || (w.pos === p && side >= 0)) d += w.h;
        return { top: r.top - d, bottom: r.bottom - d };
      },
      lineHeight: (i) => parseFloat(getComputedStyle(elOf(i)).lineHeight),
      noteH: this.noteMeasurer(geoms[0].contentW),
      columns: () => {
        throw new NeedFull();
      },
    };
    return layoutFlow({ units, geoms, bodyW, gap: g.gap, Y, screenY, scale, slotH: 0, m, changed }, cache);
  }

  private hide(el: HTMLElement) {
    this.hidden.push([el, el.style.display]);
    el.style.display = "none";
  }

  private restoreHidden() {
    for (const [el, d] of this.hidden) el.style.display = d;
    this.hidden = [];
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
  private columns(u: DomUnit, c: ColumnCtx, geoms: SecGeom[], Y: (n: number) => number, screenY: (n: number) => number): ColumnResult {
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
        return { ...k, top: Y(rr.top), bottom: Y(rr.bottom), mt: parseFloat(cs.marginTop) || 0, mb: parseFloat(cs.marginBottom) || 0 };
      });
      const n0 = Y(el.getBoundingClientRect().top);
      const cols = g.cols;
      const slack = 4;
      let setTop = c.top;
      let A = 0;
      let cap = (c.contentEnd(p) - setTop) * cols - slack;
      let forceNext = false;
      const coords = (pos: number, side = 1) => viewCoords(this.view, pos, side);
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
            const split = lineSplitAt(coords, k.pos + 1, k.pos + k.node.nodeSize - 1, from, screenY(n0 + A + cap));
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
      this.needFull = true;
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
