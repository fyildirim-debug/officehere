// Sayfalama algoritması: ölçülen blok konumlarından sayfa listesi ve sayfa boşluğu parçacıkları.
// Sayfalar sıralı oluşturulur; her sayfa bölümünün boyut/kenar boşluğunu taşır. Çok sütunlu
// bölümler CSS sütunlarıyla akar; sayfa sınırları "column-span: all" ara parçalarıyla kesilir.
import { repeatHeaderHeight } from "./tableTools";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { EditorView } from "@tiptap/pm/view";
import type { SecGeom } from "./sections";

export interface PageInfo {
  /** Sayfanın sayfa kapsayıcısındaki üst konumu (px, ölçeksiz). */
  top: number;
  /** Gövde genişliği içinde sayfanın sol konumu (dar sayfalar ortalanır). */
  left: number;
  w: number;
  h: number;
  mt: number;
  mb: number;
  ml: number;
  mr: number;
  headerY: number;
  footerY: number;
  /** Bölüm sırası (0 tabanlı). */
  section: number;
  /** Bölümün ilk sayfası mı (Farklı İlk Sayfa için). */
  first: boolean;
  /** Bölümde "Farklı İlk Sayfa" açık mı. */
  titlePage?: boolean;
  /** Tek/çift sayfa bölüm sonu için eklenen boş sayfa. */
  blank?: boolean;
}

export interface LayoutInfo {
  pageCount: number;
  /** Her sayfanın ilk içerik konumu. */
  pageStarts: number[];
  /** Her sayfanın altında gösterilecek dipnot kimlikleri. */
  footnotes: string[][];
  pages: PageInfo[];
  /** Gövde genişliği (en geniş sayfa). */
  bodyW: number;
  /** Yan yana dizimde satır kaydırma mümkün mü (false: tek satır). */
  tileOk?: boolean;
}

export interface PageBreakMark {
  pos: number;
  kind: "block" | "row" | "rowhead" | "line" | "span" | "spanline";
  h: number;
  /** Yan yana sayfa diziminde yeni satır başlatır (tüm sütunları kesen sıfır yükseklikli parça). */
  tileRow?: boolean;
}

/** Bölüm düğümüne uygulanacak ek biçim (son sayfada sıralı sütun doldurma). */
export interface SectionStyle {
  pos: number;
  height: number;
}

/** Dipnot alanı ayıracı (çizgi + boşluk) yüksekliği (px). */
export const FOOTNOTE_SEP = 14;

export interface Unit {
  pos: number;
  node: PMNode;
  el: HTMLElement;
  breakPos: number;
  kind: "block" | "row" | "section";
  top: number;
  bottom: number;
  keepNext: boolean;
  pageBreak: boolean;
  colBreak: boolean;
  breakBefore: boolean;
  notes: { id: string; text: string }[];
  /** Bölüm sırası ve bölümün ilk birimi mi. */
  sec: number;
  secStart: boolean;
}

export type RawUnit = Omit<Unit, "top" | "bottom">;

export function footnotesIn(node: PMNode): { id: string; text: string }[] {
  const out: { id: string; text: string }[] = [];
  node.descendants((n) => {
    if (n.type.name === "note" && n.attrs.kind === "footnote") out.push({ id: n.attrs.id as string, text: noteKey(n) });
    return true;
  });
  return out;
}

/** Dipnot yüksekliği önbelleği için anahtar (zengin içerik ya da düz metin). */
export function noteKey(n: PMNode): string {
  const c = n.attrs.content as unknown;
  return c ? JSON.stringify(c) : ((n.attrs.text as string) ?? "");
}

const LIST_TYPES = new Set(["bulletList", "orderedList"]);

/** Ölçüm birimleri: tek sütunlu bölümlerde bloklar, liste öğeleri ve tablo satırları; çok sütunlu bölüm tek birimdir. */
export function collectUnits(view: EditorView, geoms: SecGeom[]): RawUnit[] {
  const out: RawUnit[] = [];
  let sec = 0;
  let secStart = true;
  const push = (node: PMNode, pos: number, breakPos: number, kind: RawUnit["kind"]) => {
    const el = view.nodeDOM(pos);
    if (!(el instanceof HTMLElement)) return;
    out.push({
      pos,
      node,
      el,
      breakPos,
      kind,
      keepNext: kind !== "section" && (!!node.attrs.keepNext || node.type.name === "heading"),
      pageBreak: node.type.name === "pageBreak",
      colBreak: node.type.name === "columnBreak",
      breakBefore: kind !== "section" && !!node.attrs.pageBreakBefore,
      notes: footnotesIn(node),
      sec,
      secStart,
    });
    secStart = false;
  };
  const visit = (node: PMNode, pos: number, breakPos: number) => {
    if (LIST_TYPES.has(node.type.name)) {
      node.forEach((li, off, idx) => {
        const liPos = pos + 1 + off;
        li.forEach((c, off2, idx2) => {
          const cPos = liPos + 1 + off2;
          const bp = idx2 === 0 ? (idx === 0 ? breakPos : liPos) : cPos;
          visit(c, cPos, bp);
        });
      });
      return;
    }
    if (node.type.name === "table") {
      node.forEach((row, off, idx) => {
        const rPos = pos + 1 + off;
        push(row, rPos, idx === 0 ? breakPos : rPos, idx === 0 ? "block" : "row");
      });
      return;
    }
    push(node, pos, breakPos, "block");
  };
  view.state.doc.forEach((section, secOff, si) => {
    sec = si;
    secStart = true;
    const g = geoms[si];
    if (g && g.cols > 1) {
      // Bölümden önceki boşluk bölümler arasına (sütunların dışına) konur.
      push(section, secOff, secOff, "section");
      return;
    }
    section.forEach((child, off) => visit(child, secOff + 1 + off, secOff + 1 + off));
  });
  return out;
}

/** Çok sütunlu bölümün dar (tek sütun) ölçümündeki birimleri. */
export function collectSectionChildren(view: EditorView, section: PMNode, secOff: number): RawUnit[] {
  const out: RawUnit[] = [];
  section.forEach((child, off) => {
    const pos = secOff + 1 + off;
    const el = view.nodeDOM(pos);
    if (!(el instanceof HTMLElement)) return;
    out.push({
      pos,
      node: child,
      el,
      breakPos: pos,
      kind: "block",
      keepNext: !!child.attrs.keepNext || child.type.name === "heading",
      pageBreak: child.type.name === "pageBreak",
      colBreak: child.type.name === "columnBreak",
      breakBefore: !!child.attrs.pageBreakBefore,
      notes: footnotesIn(child),
      sec: 0,
      secStart: false,
    });
  });
  return out;
}

const marginTopOf = (el: HTMLElement) => parseFloat(getComputedStyle(el).marginTop) || 0;

export interface FlowCtx {
  view: EditorView;
  units: Unit[];
  geoms: SecGeom[];
  bodyW: number;
  gap: number;
  /** Ekran y → sayfa kapsayıcısı y. */
  Y: (sy: number) => number;
  /** Sayfa kapsayıcısı y → ekran y. */
  screenY: (y: number) => number;
  scale: number;
  noteH: (key: string) => number;
  /** Birden Çok Sayfa: her sayfa en büyük sayfa yüksekliğinde bir yuvaya yerleşir (0 = kapalı). */
  slotH: number;
  /** Çok sütunlu bölüm: dar ölçüm ve sayfalara bölme (dışarıda uygulanır). */
  columns: (u: Unit, ctx: ColumnCtx) => ColumnResult;
}

export interface ColumnCtx {
  /** Bölümün (kaydırılmış) üst konumu. */
  top: number;
  /** Geçerli sayfa sırası ve içerik sonu hesaplayıcısı. */
  page: number;
  contentEnd: (p: number) => number;
  nextPage: () => number;
  pages: PageInfo[];
  /** Bölümden sonra sayfa sonu mu geliyor (sonraki bölüm yeni sayfada başlıyor ya da son bölüm). */
  endsPage: boolean;
}

export interface ColumnResult {
  breaks: PageBreakMark[];
  /** Bölümün yeni alt konumu (kaydırılmış koordinat). */
  bottom: number;
  page: number;
  style?: SectionStyle;
  starts: { page: number; pos: number }[];
}

export function layoutFlow(ctx: FlowCtx) {
  const { view, units, geoms, bodyW, gap, Y } = ctx;
  const breaks: PageBreakMark[] = [];
  const styles: SectionStyle[] = [];
  const pages: PageInfo[] = [];
  const pageStarts: number[] = [];
  const footnotes: string[][] = [];
  const footUsed: number[] = [];
  const unitPage: number[] = [];
  const sBefore: number[] = [];

  const addPage = (g: SecGeom, first: boolean, start: number): number => {
    const prev = pages[pages.length - 1];
    const top = prev ? prev.top + (ctx.slotH || prev.h) + gap : 0;
    pages.push({ top, left: (bodyW - g.w) / 2, w: g.w, h: g.h, mt: g.mt, mb: g.mb, ml: g.ml, mr: g.mr, headerY: g.headerY, footerY: g.footerY, section: g.index, first, titlePage: g.titlePage });
    pageStarts.push(start);
    footnotes.push([]);
    footUsed.push(0);
    return pages.length - 1;
  };
  const contentEnd = (p: number) => pages[p].top + pages[p].h - pages[p].mb - footUsed[p];
  addPage(geoms[0], true, 0);

  let S = 0;
  let forceNext = false;

  const outerTop = (u: Unit) => {
    const el = view.nodeDOM(u.breakPos);
    if (el instanceof HTMLElement && el !== u.el) return Y(el.getBoundingClientRect().top) - marginTopOf(el);
    return u.top - marginTopOf(u.el);
  };
  const notesH = (u: Unit) => u.notes.reduce((sum, n) => sum + ctx.noteH(n.text), 0);
  const record = (k: number, page: number) => {
    const u = units[k];
    unitPage[k] = page;
    if (!u.notes.length) return;
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
  /** Tek/çift sayfa bölüm sonu: yeni bölümün ilk sayfa numarası uymuyorsa önce boş sayfa eklenir. */
  const parityFix = (g: SecGeom, start: number) => {
    const want = g.kind === "evenPage" ? 0 : g.kind === "oddPage" ? 1 : -1;
    if (want >= 0 && (pages.length + 1) % 2 !== want) {
      const b = addPage(g, false, start);
      pages[b].blank = true;
    }
  };
  /** Bölümün ilk sayfası olarak yeni sayfa (gerekirse önce boş sayfa). */
  const sectionPage = (g: SecGeom, start: number, secForce: boolean) => {
    if (secForce) parityFix(g, start);
    return addPage(g, secForce, start);
  };
  const sameSize = (pg: PageInfo, g: SecGeom) => Math.abs(pg.w - g.w) < 0.5 && Math.abs(pg.h - g.h) < 0.5;
  const sameGeom = (pg: PageInfo, g: SecGeom) => sameSize(pg, g) && Math.abs(pg.mt - g.mt) < 0.5;

  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    const g = geoms[u.sec] ?? geoms[0];
    sBefore[i] = S;
    let p = pages.length - 1;
    const top = u.top + S;
    const bottom = u.bottom + S;
    // Önceki uzun birim sayfalar boyunca taşmışsa sayfa ekle.
    while (top > pages[p].top + pages[p].h + gap / 2) p = addPage(g, false, u.pos);
    // Sürekli bölüm farklı boyutlu sayfaya geçemez: sonraki sayfada başlar.
    const sizeChange = u.secStart && u.sec > 0 && !sameSize(pages[p], g);
    const secForce = u.secStart && u.sec > 0 && (g.kind !== "continuous" || sizeChange);
    // Açık sayfa sonu her zaman sayfayı bitirir (sayfanın tek içeriği olsa bile; Word'deki gibi boş sayfa kalır).
    const hard = forceNext;
    const force = forceNext || u.breakBefore || secForce;
    forceNext = u.pageBreak || (u.colBreak && g.cols === 1);
    const ownNotes = u.notes.length ? notesH(u) + (footnotes[p].length ? 0 : FOOTNOTE_SEP) : 0;
    const end = contentEnd(p) - ownNotes;
    const inMargin = top > end - 0.5;
    const oTop = () => outerTop(u) + S;
    const atTop = () => oTop() <= pages[p].top + pages[p].mt + 1;
    // Zorunlu kesme ama birim zaten sayfanın başında: yeni sayfa gerekmez (bölüm başında geometri/parite uymalı).
    const startHere = force && !hard && atTop() && (!secForce || (sameGeom(pages[p], g) && (g.kind === "nextPage" || g.kind === "continuous" || (pages.length % 2 === 0) === (g.kind === "evenPage"))));

    if (u.kind === "section") {
      // Çok sütunlu bölüm
      let sp = p;
      if (startHere && secForce) {
        pages[p].first = true;
        pages[p].section = u.sec;
        pages[p].titlePage = g.titlePage;
      } else if (force && !startHere) {
        sp = sectionPage(g, u.pos, secForce);
        const h = pages[sp].top + pages[sp].mt - oTop();
        if (h > 0) {
          breaks.push({ pos: u.breakPos, kind: "block", h });
          S += h;
        }
      }
      const nextSec = geoms[u.sec + 1];
      const res = ctx.columns(u, {
        top: u.top + S,
        page: sp,
        contentEnd,
        nextPage: () => addPage(g, false, -1),
        pages,
        endsPage: !nextSec || nextSec.kind !== "continuous" || Math.abs(nextSec.w - g.w) > 0.5 || Math.abs(nextSec.h - g.h) > 0.5,
      });
      for (const b of res.breaks) breaks.push(b);
      for (const st of res.starts) pageStarts[st.page] = st.pos;
      if (res.style) styles.push(res.style);
      unitPage[i] = res.page;
      S = res.bottom - u.bottom;
      continue;
    }

    const fits = !force && !inMargin && bottom <= end + 0.5;
    const textUnit = !u.notes.length && u.kind === "block" && u.node.isTextblock && u.node.content.size > 0;
    const splittable = textUnit && !u.node.attrs.keepLines;
    if (fits || startHere || (!force && atTop())) {
      if (startHere && secForce) {
        pages[p].first = true;
        pages[p].section = u.sec;
        pages[p].titlePage = g.titlePage;
      }
      // Sayfa başında başlayan ve sayfadan uzun paragraf satır sınırından bölünür.
      // (Satırları birlikte tut sayfadan uzun paragrafta uygulanamaz.)
      if (!fits && textUnit && bottom > end + 0.5) {
        const split = splitParagraph(ctx, u, S, p, pages, contentEnd, () => addPage(g, false, -1));
        if (split) {
          for (const b of split.breaks) breaks.push(b);
          split.breaks.forEach((b, k) => (pageStarts[p + 1 + k] = b.pos));
          S += split.added;
          record(i, p);
          continue;
        }
      }
      record(i, p);
      // Sayfadan uzun birim: kapsadığı sayfalar eklenir.
      while (bottom > pages[pages.length - 1].top + pages[pages.length - 1].h) addPage(g, false, u.pos);
      continue;
    }

    // Paragrafı satır düzeyinde bölmeyi dene (dipnot içermeyenlerde).
    if (!force && !inMargin && splittable) {
      const split = splitParagraph(ctx, u, S, p, pages, contentEnd, () => addPage(g, false, -1));
      if (split) {
        for (const b of split.breaks) breaks.push(b);
        split.breaks.forEach((b, k) => (pageStarts[p + 1 + k] = b.pos));
        S += split.added;
        record(i, p);
        continue;
      }
    }

    // keepNext zinciri: başlık bir sonraki paragrafla birlikte kalsın.
    let target = i;
    if (!force) {
      let j = i - 1;
      while (j >= 0 && units[j].keepNext && !units[j].pageBreak && unitPage[j] === p) j--;
      const cand = j + 1;
      if (cand < i) {
        const ct = outerTop(units[cand]) + sBefore[cand];
        if (ct > pages[p].top + pages[p].mt + 1) target = cand;
      }
    }
    for (let k = target; k < i; k++) unrecord(k);
    S = sBefore[target];
    const t = units[target];
    const tg = geoms[t.sec] ?? g;
    const before = pages.length;
    const np = sectionPage(tg, t.pos, secForce && target === i);
    const h = pages[np].top + pages[np].mt - (outerTop(t) + S);
    if (h <= 0) {
      record(i, p);
      continue;
    }
    const existing = breaks.find((b) => b.pos === t.breakPos && (b.kind === "block" || b.kind === "row"));
    if (existing) existing.h += h;
    else breaks.push({ pos: t.breakPos, kind: t.kind === "row" ? "row" : "block", h });
    S += h;
    // Başlık Satırlarını Yinele: yeni sayfadaki ilk satırın üstüne başlık satırının kopyası gelir.
    const hh = t.kind === "row" ? repeatHeaderHeight(view, t.pos, ctx.scale) : 0;
    if (hh > 0 && !breaks.some((b) => b.pos === t.breakPos && b.kind === "rowhead")) {
      breaks.push({ pos: t.breakPos, kind: "rowhead", h: hh });
      S += hh;
    }
    for (let k = target; k <= i; k++) {
      sBefore[k] = S;
      record(k, np);
    }
    for (let q = before; q <= np; q++) pageStarts[q] = t.pos;
  }

  // Eksik başlangıçlar (boş sayfalar) bir sonraki sayfanınkini alır.
  for (let q = pageStarts.length - 1; q >= 0; q--) if (pageStarts[q] < 0) pageStarts[q] = pageStarts[q + 1] ?? pageStarts[q - 1] ?? 0;
  pageStarts[0] = 0;
  for (let q = 1; q < pageStarts.length; q++) if (pageStarts[q] < pageStarts[q - 1]) pageStarts[q] = pageStarts[q - 1];
  const rank = (k: PageBreakMark["kind"]) => (k === "line" ? 2 : k === "rowhead" ? 1 : 0);
  breaks.sort((a, b) => a.pos - b.pos || rank(a.kind) - rank(b.kind));
  return { breaks, styles, info: { pageCount: pages.length, pageStarts, footnotes, pages, bodyW } as LayoutInfo };
}

/** Ekran koordinatında satır bölme noktası: alt kenarı `limit`i aşan ilk satırın başı. */
export function lineSplitAt(view: EditorView, u: Unit, from: number, limitScreen: number) {
  const start = u.pos + 1;
  const end = u.pos + u.node.nodeSize - 1;
  const coords = (p: number, side = 1) => {
    try {
      return view.coordsAtPos(p, side);
    } catch {
      return null;
    }
  };
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
  const p = firstPosBelow(Math.max(start, from), limitScreen);
  if (p < 0) return null;
  const pc = coords(p);
  if (!pc) return null;
  const q = firstPosBelow(Math.max(start, from), pc.top + 1);
  if (q < 0) return null;
  return { pos: q, top: coords(q)?.top ?? pc.top, start, end, coords, firstPosBelow };
}

/** Uzun paragrafı satır sınırından böler (dul/öksüz satır denetimiyle). */
function splitParagraph(ctx: FlowCtx, u: Unit, S0: number, page0: number, pages: PageInfo[], contentEnd: (p: number) => number, nextPage: () => number) {
  const { view, Y, scale, screenY } = ctx;
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
  const paraBottomScreen = screenY(u.bottom);
  const breaks: PageBreakMark[] = [];
  let S = S0;
  let page = page0;
  let from = start;
  let lineTopFirst = first.top;
  for (let guard = 0; guard < 200; guard++) {
    const cEnd = contentEnd(page);
    if (u.bottom + S <= cEnd + 0.5) break;
    const split = lineSplitAt(view, u, from, screenY(cEnd - S));
    if (!split) break;
    let { pos: q, top: T } = split;
    // Tek satır sonraki sayfaya kalıyorsa bir satır geri al.
    if (paraBottomScreen - T < lineH * 1.5 && q > from) {
      const prev = coords(q - 1, -1);
      if (prev) {
        const q2 = split.firstPosBelow(from, prev.top + 1);
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
    const np = nextPage();
    const h = pages[np].top + pages[np].mt - lineTop;
    if (h <= 0) break;
    breaks.push({ pos: q, kind: "line", h });
    S += h;
    page = np;
    from = q;
    lineTopFirst = T;
  }
  if (!breaks.length) return null;
  return { breaks, added: S - S0, page };
}
