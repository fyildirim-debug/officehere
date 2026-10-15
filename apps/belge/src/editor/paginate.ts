// Sayfalama düzen birimi (DOM'suz): ölçülmüş doğal blok konumlarından sayfa listesi ve sayfa boşluğu
// parçacıkları. Sayfalar sıralı oluşturulur; her sayfa bölümünün boyut/kenar boşluğunu taşır. Çok sütunlu
// bölümler CSS sütunlarıyla akar; sayfa sınırları "column-span: all" ara parçalarıyla kesilir. DOM ölçümü
// `FlowMeasure` arayüzüyle dışarıdan verilir (testte taklit edilir).
//
// Artımlı çalışma: her birimin başındaki akış durumu (anlık görüntü) önbellekte tutulur. Yeni çalışmada
// değişmeyen ön ek atlanır; değişen ilk birimden önceki son sayfa başından devam edilir ve değişmeyen son
// ekteki bir birimde durum (sayfa, sayfa içi konum, dipnotlar) öncekiyle örtüşünce önceki sonucun kalanı
// konumları kaydırılarak eklenir. Sonuç her durumda tam hesapla aynıdır.
import type { Node as PMNode } from "@tiptap/pm/model";
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
  /** Birden Çok Sayfa: satırların ilk sayfaları, sütun sayısı ve hesaplandığı hedef satır başına sayfa. */
  tileRows?: number[];
  tileCols?: number;
  tileTarget?: number;
}

export interface PageBreakMark {
  pos: number;
  kind: "block" | "row" | "rowhead" | "line" | "span" | "spanline";
  h: number;
  /** Yan yana sayfa diziminde yeni satır başlatır (tüm sütunları kesen sıfır yükseklikli parça). */
  tileRow?: boolean;
  /** Satır erken bittiyse boş sayfa yuvalarını dolduran ek yükseklik. */
  tileFill?: number;
}

/** Bölüm düğümüne uygulanacak ek biçim (son sayfada sıralı sütun doldurma). */
export interface SectionStyle {
  pos: number;
  height: number;
}

/** Dipnot alanı ayıracı (çizgi + boşluk) yüksekliği (px). */
export const FOOTNOTE_SEP = 14;

/** Düzen birimi: tek sütunlu bölümlerde blok, liste paragrafı ya da tablo satırı; çok sütunlu bölüm tek birimdir. */
export interface FlowUnit {
  /** Kimlik: değişmeyen düğüm aynı nesne kalır (ProseMirror yapısal paylaşımı). */
  key: object;
  /** Görünümü etkileyen ataların (liste, tablo) öznitelik nesneleri; kimlikle karşılaştırılır. */
  ctx: readonly unknown[];
  pos: number;
  breakPos: number;
  /** Düğüm sonu (pos + nodeSize). */
  end: number;
  kind: "block" | "row" | "section";
  /** Doğal (ara parçasız) konumlar, sayfa kapsayıcısı koordinatında. */
  top: number;
  bottom: number;
  keepNext: boolean;
  pageBreak: boolean;
  colBreak: boolean;
  breakBefore: boolean;
  keepLines: boolean;
  /** İçerikli metin bloğu (satır sınırından bölünebilir). */
  text: boolean;
  notes: { id: string; text: string }[];
  /** Bölüm sırası ve bölümün ilk birimi mi. */
  sec: number;
  secStart: boolean;
  /** Tablo satırı: yeni sayfada yinelenen başlık satırının yüksekliği (0 = yok). */
  hh: number;
  /** Ölçüm katmanının önbelleği: kesme kapsayıcısının dış üst kenarının birimin üstüne uzaklığı. */
  ot?: number;
}

export interface Rect {
  top: number;
  bottom: number;
}

/** Düzen biriminin DOM'dan istediği ölçümler (doğal, ara parçasız düzende). */
export interface FlowMeasure {
  /** Birimin (kesme konumundaki kapsayıcısıyla) kenar boşluğu dahil üst konumu (kapsayıcı koordinatı). */
  outerTop(i: number): number;
  /** Birim içindeki konumun ekran dikdörtgeni (`coordsAtPos`); ölçülemezse null. */
  coords(i: number, pos: number, side: number): Rect | null;
  /** Hesaplanan satır yüksekliği (CSS px; "normal" ise NaN). */
  lineHeight(i: number): number;
  /** Dipnot satırının yüksekliği. */
  noteH(key: string): number;
  /** Çok sütunlu bölüm: dar ölçüm ve sayfalara bölme. */
  columns(i: number, c: ColumnCtx): ColumnResult;
}

export interface FlowCtx {
  units: FlowUnit[];
  geoms: SecGeom[];
  bodyW: number;
  gap: number;
  /** Ekran y → sayfa kapsayıcısı y. */
  Y: (sy: number) => number;
  /** Sayfa kapsayıcısı y → ekran y. */
  screenY: (y: number) => number;
  scale: number;
  /** Birden Çok Sayfa: her sayfa en büyük sayfa yüksekliğinde bir yuvaya yerleşir (0 = kapalı). */
  slotH: number;
  m: FlowMeasure;
  /** Kaydı aynı olsa da değişmiş sayılacak birimler (ör. DOM'u belge dışından değişen; 1 = değişti). */
  changed?: Uint8Array;
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

/** Birim başındaki akış durumu. */
interface Snap {
  S: number;
  force: boolean;
  nPages: number;
  nBreaks: number;
  nStyles: number;
  /** Geçerli (son) sayfa ve dipnot durumu (sonraki birimler değiştirebilir). */
  cur: PageInfo;
  foot: string[];
  footUsed: number;
  /** Bu birimden geriye birlikte tut (keepNext) zinciri uzanamaz. */
  free: boolean;
}

/** Önceki çalışmanın artımlı devam için sakladıkları (diziler akış sırasındaki ham hâlleriyle). */
export interface FlowCache {
  env: string;
  units: FlowUnit[];
  index: Map<object, number>;
  pages: PageInfo[];
  pageStarts: number[];
  footnotes: string[][];
  footUsed: number[];
  breaks: PageBreakMark[];
  styles: SectionStyle[];
  unitPage: number[];
  sBefore: number[];
  /** Her birim başının (ve sonun) durumu: snaps.length === units.length + 1. */
  snaps: Snap[];
}

export interface FlowStats {
  /** Hesaba başlanan birim (öncesi önbellekten). */
  from: number;
  /** Hesabın durduğu birim (sonrası önbellekten kaydırılarak). */
  to: number;
  incremental: boolean;
}

export interface FlowResult {
  breaks: PageBreakMark[];
  styles: SectionStyle[];
  info: LayoutInfo;
  cache: FlowCache;
  stats: FlowStats;
}

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

/** Konumların yuvarlama farkı toleransı (px). */
const EPS = 1e-6;
const near = (a: number, b: number) => Math.abs(a - b) < EPS;

/** Düzeni etkileyen ortam: bölüm geometrileri (konumlar hariç), gövde genişliği, aralık, ölçek. */
export function flowEnv(geoms: SecGeom[], bodyW: number, gap: number, slotH: number, scale: number): string {
  const g = geoms.map((x) => [x.index, x.kind, x.w, x.h, x.mt, x.mb, x.ml, x.mr, x.headerY, x.footerY, x.contentW, x.cols, x.colGap, x.colW, x.titlePage ? 1 : 0].join(","));
  return `${g.join(";")}|${bodyW}|${gap}|${slotH}|${scale}`;
}

/** Ölçüm dışındaki içerik aynı mı (düğüm kimliği, ata öznitelikleri, bayraklar, dipnotlar, başlık satırı). */
function sameUnit(a: FlowUnit, b: FlowUnit): boolean {
  if (a.key !== b.key || a.kind !== b.kind || a.sec !== b.sec || a.secStart !== b.secStart || a.hh !== b.hh) return false;
  if (a.end - a.pos !== b.end - b.pos || a.breakPos - a.pos !== b.breakPos - b.pos) return false;
  if (a.keepNext !== b.keepNext || a.pageBreak !== b.pageBreak || a.colBreak !== b.colBreak || a.breakBefore !== b.breakBefore || a.keepLines !== b.keepLines || a.text !== b.text) return false;
  if (a.ctx.length !== b.ctx.length || a.ctx.some((c, k) => c !== b.ctx[k])) return false;
  return a.notes.length === b.notes.length && a.notes.every((n, k) => n.id === b.notes[k].id && n.text === b.notes[k].text);
}

/** Yeni birimlerin öncekilerle eşleşmesi (sıra korunur; -1: yeni ya da değişmiş). */
export function alignUnits(units: FlowUnit[], prev: { units: FlowUnit[]; index: Map<object, number> }): Int32Array {
  const align = new Int32Array(units.length).fill(-1);
  let last = -1;
  for (let i = 0; i < units.length; i++) {
    const j = prev.index.get(units[i].key);
    if (j === undefined || j <= last || !sameUnit(prev.units[j], units[i])) continue;
    align[i] = j;
    last = j;
  }
  return align;
}

const samePage = (a: PageInfo, b: PageInfo) =>
  near(a.top, b.top) &&
  near(a.left, b.left) &&
  near(a.w, b.w) &&
  near(a.h, b.h) &&
  near(a.mt, b.mt) &&
  near(a.mb, b.mb) &&
  near(a.ml, b.ml) &&
  near(a.mr, b.mr) &&
  a.headerY === b.headerY &&
  a.footerY === b.footerY &&
  a.section === b.section &&
  a.first === b.first &&
  !!a.titlePage === !!b.titlePage &&
  !!a.blank === !!b.blank;

interface Reuse {
  prev: FlowCache;
  align: Int32Array;
  /** Devam edilecek birim (önceki çalışmanın geçerli anlık görüntüsü). */
  from: number;
  /** Örtüşme denetiminin başladığı birim: buradan sonrası öncekiyle aynı kaydırmayla eşleşir. */
  tail: number;
  dPos: number;
  dTop: number;
}

/** Artımlı devam planı: değişmeyen ön ek, devam noktası ve aynı kaydırmalı son ek. */
function planReuse(units: FlowUnit[], prev: FlowCache, changed?: Uint8Array): Reuse {
  const n = units.length;
  const on = prev.units.length;
  const old = prev.units;
  const align = alignUnits(units, prev);
  if (changed) for (let i = 0; i < n; i++) if (changed[i]) align[i] = -1;
  // Ön ek: aynı birim, aynı konum ve aynı doğal yerleşim.
  let d = 0;
  const lim = Math.min(n, on);
  while (d < lim && align[d] === d && units[d].pos === old[d].pos && units[d].top === old[d].top && units[d].bottom === old[d].bottom) d++;
  // Son ek: sona kadar ardışık eşleşen ve hepsi aynı miktar kaymış birimler.
  let tail = n;
  let dPos = 0;
  let dTop = 0;
  if (n > 0 && on > 0 && align[n - 1] === on - 1) {
    dPos = units[n - 1].pos - old[on - 1].pos;
    dTop = units[n - 1].top - old[on - 1].top;
    const ok = (k: number) => {
      const o = old[align[k]];
      return units[k].pos - o.pos === dPos && near(units[k].top - o.top, dTop) && near(units[k].bottom - o.bottom, dTop);
    };
    if (ok(n - 1)) {
      tail = n - 1;
      while (tail > 0 && align[tail - 1] >= 0 && align[tail - 1] === align[tail] - 1 && ok(tail - 1)) tail--;
    }
  }
  // Devam noktası: d'den önceki son sayfa başı birimi (sonraki birimler bundan öncekileri hiç değiştirmez)
  // ya da önceki çalışmanın sonu.
  let from = d;
  while (from > 0 && from < on && !(prev.unitPage[from] > prev.unitPage[from - 1])) from--;
  return { prev, align, from, tail: Math.max(tail, from), dPos, dTop };
}

/**
 * Sayfalama akışı. `prev` verilirse ve ortam aynıysa önceki sonuçtan artımlı devam edilir; sonuç tam
 * hesapla birebir aynıdır.
 */
export function layoutFlow(ctx: FlowCtx, prev?: FlowCache | null): FlowResult {
  const { units, geoms, bodyW, gap, m } = ctx;
  const n = units.length;
  const env = flowEnv(geoms, bodyW, gap, ctx.slotH, ctx.scale);
  let breaks: PageBreakMark[] = [];
  let styles: SectionStyle[] = [];
  let pages: PageInfo[] = [];
  let pageStarts: number[] = [];
  let footnotes: string[][] = [];
  let footUsed: number[] = [];
  let unitPage: number[] = [];
  let sBefore: number[] = [];
  let snaps: Snap[] = [];
  let S = 0;
  let forceNext = false;

  const addPage = (g: SecGeom, first: boolean, start: number): number => {
    const prevPage = pages[pages.length - 1];
    const top = prevPage ? prevPage.top + (ctx.slotH || prevPage.h) + gap : 0;
    pages.push({ top, left: (bodyW - g.w) / 2, w: g.w, h: g.h, mt: g.mt, mb: g.mb, ml: g.ml, mr: g.mr, headerY: g.headerY, footerY: g.footerY, section: g.index, first, titlePage: g.titlePage });
    pageStarts.push(start);
    footnotes.push([]);
    footUsed.push(0);
    return pages.length - 1;
  };
  const contentEnd = (p: number) => pages[p].top + pages[p].h - pages[p].mb - footUsed[p];

  const notesH = (u: FlowUnit) => u.notes.reduce((sum, nt) => sum + m.noteH(nt.text), 0);
  const record = (k: number, page: number) => {
    const u = units[k];
    unitPage[k] = page;
    if (!u.notes.length) return;
    footUsed[page] += notesH(u) + (footnotes[page].length ? 0 : FOOTNOTE_SEP);
    for (const nt of u.notes) footnotes[page].push(nt.id);
  };
  const unrecord = (k: number) => {
    const u = units[k];
    const page = unitPage[k];
    if (page === undefined || !u.notes.length || !footnotes[page]) return;
    footnotes[page] = footnotes[page].filter((id) => !u.notes.some((nt) => nt.id === id));
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
  /** Bu birimden geriye keepNext zinciri uzanamaz (önceki birim zincire katılamaz). */
  const freeAt = (i: number) => i === 0 || !(units[i - 1].keepNext && !units[i - 1].pageBreak && unitPage[i - 1] === pages.length - 1);
  const snap = (free: boolean): Snap => {
    const last = pages.length - 1;
    return { S, force: forceNext, nPages: pages.length, nBreaks: breaks.length, nStyles: styles.length, cur: { ...pages[last] }, foot: footnotes[last].slice(), footUsed: footUsed[last], free };
  };

  const place = (i: number) => {
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
    const oTop = () => m.outerTop(i) + S;
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
      const res = m.columns(i, {
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
      return;
    }

    const fits = !force && !inMargin && bottom <= end + 0.5;
    const textUnit = !u.notes.length && u.kind === "block" && u.text;
    const splittable = textUnit && !u.keepLines;
    if (fits || startHere || (!force && atTop())) {
      if (startHere && secForce) {
        pages[p].first = true;
        pages[p].section = u.sec;
        pages[p].titlePage = g.titlePage;
      }
      // Sayfa başında başlayan ve sayfadan uzun paragraf satır sınırından bölünür.
      // (Satırları birlikte tut sayfadan uzun paragrafta uygulanamaz.)
      if (!fits && textUnit && bottom > end + 0.5) {
        const split = splitParagraph(ctx, i, S, p, pages, contentEnd, () => addPage(g, false, -1));
        if (split) {
          for (const b of split.breaks) breaks.push(b);
          split.breaks.forEach((b, k) => (pageStarts[p + 1 + k] = b.pos));
          S += split.added;
          record(i, p);
          return;
        }
      }
      record(i, p);
      // Sayfadan uzun birim: kapsadığı sayfalar eklenir.
      while (bottom > pages[pages.length - 1].top + pages[pages.length - 1].h) addPage(g, false, u.pos);
      return;
    }

    // Paragrafı satır düzeyinde bölmeyi dene (dipnot içermeyenlerde).
    if (!force && !inMargin && splittable) {
      const split = splitParagraph(ctx, i, S, p, pages, contentEnd, () => addPage(g, false, -1));
      if (split) {
        for (const b of split.breaks) breaks.push(b);
        split.breaks.forEach((b, k) => (pageStarts[p + 1 + k] = b.pos));
        S += split.added;
        record(i, p);
        return;
      }
    }

    // keepNext zinciri: başlık bir sonraki paragrafla birlikte kalsın.
    let target = i;
    if (!force) {
      let j = i - 1;
      while (j >= 0 && units[j].keepNext && !units[j].pageBreak && unitPage[j] === p) j--;
      const cand = j + 1;
      if (cand < i) {
        const ct = m.outerTop(cand) + sBefore[cand];
        if (ct > pages[p].top + pages[p].mt + 1) target = cand;
      }
    }
    for (let k = target; k < i; k++) unrecord(k);
    S = sBefore[target];
    const t = units[target];
    const tg = geoms[t.sec] ?? g;
    const before = pages.length;
    const np = sectionPage(tg, t.pos, secForce && target === i);
    const h = pages[np].top + pages[np].mt - (m.outerTop(target) + S);
    if (h <= 0) {
      record(i, p);
      return;
    }
    const existing = breaks.find((b) => b.pos === t.breakPos && (b.kind === "block" || b.kind === "row"));
    if (existing) existing.h += h;
    else breaks.push({ pos: t.breakPos, kind: t.kind === "row" ? "row" : "block", h });
    S += h;
    // Başlık Satırlarını Yinele: yeni sayfadaki ilk satırın üstüne başlık satırının kopyası gelir.
    const hh = t.kind === "row" ? t.hh : 0;
    if (hh > 0 && !breaks.some((b) => b.pos === t.breakPos && b.kind === "rowhead")) {
      breaks.push({ pos: t.breakPos, kind: "rowhead", h: hh });
      S += hh;
    }
    for (let k = target; k <= i; k++) {
      sBefore[k] = S;
      record(k, np);
    }
    for (let q = before; q <= np; q++) pageStarts[q] = t.pos;
  };

  let start = 0;
  let reuse: Reuse | null = null;
  if (prev && prev.env === env && prev.snaps.length === prev.units.length + 1) {
    reuse = planReuse(units, prev, ctx.changed);
    // Devam noktasındaki durum geri yüklenir (sayfa başından öncesi sonradan hiç değişmez).
    const f = reuse.from;
    const sn = prev.snaps[f];
    const last = sn.nPages - 1;
    pages = prev.pages.slice(0, sn.nPages).map((pg) => ({ ...pg }));
    pages[last] = { ...sn.cur };
    footnotes = prev.footnotes.slice(0, sn.nPages).map((a) => a.slice());
    footnotes[last] = sn.foot.slice();
    footUsed = prev.footUsed.slice(0, sn.nPages);
    footUsed[last] = sn.footUsed;
    pageStarts = prev.pageStarts.slice(0, sn.nPages);
    breaks = prev.breaks.slice(0, sn.nBreaks).map((b) => ({ ...b }));
    styles = prev.styles.slice(0, sn.nStyles).map((s) => ({ ...s }));
    unitPage = prev.unitPage.slice(0, f);
    sBefore = prev.sBefore.slice(0, f);
    snaps = prev.snaps.slice(0, f);
    S = sn.S;
    forceNext = sn.force;
    start = f;
  } else addPage(geoms[0], true, 0);

  let to = n;
  for (let i = start; i < n; i++) {
    const free = freeAt(i);
    if (reuse && i >= reuse.tail && splice(i, free, reuse)) {
      to = i;
      break;
    }
    snaps.push(snap(free));
    place(i);
  }
  if (to === n) snaps.push(snap(freeAt(n)));

  /** Durum önceki çalışmadaki eşiyle örtüşüyorsa kalan sonuç önceki çalışmadan kaydırılarak eklenir. */
  function splice(i: number, free: boolean, r: Reuse): boolean {
    const old = r.prev;
    const j = r.align[i];
    const sj = old.snaps[j];
    const last = pages.length - 1;
    if (!free || !sj.free || sj.nPages !== pages.length || sj.force !== forceNext || !near(S, sj.S - r.dTop)) return false;
    if (!near(footUsed[last], sj.footUsed) || footnotes[last].length !== sj.foot.length || footnotes[last].some((id, k) => id !== sj.foot[k]) || !samePage(pages[last], sj.cur)) return false;
    const dB = breaks.length - sj.nBreaks;
    const dS = styles.length - sj.nStyles;
    const mapPos = (p: number) => (p < 0 ? p : p + r.dPos);
    pages[last] = { ...old.pages[last] };
    footnotes[last] = old.footnotes[last].slice();
    footUsed[last] = old.footUsed[last];
    for (let q = sj.nPages; q < old.pages.length; q++) {
      pages.push({ ...old.pages[q] });
      footnotes.push(old.footnotes[q].slice());
      footUsed.push(old.footUsed[q]);
      pageStarts.push(mapPos(old.pageStarts[q]));
    }
    for (let k = sj.nBreaks; k < old.breaks.length; k++) breaks.push({ ...old.breaks[k], pos: old.breaks[k].pos + r.dPos });
    for (let k = sj.nStyles; k < old.styles.length; k++) styles.push({ ...old.styles[k], pos: old.styles[k].pos + r.dPos });
    for (let k = j; k < old.units.length; k++) {
      unitPage.push(old.unitPage[k]);
      sBefore.push(old.sBefore[k] - r.dTop);
    }
    // Anlık görüntüler değişmez: kayma yoksa aynı nesneler kullanılır.
    const same = r.dTop === 0 && dB === 0 && dS === 0;
    for (let k = j; k < old.snaps.length; k++) {
      const o = old.snaps[k];
      snaps.push(same ? o : { ...o, S: o.S - r.dTop, nBreaks: o.nBreaks + dB, nStyles: o.nStyles + dS });
    }
    return true;
  }

  // Eksik başlangıçlar (boş sayfalar) bir sonraki sayfanınkini alır.
  const starts = pageStarts.slice();
  for (let q = starts.length - 1; q >= 0; q--) if (starts[q] < 0) starts[q] = starts[q + 1] ?? starts[q - 1] ?? 0;
  starts[0] = 0;
  for (let q = 1; q < starts.length; q++) if (starts[q] < starts[q - 1]) starts[q] = starts[q - 1];
  const rank = (k: PageBreakMark["kind"]) => (k === "line" ? 2 : k === "rowhead" ? 1 : 0);
  const outBreaks = breaks.map((b) => ({ ...b })).sort((a, b) => a.pos - b.pos || rank(a.kind) - rank(b.kind));
  const index = new Map<object, number>();
  units.forEach((u, k) => {
    if (!index.has(u.key)) index.set(u.key, k);
  });
  return {
    breaks: outBreaks,
    styles: styles.map((s) => ({ ...s })),
    info: { pageCount: pages.length, pageStarts: starts, footnotes: footnotes.map((a) => a.slice()), pages: pages.map((pg) => ({ ...pg })), bodyW },
    cache: { env, units, index, pages, pageStarts, footnotes, footUsed, breaks, styles, unitPage, sBefore, snaps },
    stats: { from: start, to, incremental: !!reuse },
  };
}

/** Ekran koordinatında satır bölme noktası: alt kenarı `limit`i aşan ilk satırın başı. */
export function lineSplitAt(coords: (p: number, side?: number) => Rect | null, start: number, end: number, from: number, limitScreen: number) {
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
  return { pos: q, top: coords(q)?.top ?? pc.top, firstPosBelow };
}

/** Uzun paragrafı satır sınırından böler (dul/öksüz satır denetimiyle). */
function splitParagraph(ctx: FlowCtx, i: number, S0: number, page0: number, pages: PageInfo[], contentEnd: (p: number) => number, nextPage: () => number) {
  const { m, Y, scale, screenY } = ctx;
  const u = ctx.units[i];
  const start = u.pos + 1;
  const end = u.end - 1;
  if (end - start < 2) return null;
  const coords = (p: number, side = 1) => m.coords(i, p, side);
  const first = coords(start);
  if (!first) return null;
  const lhRaw = m.lineHeight(i);
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
    const split = lineSplitAt(coords, start, end, from, screenY(cEnd - S));
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
