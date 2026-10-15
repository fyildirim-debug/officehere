// Sayfalama düzen birimi (DOM'suz): ölçüm arayüzü taklit edilir. Artımlı sonuç, rastgele düzenleme
// dizilerinde her adımda tam hesapla birebir aynı olmalı; değişmeyen sayfalar yeniden hesaplanmamalı.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { layoutFlow, type ColumnCtx, type ColumnResult, type FlowCache, type FlowCtx, type FlowMeasure, type FlowResult, type FlowUnit, type PageBreakMark } from "../paginate";
import type { SecGeom } from "../sections";
import type { PageSetup, SectionKind } from "../../model/document";

/* ------------------------------------------------------------------ */
/* Taklit belge: doğal düzen ve ölçüm                                    */
/* ------------------------------------------------------------------ */

type Kind = "p" | "h" | "row" | "pb" | "cols";

interface Block {
  /** Değişen blok yeni bir kimlik nesnesi alır (ProseMirror'da yeni düğüm). */
  key: object;
  kind: Kind;
  /** Satır yükseklikleri (metin blokları) ya da tek parça yükseklik. */
  lines: number[];
  /** Satır başına karakter (konum → satır eşlemesi). */
  cpl: number;
  mt: number;
  mb: number;
  keepNext: boolean;
  keepLines: boolean;
  breakBefore: boolean;
  /** Satır yüksekliği "normal" (NaN) mı. */
  normalLh: boolean;
  notes: { id: string; text: string }[];
  /** Tablo satırı: yinelenen başlık yüksekliği. */
  hh: number;
}

interface Doc {
  secs: { kind: SectionKind; land: boolean; cols: number; blocks: Block[] }[];
}

const PAGE = {} as PageSetup;

/** Ölçüler 1/4 px katı: kayan nokta toplamları kesin, karşılaştırma birebir. */
function geom(index: number, kind: SectionKind, land: boolean, cols: number): SecGeom {
  const w = land ? 1056 : 816;
  const h = land ? 816 : 1056;
  return { index, pos: 0, end: 0, kind, page: PAGE, w, h, mt: 96, mb: 96, ml: 96, mr: 96, headerY: 48, footerY: 48, contentW: w - 192, cols, colGap: cols > 1 ? 48 : 0, colW: cols > 1 ? (w - 192 - 48) / 2 : w - 192, line: false, titlePage: index === 1 };
}

const geomsOf = (doc: Doc) => doc.secs.map((s, i) => geom(i, i === 0 ? "nextPage" : s.kind, s.land, s.cols));

const heightOf = (b: Block) => b.lines.reduce((a, x) => a + x, 0);
const sizeOf = (b: Block) => (b.kind === "p" || b.kind === "h" ? 2 + b.lines.length * b.cpl : b.kind === "cols" ? 2 + b.lines.length * 7 : 2);

/** Doğal düzen (esnek sütun: kenar boşlukları birleşmez) ve düzen birimleri. */
function unitsOf(doc: Doc): { units: FlowUnit[]; blocks: Block[] } {
  const units: FlowUnit[] = [];
  const blocks: Block[] = [];
  // İlk birim ilk sayfanın üst kenar boşluğunun altında başlar.
  let y = 96;
  let pos = 0;
  doc.secs.forEach((s, si) => {
    let p = pos + 1;
    s.blocks.forEach((b, k) => {
      y += b.mt;
      const top = y;
      const bottom = top + heightOf(b);
      y = bottom + b.mb;
      const size = sizeOf(b);
      units.push({
        key: b.key,
        ctx: [],
        pos: p,
        breakPos: p,
        end: p + size,
        kind: b.kind === "row" ? "row" : b.kind === "cols" ? "section" : "block",
        top,
        bottom,
        keepNext: b.kind === "h" || b.keepNext,
        pageBreak: b.kind === "pb",
        colBreak: false,
        breakBefore: b.breakBefore,
        keepLines: b.keepLines,
        text: (b.kind === "p" || b.kind === "h") && b.lines.length > 0,
        notes: b.notes,
        sec: si,
        secStart: k === 0,
        hh: b.kind === "row" ? b.hh : 0,
      });
      blocks.push(b);
      p += size;
    });
    pos = p + 1;
    y += 8; // bölüm arası
  });
  return { units, blocks };
}

/** Taklit ölçüm: değişmeyen blok için sonuç yalnız bloğa ve doğal konumuna bağlıdır. */
function measureOf(units: FlowUnit[], blocks: Block[], geoms: SecGeom[], calls: { coords: number }): FlowMeasure {
  return {
    outerTop: (i) => units[i].top - blocks[i].mt,
    coords: (i, pos, side) => {
      calls.coords++;
      const u = units[i];
      const b = blocks[i];
      const c = pos - (u.pos + 1);
      if (c < 0 || c > b.lines.length * b.cpl) return null;
      let line = Math.min(b.lines.length - 1, Math.floor(c / b.cpl));
      // Satır başında geriye doğru: önceki satırın sonu.
      if (side < 0 && c > 0 && c % b.cpl === 0) line = Math.min(line, c / b.cpl - 1);
      let top = u.top;
      for (let k = 0; k < line; k++) top += b.lines[k];
      // Karakter kutusu satır kutusundan dar (yarım aralık 2 px).
      return { top: top + 2, bottom: top + b.lines[line] - 2 };
    },
    lineHeight: (i) => (blocks[i].normalLh ? NaN : blocks[i].lines[0] ?? NaN),
    noteH: (key) => 16 + (key.length % 3) * 8,
    columns: (i, c) => columnsMock(units[i], blocks[i], geoms[units[i].sec], c),
  };
}

/** Çok sütunlu bölüm taklidi: tek sayfaya sığarsa sütunlar dengelenir, sığmazsa yarım satırlık parçalarla sayfalara bölünür. */
function columnsMock(u: FlowUnit, b: Block, g: SecGeom, c: ColumnCtx): ColumnResult {
  const H = u.bottom - u.top;
  const breaks: PageBreakMark[] = [];
  const starts: { page: number; pos: number }[] = [];
  let p = c.page;
  if (c.top + H <= c.contentEnd(p) + 0.5) {
    if (c.endsPage) return { breaks, bottom: c.contentEnd(p), page: p, style: { pos: u.pos, height: Math.max(0, c.contentEnd(p) - c.top) }, starts };
    return { breaks, bottom: c.top + H, page: p, starts };
  }
  // Tek sütundaki toplam yükseklik (dar ölçüm) sütun sayısıyla çarpılır; her sayfa sütun sayısı kadar doldurulur.
  const narrow = H * g.cols;
  let setTop = c.top;
  let A = 0;
  let k = 0;
  while (narrow - A > (c.contentEnd(p) - setTop) * g.cols + 0.5 && k < b.lines.length) {
    const cap = (c.contentEnd(p) - setTop) * g.cols;
    // Bölme satır sınırında (satırlar tek sütunda üst üste).
    let used = 0;
    let line = 0;
    while (line < b.lines.length && used + b.lines[line] * g.cols <= A + cap) used += b.lines[line++] * g.cols;
    if (used <= A) break;
    const np = c.nextPage();
    const target = c.pages[np].top + c.pages[np].mt;
    const at = u.pos + 1 + line * 7;
    breaks.push({ pos: at, kind: "span", h: Math.max(1, target - (setTop + (used - A) / g.cols)) });
    starts.push({ page: np, pos: at });
    p = np;
    setTop = target;
    A = used;
    k++;
  }
  if (c.endsPage) return { breaks, bottom: c.contentEnd(p), page: p, style: { pos: u.pos, height: Math.max(0, c.contentEnd(p) - c.top) }, starts };
  return { breaks, bottom: setTop + (narrow - A) / g.cols, page: p, starts };
}

/** `touched`: kaydı aynı kalsa da değişmiş bildirilen birimler (ör. DOM'u belge dışından değişen). */
function run(doc: Doc, prev?: FlowCache | null, touched: number[] = []): FlowResult & { coords: number; n: number } {
  const geoms = geomsOf(doc);
  const { units, blocks } = unitsOf(doc);
  const calls = { coords: 0 };
  const changed = new Uint8Array(units.length);
  for (const k of touched) if (units.length) changed[k % units.length] = 1;
  const ctx: FlowCtx = { units, geoms, bodyW: Math.max(...geoms.map((g) => g.w)), gap: 16, Y: (y) => y, screenY: (y) => y, scale: 1, slotH: 0, m: measureOf(units, blocks, geoms, calls), changed };
  return { ...layoutFlow(ctx, prev), coords: calls.coords, n: units.length };
}

const out = (r: FlowResult) => ({ breaks: r.breaks, styles: r.styles, info: r.info });
/**
 * Beklenen sonuçtaki sayılar 1e-6 duyarlılıkla karşılaştırılır: artımlı çalışma önceki sonucu eski kaydırmadan
 * (near() ile kabul edilen eşleşme) sürdürdüğü için toplama sırası son basamakta fark yaratabilir (ör. 340.3000000000011
 * ile 340.3000000000002). Konumlar ve sayılar tam sayı olduğundan yine birebir karşılaştırılmış olur.
 */
const yakin = (v: unknown): unknown =>
  typeof v === "number" ? expect.closeTo(v, 6) : Array.isArray(v) ? v.map(yakin) : v && typeof v === "object" ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, yakin(x)])) : v;

/* ------------------------------------------------------------------ */
/* Belge üreticileri                                                     */
/* ------------------------------------------------------------------ */

let seq = 0;
const para = (lines: number, o: Partial<Block> = {}): Block => ({ key: {}, kind: "p", lines: Array(lines).fill(20), cpl: 60, mt: 0, mb: 12, keepNext: false, keepLines: false, breakBefore: false, normalLh: false, notes: [], hh: 0, ...o });
const heading = (o: Partial<Block> = {}): Block => para(1, { kind: "h", lines: [32], mt: 24, mb: 8, ...o });
const row = (h: number, hh = 0): Block => para(0, { kind: "row", lines: [h], mb: 0, hh });
const pageBreak = (): Block => para(0, { kind: "pb", lines: [20], mb: 0 });

/** Uzun metin belgesi: başlıklar, 3-8 satırlık paragraflar ve ara sıra tablolar. */
function textDoc(chapters: number): Doc {
  const blocks: Block[] = [];
  for (let c = 0; c < chapters; c++) {
    blocks.push(heading());
    for (let k = 0; k < 12; k++) blocks.push(para(3 + ((c * 7 + k * 3) % 6)));
    if (c % 2 === 1) for (let r = 0; r < 6; r++) blocks.push(row(26 + (r % 3) * 8, r ? 26 : 0));
  }
  return { secs: [{ kind: "nextPage", land: false, cols: 1, blocks }] };
}

const replaceBlock = (doc: Doc, s: number, i: number, b: Block): Doc => ({ secs: doc.secs.map((x, k) => (k === s ? { ...x, blocks: x.blocks.map((y, j) => (j === i ? b : y)) } : x)) });

/* ------------------------------------------------------------------ */
/* Belirli durumlar                                                      */
/* ------------------------------------------------------------------ */

describe("sayfalama düzen birimi", () => {
  it("kısa belge tek sayfa; uzun paragraf satır sınırından bölünür", () => {
    const one = run({ secs: [{ kind: "nextPage", land: false, cols: 1, blocks: [para(3), para(4)] }] });
    expect(one.info.pageCount).toBe(1);
    expect(one.breaks).toEqual([]);
    const long = run({ secs: [{ kind: "nextPage", land: false, cols: 1, blocks: [para(100)] }] });
    expect(long.info.pageCount).toBe(3);
    expect(long.breaks.every((b) => b.kind === "line" && b.h > 0)).toBe(true);
    // Her sayfa 43 satır alır (864 px içerik / 20 px), ikinci sayfa 44. satırdan başlar.
    expect(long.breaks[0].pos).toBe(1 + 1 + 43 * 60);
  });

  it("sayfa sonu, başlığın sonrakiyle tutulması, tablo satırı ve yinelenen başlık", () => {
    const doc: Doc = { secs: [{ kind: "nextPage", land: false, cols: 1, blocks: [para(38), heading(), para(5), pageBreak(), para(2), ...Array.from({ length: 40 }, (_, k) => row(30, k ? 24 : 0))] }] };
    const r = run(doc);
    // Başlık sayfa sonuna sığsa da paragrafıyla yeni sayfaya geçer.
    const hPos = r.cache.units[1].pos;
    expect(r.breaks.some((b) => b.pos === hPos && b.kind === "block")).toBe(true);
    expect(r.breaks.some((b) => b.kind === "row")).toBe(true);
    expect(r.breaks.some((b) => b.kind === "rowhead" && b.h === 24)).toBe(true);
    expect(r.info.pageCount).toBeGreaterThanOrEqual(4);
  });

  it("dipnot sayfanın içerik alanını daraltır ve sayfasında listelenir", () => {
    const notes = [{ id: "n1", text: "dipnot metni" }];
    const r = run({ secs: [{ kind: "nextPage", land: false, cols: 1, blocks: [para(30), para(2, { notes }), para(30)] }] });
    expect(r.info.footnotes[0]).toEqual(["n1"]);
  });

  it("bölümler: tek sayfa bölüm sonu boş sayfa ekler, yatay bölüm yeni sayfada başlar", () => {
    const r = run({ secs: [
      { kind: "nextPage", land: false, cols: 1, blocks: [para(5)] },
      { kind: "oddPage", land: false, cols: 1, blocks: [para(5)] },
      { kind: "continuous", land: true, cols: 1, blocks: [para(5)] },
    ] });
    expect(r.info.pages.map((p) => [p.section, !!p.blank, p.w])).toEqual([[0, false, 816], [1, true, 816], [1, false, 816], [2, false, 1056]]);
  });

  it("artımlı: son sayfada yazmak yalnız son sayfayı hesaplar", () => {
    const doc = textDoc(12);
    const full = run(doc);
    const blocks = doc.secs[0].blocks;
    const last = blocks.length - 1;
    // Satır sayısı değişmeden son paragrafa yazma: yalnız düğüm kimliği değişir.
    const typed = replaceBlock(doc, 0, last, { ...blocks[last], key: {} });
    const inc = run(typed, full.cache);
    expect(out(inc)).toEqual(yakin(out(run(typed))));
    expect(inc.stats.incremental).toBe(true);
    const lastPage = full.info.pageCount - 1;
    expect(inc.cache.unitPage[inc.stats.from]).toBeGreaterThanOrEqual(lastPage - 1);
  });

  it("artımlı: baştaki paragrafa satır eklenince değişen sayfalardan sonra önceki sonuç kullanılır", () => {
    const doc = textDoc(12);
    const full = run(doc);
    const blocks = doc.secs[0].blocks;
    const grown = replaceBlock(doc, 0, 2, { ...blocks[2], key: {}, lines: [...blocks[2].lines, 20] });
    const ref = run(grown);
    const inc = run(grown, full.cache);
    expect(out(inc)).toEqual(yakin(out(ref)));
    expect(inc.stats.from).toBe(0);
    expect(inc.stats.to).toBeLessThan(inc.n);
    // Kesmelerin kayması bir yerde durur: hesaplanan bölüm tüm belgeden çok daha kısa, ölçüm de az.
    expect(inc.coords).toBeLessThan(ref.coords / 2);
  });

  it("artımlı: aynı yükseklikte başka dipnot önceki sayfanın dipnot listesiyle birleştirilmez", () => {
    const doc = textDoc(8);
    const blocks = doc.secs[0].blocks;
    const withNote = replaceBlock(doc, 0, 3, { ...blocks[3], key: {}, notes: [{ id: "a", text: "kısa" }] });
    const before = run(withNote);
    // Aynı uzunlukta (aynı yükseklikte) başka kimlikli dipnot: dipnot alanı aynı, liste farklı.
    const swapped = replaceBlock(withNote, 0, 3, { ...blocks[3], key: {}, notes: [{ id: "b", text: "abcd" }] });
    const inc = run(swapped, before.cache);
    expect(out(inc)).toEqual(yakin(out(run(swapped))));
    expect(inc.info.footnotes.flat()).toEqual(["b"]);
    expect(inc.stats.to).toBeLessThan(inc.n);
  });

  it("ortam (sayfa yapısı) değişince tam hesap yapılır", () => {
    const doc = textDoc(4);
    const full = run(doc);
    const land: Doc = { secs: doc.secs.map((s) => ({ ...s, land: true })) };
    const inc = run(land, full.cache);
    expect(inc.stats.incremental).toBe(false);
    expect(out(inc)).toEqual(yakin(out(run(land))));
  });
});

/* ------------------------------------------------------------------ */
/* Özellik: rastgele düzenleme dizileri                                  */
/* ------------------------------------------------------------------ */

const blockArb: fc.Arbitrary<Block> = fc
  .record({
    kind: fc.constantFrom<Kind>("p", "p", "p", "p", "h", "row", "row", "pb"),
    n: fc.integer({ min: 1, max: 14 }),
    lh: fc.constantFrom(20, 20, 18.5, 24.25, 32),
    cpl: fc.constantFrom(40, 60),
    mt: fc.constantFrom(0, 0, 6, 12.5),
    mb: fc.constantFrom(0, 8, 12),
    keepNext: fc.constantFrom(false, false, false, false, true),
    keepLines: fc.boolean(),
    breakBefore: fc.constantFrom(false, false, false, false, false, false, false, true),
    normalLh: fc.constantFrom(false, false, false, true),
    note: fc.constantFrom(null, null, null, null, null, null, "kısa", "daha uzun dipnot"),
    hh: fc.constantFrom(0, 0, 22),
    big: fc.constantFrom(false, false, false, false, false, false, false, false, false, true),
  })
  .map((r) => {
    const n = r.big ? 60 + r.n * 4 : r.n;
    if (r.kind === "row") return row(18 + r.n * 6, r.hh);
    if (r.kind === "pb") return pageBreak();
    if (r.kind === "h") return heading({ lines: [r.lh + 12], keepLines: r.keepLines });
    return para(n, { lines: Array.from({ length: n }, (_, k) => (k % 5 === 4 ? r.lh + 4 : r.lh)), cpl: r.cpl, mt: r.mt, mb: r.mb, keepNext: r.keepNext, keepLines: r.keepLines && n < 10, breakBefore: r.breakBefore, normalLh: r.normalLh, notes: r.note ? [{ id: `n${seq++}`, text: r.note }] : [] });
  });

type Edit =
  | { op: "ins"; at: number; b: Block }
  | { op: "del"; at: number }
  | { op: "rep"; at: number; b: Block }
  | { op: "type"; at: number }
  | { op: "line"; at: number; add: boolean }
  | { op: "flag"; at: number; f: "keepNext" | "keepLines" | "breakBefore" }
  | { op: "note"; at: number }
  | { op: "multi"; at: number; n: number };

const editArb: fc.Arbitrary<Edit> = fc.oneof(
  { weight: 3, arbitrary: fc.record({ op: fc.constant("ins" as const), at: fc.nat(), b: blockArb }) },
  { weight: 2, arbitrary: fc.record({ op: fc.constant("del" as const), at: fc.nat() }) },
  { weight: 2, arbitrary: fc.record({ op: fc.constant("rep" as const), at: fc.nat(), b: blockArb }) },
  { weight: 5, arbitrary: fc.record({ op: fc.constant("type" as const), at: fc.nat() }) },
  { weight: 5, arbitrary: fc.record({ op: fc.constant("line" as const), at: fc.nat(), add: fc.boolean() }) },
  { weight: 2, arbitrary: fc.record({ op: fc.constant("flag" as const), at: fc.nat(), f: fc.constantFrom("keepNext" as const, "keepLines" as const, "breakBefore" as const) }) },
  { weight: 1, arbitrary: fc.record({ op: fc.constant("note" as const), at: fc.nat() }) },
  { weight: 1, arbitrary: fc.record({ op: fc.constant("multi" as const), at: fc.nat(), n: fc.integer({ min: 2, max: 6 }) }) },
);

function applyEdit(doc: Doc, e: Edit): Doc {
  const total = doc.secs.reduce((a, s) => a + s.blocks.length, 0);
  if (total === 0 && e.op !== "ins") return doc;
  // Düz sıra → (bölüm, blok)
  const locate = (at: number, ins = false) => {
    let k = at % (total + (ins ? 1 : 0) || 1);
    for (let s = 0; s < doc.secs.length; s++) {
      const len = doc.secs[s].blocks.length + (ins && s === doc.secs.length - 1 ? 1 : 0);
      if (k < len || (ins && k <= doc.secs[s].blocks.length)) return { s, i: Math.min(k, doc.secs[s].blocks.length) };
      k -= doc.secs[s].blocks.length;
    }
    return { s: doc.secs.length - 1, i: doc.secs[doc.secs.length - 1].blocks.length };
  };
  const edit = (s: number, f: (bl: Block[]) => Block[]): Doc => ({ secs: doc.secs.map((x, k) => (k === s ? { ...x, blocks: f(x.blocks) } : x)) });
  switch (e.op) {
    case "ins": {
      const { s, i } = locate(e.at, true);
      return edit(s, (bl) => [...bl.slice(0, i), { ...e.b, key: {} }, ...bl.slice(i)]);
    }
    case "del": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) => (bl.length > 1 ? bl.filter((_, k) => k !== i) : bl));
    }
    case "rep": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) => bl.map((b, k) => (k === i ? { ...e.b, key: {} } : b)));
    }
    case "type": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) => bl.map((b, k) => (k === i ? { ...b, key: {} } : b)));
    }
    case "line": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) =>
        bl.map((b, k) => {
          if (k !== i || (b.kind !== "p" && b.kind !== "cols")) return b;
          const lines = e.add ? [...b.lines, b.lines[b.lines.length - 1] ?? 20] : b.lines.slice(0, Math.max(1, b.lines.length - 1));
          return { ...b, key: {}, lines };
        }),
      );
    }
    case "flag": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) => bl.map((b, k) => (k === i && b.kind === "p" ? { ...b, key: {}, [e.f]: !b[e.f] } : b)));
    }
    case "note": {
      const { s, i } = locate(e.at);
      return edit(s, (bl) => bl.map((b, k) => (k === i && b.kind === "p" ? { ...b, key: {}, notes: b.notes.length ? [] : [{ id: `n${seq++}`, text: "yeni dipnot" }] } : b)));
    }
    case "multi": {
      // Birbirinden uzak birkaç bloğa aynı anda yazma (ör. Tümünü Değiştir).
      let d = doc;
      for (let k = 0; k < e.n; k++) d = applyEdit(d, { op: "line", at: e.at + k * 17, add: k % 2 === 0 });
      return d;
    }
  }
}

const sectionArb = fc.record({
  kind: fc.constantFrom<SectionKind>("nextPage", "continuous", "evenPage", "oddPage"),
  land: fc.constantFrom(false, false, true),
  cols: fc.constantFrom(1, 1, 1, 2),
  blocks: fc.array(blockArb, { minLength: 1, maxLength: 70 }),
});

const docArb: fc.Arbitrary<Doc> = fc.array(sectionArb, { minLength: 1, maxLength: 3 }).map((secs) => ({
  // Çok sütunlu bölüm tek birimdir (taklit: tek "cols" blok).
  secs: secs.map((s) => (s.cols > 1 ? { ...s, blocks: [para(0, { kind: "cols", lines: s.blocks.slice(0, 30).map(() => 20), mb: 0 })] } : s)),
}));

describe("artımlı sayfalama = tam sayfalama (rastgele düzenleme dizileri)", () => {
  it("her adımda kesmeler, sayfalar, sayfa başları ve dipnotlar birebir aynı", () => {
    const seen = { steps: 0, resumed: 0, spliced: 0 };
    fc.assert(
      fc.property(docArb, fc.array(editArb, { minLength: 1, maxLength: 12 }), fc.array(fc.nat(), { maxLength: 3 }), (doc0, edits, touched) => {
        let doc = doc0;
        let cache = run(doc).cache;
        for (const e of edits) {
          doc = applyEdit(doc, e);
          const inc = run(doc, cache, touched);
          const full = run(doc);
          expect(out(inc)).toEqual(yakin(out(full)));
          seen.steps++;
          if (inc.stats.from > 0) seen.resumed++;
          if (inc.stats.to < inc.n) seen.spliced++;
          cache = inc.cache;
        }
      }),
      { numRuns: 300 },
    );
    // Denetim anlamlı olsun: adımların önemli bir kısmı ortadan başlar ve önceki sonuçla birleşir.
    expect(seen.resumed).toBeGreaterThan(seen.steps / 5);
    expect(seen.spliced).toBeGreaterThan(seen.steps / 5);
  });

  it("uzun metin belgesinde yazma dizisi: her adım tam hesapla aynı ve çoğu adım küçük bir bölümü hesaplar", () => {
    const seen = { steps: 0, small: 0 };
    fc.assert(
      fc.property(fc.integer({ min: 6, max: 16 }), fc.array(editArb, { minLength: 4, maxLength: 20 }), (chapters, edits) => {
        let doc = textDoc(chapters);
        let cache = run(doc).cache;
        for (const e of edits) {
          doc = applyEdit(doc, e);
          const inc = run(doc, cache);
          expect(out(inc)).toEqual(yakin(out(run(doc))));
          seen.steps++;
          if (inc.stats.to - inc.stats.from < inc.n / 4) seen.small++;
          cache = inc.cache;
        }
      }),
      { numRuns: 150 },
    );
    expect(seen.small).toBeGreaterThan(seen.steps / 2);
  });
});
