/**
 * Asıl Slayt Görünümü: asıl slayt ve düzenler, düzenleyicinin tüm araçlarıyla
 * düzenlenebilsin diye geçici bir "görünüm sunusuna" dönüştürülür (ilk slayt
 * asıl slayt, ardından düzenler). Görünüm kapatılınca değişiklikler asıl
 * slayta/düzenlere geri yazılır: yer tutucu örnek metnindeki biçimler düzey
 * stillerine (lstStyle) katlanır, konumu düzenden gelen slayt yer tutucuları
 * yeni konuma taşınır.
 */
import { uid } from "./factory";
import { layoutByType, relayoutSlide } from "./presentation";
import { findPlaceholder } from "./textStyle";
import type { Layout, LevelStyle, Paragraph, Placeholder, Presentation, Shape, Slide, SpShape, TextBody } from "./types";

export const MASTER_SLIDE_ID = "__master";
const PL = "__pl:";

const LEVEL_NAMES = ["Asıl metin stillerini düzenlemek için tıklayın", "İkinci düzey", "Üçüncü düzey", "Dördüncü düzey", "Beşinci düzey"];

function today(): string {
  return new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(new Date());
}

/** Yer tutucunun asıl görünümdeki örnek paragrafları. */
export function samplePlaceholderText(ph: Placeholder, isMaster: boolean, prompt?: string, footerText?: string): Paragraph[] {
  const p = (text: string, level = 0): Paragraph => ({ runs: [{ text, props: {} }], level, props: {} });
  if (prompt && !isMaster && ph.type !== "dt" && ph.type !== "ftr" && ph.type !== "sldNum") return prompt.split("\n").map((t) => p(t));
  switch (ph.type) {
    case "title":
    case "ctrTitle":
      return [p("Asıl başlık stilini düzenlemek için tıklayın")];
    case "subTitle":
      return [p("Asıl alt başlık stilini düzenlemek için tıklayın")];
    case "body":
    case "obj":
      return LEVEL_NAMES.map((t, i) => p(t, i));
    case "pic":
      return [p("Resim eklemek için simgeye tıklayın")];
    case "tbl":
      return [p("Tablo eklemek için simgeye tıklayın")];
    case "dt":
      return [p(today())];
    case "ftr":
      return [p(footerText || "Alt Bilgi")];
    case "sldNum":
      return [p("‹#›")];
  }
}

function plain(b: TextBody | undefined): string {
  return (b?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

function withSample(s: Shape, isMaster: boolean): Shape {
  if (s.kind !== "sp" || !s.ph) return s;
  const text = s.text ?? { paragraphs: [] };
  const footerText = s.ph.type === "ftr" ? plain(s.text) : undefined;
  let paragraphs = samplePlaceholderText(s.ph, isMaster, text.prompt, footerText);
  // Madde işaretsiz (başlık benzeri) metin yer tutucularında yalnız ilk düzey gösterilir.
  if (!isMaster && (s.ph.type === "body" || s.ph.type === "obj") && text.levels?.[0]?.para?.bullet?.type === "none") paragraphs = paragraphs.slice(0, 1);
  return { ...s, text: { ...text, paragraphs } };
}

/** Sunuyu asıl görünüm sunusuna çevirir. */
export function toMasterView(p: Presentation): Presentation {
  const masterShapes = p.master.shapes.map((s) => withSample(s, true));
  const masterSlide: Slide = { id: MASTER_SLIDE_ID, layoutId: `${PL}${MASTER_SLIDE_ID}`, shapes: masterShapes, background: p.master.background, notes: "", animations: [] };
  const layoutSlides: Slide[] = p.layouts.map((l) => ({
    id: l.id,
    layoutId: `${PL}${l.id}`,
    shapes: l.shapes.map((s) => withSample(s, false)),
    background: l.background,
    hideMasterShapes: l.showMasterShapes ? undefined : true,
    notes: "",
    animations: [],
  }));
  const layouts: Layout[] = [
    { id: `${PL}${MASTER_SLIDE_ID}`, type: "cust", name: "Asıl Slayt", shapes: [], showMasterShapes: false },
    ...p.layouts.map((l) => ({ id: `${PL}${l.id}`, type: l.type, name: l.name, shapes: [], showMasterShapes: true })),
  ];
  return { ...p, master: { ...p.master, shapes: masterShapes }, layouts, slides: [masterSlide, ...layoutSlides] };
}

/** Görünüm sunusunda asıl slayt (ilk slayt) değişince `master` alanını eşitler. */
export function syncMasterView(mv: Presentation): Presentation {
  const ms = mv.slides.find((s) => s.id === MASTER_SLIDE_ID);
  if (!ms) return mv;
  const bg = ms.background ?? mv.master.background;
  if (syncedFrom.get(mv.master.shapes) === ms.shapes && bg === mv.master.background) return mv;
  // Asıl slayttaki örnek metne verilen biçimler (ör. kalın başlık) düzey stillerine katlanır;
  // böylece düzenler ve onların küçük resimleri değişikliği anında yansıtır.
  const shapes = ms.shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph || !s.text) return s;
    const f = fold(s, true);
    return f.kind === "sp" && f.text ? { ...s, text: { ...s.text, levels: f.text.levels } } : s;
  });
  syncedFrom.set(shapes, ms.shapes);
  return { ...mv, master: { ...mv.master, shapes, background: bg } };
}

/** Eşitlenmiş asıl şekil dizisi → kaynağı olan asıl slayt şekilleri. */
const syncedFrom = new WeakMap<Shape[], Shape[]>();

export function isLayoutSlide(mv: Presentation, slideId: string | undefined): boolean {
  return !!slideId && slideId !== MASTER_SLIDE_ID && mv.layouts.some((l) => l.id === `${PL}${slideId}`);
}

/** Görünüm sunusundaki düzen adı/türü. */
export function pseudoLayoutOf(mv: Presentation, slideId: string): Layout | undefined {
  return mv.layouts.find((l) => l.id === `${PL}${slideId}`);
}

function strip<T extends object>(o: T | undefined): Partial<T> {
  const r: Partial<T> = {};
  if (!o) return r;
  for (const [k, v] of Object.entries(o)) if (v !== undefined) (r as Record<string, unknown>)[k] = v;
  return r;
}

/** Örnek metindeki paragraf/çalıştırma biçimlerini düzey stillerine katlar ve metni boşaltır. */
function fold(s: Shape, isMaster: boolean): Shape {
  if (s.kind !== "sp" || !s.ph || !s.text) return s;
  const body = s.text;
  const levels: Partial<LevelStyle>[] = (body.levels ?? []).map((l) => ({ para: { ...(l?.para ?? {}) }, run: { ...(l?.run ?? {}) } }));
  const seen = new Set<number>();
  for (const p of body.paragraphs) {
    if (seen.has(p.level)) continue;
    seen.add(p.level);
    const run = p.runs.find((r) => r.text)?.props ?? p.endProps;
    const pp = strip(p.props);
    const rp = strip(run);
    delete rp.link;
    delete rp.field;
    if (!Object.keys(pp).length && !Object.keys(rp).length) continue;
    while (levels.length <= p.level) levels.push({});
    const cur = levels[p.level] ?? {};
    levels[p.level] = { para: { ...(cur.para ?? {}), ...pp }, run: { ...(cur.run ?? {}), ...rp } };
  }
  const type = s.ph.type;
  const footer = type === "dt" || type === "ftr" || type === "sldNum";
  const text = plain(body);
  const full = samplePlaceholderText(s.ph, isMaster, body.prompt);
  const expectedSet = new Set([plain({ paragraphs: full }), plain({ paragraphs: full.slice(0, 1) })]);
  let prompt = body.prompt;
  let paragraphs: Paragraph[] = [{ runs: [], level: 0, props: {} }];
  if (footer) {
    // Alt bilgiye yazılan metin korunur; örnek metin boşaltılır.
    const sample = plain({ paragraphs: samplePlaceholderText(s.ph, isMaster) });
    if (type === "ftr" && text && text !== sample) paragraphs = [{ runs: [{ text, props: {} }], level: 0, props: {} }];
  } else if (!isMaster && text.trim() && !expectedSet.has(text)) {
    // Düzende değiştirilen örnek metin özel istem metni olur.
    prompt = text;
  }
  const levelsOut = levels.some((l) => l && (Object.keys(l.para ?? {}).length || Object.keys(l.run ?? {}).length)) ? levels : undefined;
  return { ...s, text: { ...body, paragraphs, prompt, levels: levelsOut, fontScale: undefined, lnSpcReduction: undefined } };
}

function sameGeom(a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }) {
  return Math.abs(a.x - b.x) < 0.6 && Math.abs(a.y - b.y) < 0.6 && Math.abs(a.w - b.w) < 0.6 && Math.abs(a.h - b.h) < 0.6;
}

/** Konumu eski kaynaktan gelen yer tutucuları yeni kaynağın konumuna taşır. */
function followGeometry(shapes: Shape[], oldSrc: Shape[], newSrc: Shape[], typeOnly: boolean): Shape[] {
  let changed = false;
  const out = shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph) return s;
    const key = typeOnly ? { type: s.ph.type } : s.ph;
    const o = findPlaceholder(oldSrc, key, !typeOnly);
    const n = findPlaceholder(newSrc, key, !typeOnly);
    if (!o || !n || !sameGeom(o, s) || sameGeom(o, n)) return s;
    changed = true;
    return { ...s, x: n.x, y: n.y, w: n.w, h: n.h };
  });
  return changed ? out : shapes;
}

/** Asıl görünüm sunusundaki değişiklikleri asıl sunuya uygular. */
export function fromMasterView(base: Presentation, mv: Presentation): Presentation {
  const ms = mv.slides.find((s) => s.id === MASTER_SLIDE_ID);
  const masterShapes = (ms?.shapes ?? base.master.shapes).map((s) => fold(s, true));
  const master = { ...base.master, background: ms?.background ?? base.master.background, shapes: masterShapes };
  const layouts: Layout[] = mv.slides
    .filter((s) => s.id !== MASTER_SLIDE_ID)
    .map((ls) => {
      const pl = pseudoLayoutOf(mv, ls.id);
      const orig = base.layouts.find((l) => l.id === ls.id);
      let shapes = ls.shapes.map((s) => fold(s, false));
      // Asıl slaytta taşınan yer tutucular, düzende aynı konumdaysa onu izler.
      shapes = followGeometry(shapes, base.master.shapes, masterShapes, true);
      return {
        id: ls.id,
        type: pl?.type ?? orig?.type ?? "cust",
        name: pl?.name ?? orig?.name ?? "Özel Düzen",
        shapes,
        background: ls.background,
        showMasterShapes: !ls.hideMasterShapes,
      };
    });
  const next: Presentation = { ...base, theme: mv.theme, master, layouts, masterEdited: true };
  const fallback = layouts.find((l) => l.type === "obj") ?? layouts[0];
  next.slides = base.slides.map((sl) => {
    const oldL = base.layouts.find((l) => l.id === sl.layoutId);
    const newL = layouts.find((l) => l.id === sl.layoutId);
    if (!newL) return fallback ? relayoutSlide(next, sl, layoutByType(next, oldL?.type ?? "obj"), oldL) : sl;
    if (!oldL) return sl;
    const shapes = followGeometry(sl.shapes, oldL.shapes, newL.shapes, false);
    return shapes === sl.shapes ? sl : { ...sl, shapes };
  });
  return next;
}

/* ------------------------------------------------------------------ */
/* Görünüm sunusunda düzen işlemleri                                     */
/* ------------------------------------------------------------------ */

function phShape(type: Placeholder["type"], idx: number | undefined, name: string, box: { x: number; y: number; w: number; h: number }): SpShape {
  return {
    kind: "sp",
    id: uid(),
    name,
    ...box,
    rot: 0,
    geom: "rect",
    fill: { type: "none" },
    line: { fill: { type: "none" }, width: 1, dash: "solid" },
    ph: idx === undefined ? { type } : { type, idx },
    text: {
      paragraphs: samplePlaceholderText({ type }, false),
      // Resim/tablo yer tutucularında madde işareti yoktur.
      levels: type === "pic" || type === "tbl" ? [{ para: { bullet: { type: "none" }, marL: 0, indent: 0, align: "ctr" }, run: { size: 18 } }] : undefined,
      anchor: type === "pic" || type === "tbl" ? "ctr" : undefined,
    },
  };
}

/** Asıl slayttaki yer tutucunun kopyası (yeni düzen ve "Başlık"/"Alt Bilgiler" onay kutuları için). */
export function placeholderFromMaster(mv: Presentation, type: Placeholder["type"]): SpShape | undefined {
  const m = findPlaceholder(mv.master.shapes, { type });
  if (!m || m.kind !== "sp") return undefined;
  const s = { ...(m as SpShape), id: uid(), text: { ...(m as SpShape).text!, levels: undefined, prompt: undefined, paragraphs: samplePlaceholderText({ type }, false) } };
  if (type === "title") s.name = "Başlık 1";
  return s;
}

/** Görünüm sunusuna yeni (özel) düzen ekler; eklenen slayt dizinini döndürür. */
export function insertLayout(mv: Presentation, afterIndex: number, name = "Özel Düzen"): { pres: Presentation; index: number } {
  const id = uid("L");
  const shapes: Shape[] = [];
  const title = placeholderFromMaster(mv, "title");
  if (title) shapes.push(title);
  for (const t of ["dt", "ftr", "sldNum"] as const) {
    const f = placeholderFromMaster(mv, t);
    if (f) shapes.push(f);
  }
  const slide: Slide = { id, layoutId: `${PL}${id}`, shapes, notes: "", animations: [] };
  const at = Math.max(1, Math.min(mv.slides.length, afterIndex + 1));
  const slides = [...mv.slides.slice(0, at), slide, ...mv.slides.slice(at)];
  const layouts = [...mv.layouts, { id: `${PL}${id}`, type: "cust" as const, name, shapes: [], showMasterShapes: true }];
  return { pres: { ...mv, slides, layouts }, index: at };
}

export function duplicateLayout(mv: Presentation, index: number): { pres: Presentation; index: number } | null {
  const src = mv.slides[index];
  if (!src || src.id === MASTER_SLIDE_ID) return null;
  const pl = pseudoLayoutOf(mv, src.id);
  const id = uid("L");
  const copy: Slide = { ...structuredClone(src), id, layoutId: `${PL}${id}`, shapes: src.shapes.map((s) => ({ ...structuredClone(s), id: uid() })) };
  const slides = [...mv.slides.slice(0, index + 1), copy, ...mv.slides.slice(index + 1)];
  const layouts = [...mv.layouts, { id: `${PL}${id}`, type: "cust" as const, name: `1_${pl?.name ?? "Özel Düzen"}`, shapes: [], showMasterShapes: true }];
  return { pres: { ...mv, slides, layouts }, index: index + 1 };
}

export function renameLayout(mv: Presentation, slideId: string, name: string): Presentation {
  return { ...mv, layouts: mv.layouts.map((l) => (l.id === `${PL}${slideId}` ? { ...l, name } : l)) };
}

/** Düzeni kullanan sunu slaytlarının sayısı. */
export function layoutUsage(base: Presentation, layoutId: string): number {
  return base.slides.filter((s) => s.layoutId === layoutId).length;
}

export function deleteLayout(mv: Presentation, slideId: string): Presentation {
  if (slideId === MASTER_SLIDE_ID) return mv;
  return { ...mv, slides: mv.slides.filter((s) => s.id !== slideId), layouts: mv.layouts.filter((l) => l.id !== `${PL}${slideId}`) };
}

/** Düzene başlık / alt bilgi yer tutucularını ekler ya da kaldırır. */
export function toggleLayoutPlaceholders(mv: Presentation, slideId: string, kind: "title" | "footers", on: boolean): Presentation {
  const types: Placeholder["type"][] = kind === "title" ? ["title", "ctrTitle"] : ["dt", "ftr", "sldNum"];
  return {
    ...mv,
    slides: mv.slides.map((s) => {
      if (s.id !== slideId || s.id === MASTER_SLIDE_ID) return s;
      const has = (t: Placeholder["type"]) => s.shapes.some((x) => x.kind === "sp" && x.ph?.type === t);
      if (!on) return { ...s, shapes: s.shapes.filter((x) => !(x.kind === "sp" && x.ph && types.includes(x.ph.type))) };
      const add: Shape[] = [];
      for (const t of kind === "title" ? (["title"] as const) : (["dt", "ftr", "sldNum"] as const)) {
        if (has(t) || (t === "title" && has("ctrTitle"))) continue;
        const p = placeholderFromMaster(mv, t);
        if (p) add.push(p);
      }
      return add.length ? { ...s, shapes: kind === "title" ? [...add, ...s.shapes] : [...s.shapes, ...add] } : s;
    }),
  };
}

const PH_NAMES: Record<string, string> = { obj: "İçerik Yer Tutucusu", body: "Metin Yer Tutucusu", pic: "Resim Yer Tutucusu", tbl: "Tablo Yer Tutucusu" };

/** Düzene yeni yer tutucu ekler (Yer Tutucu Ekle). */
export function addLayoutPlaceholder(mv: Presentation, slideId: string, type: "obj" | "body" | "pic" | "tbl"): { pres: Presentation; shapeId: string } | null {
  const s = mv.slides.find((x) => x.id === slideId);
  if (!s || slideId === MASTER_SLIDE_ID) return null;
  const idx = Math.max(9, ...s.shapes.map((x) => (x.kind === "sp" && x.ph?.idx !== undefined ? x.ph.idx : 0))) + 1;
  const n = s.shapes.filter((x) => x.kind === "sp" && x.ph).length + 1;
  const box = { x: mv.width * 0.3, y: mv.height * 0.3, w: mv.width * 0.4, h: mv.height * 0.4 };
  const ph = phShape(type, idx, `${PH_NAMES[type]} ${n}`, box);
  return { pres: { ...mv, slides: mv.slides.map((x) => (x.id === slideId ? { ...x, shapes: [...x.shapes, ph] } : x)) }, shapeId: ph.id };
}
