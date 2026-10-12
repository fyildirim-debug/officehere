import { buildMasterAndLayouts } from "../themes/build";
import { getThemeDef, THEMES } from "../themes/themes";
import { emptyBody, uid } from "./factory";
import type { HeaderFooter, Layout, LayoutType, Presentation, Shape, Slide, SpShape, Theme } from "./types";
import { findPlaceholder } from "./textStyle";

export const SLIDE_16_9 = { width: 1280, height: 720 };
export const SLIDE_4_3 = { width: 960, height: 720 };

export const DEFAULT_HF: HeaderFooter = {
  dateTime: false,
  dateAuto: true,
  dateFixed: "",
  slideNumber: false,
  footer: false,
  footerText: "",
  hideOnTitle: true,
};

export function themeFromDef(id: string, variant = 0): Theme {
  const def = getThemeDef(id) ?? THEMES[0];
  const v = def.variants[Math.min(variant, def.variants.length - 1)];
  return { name: def.name, id: def.id, variant, colors: { ...v.colors }, fonts: { ...def.fonts } };
}

export function newPresentation(themeId = "varsayilan", variant = 0, size = SLIDE_16_9): Presentation {
  const def = getThemeDef(themeId) ?? THEMES[0];
  const { master, layouts } = buildMasterAndLayouts(size.width, size.height, def.style);
  const pres: Presentation = {
    width: size.width,
    height: size.height,
    theme: themeFromDef(def.id, variant),
    master,
    layouts,
    slides: [],
    headerFooter: { ...DEFAULT_HF },
    props: { created: new Date().toISOString() },
  };
  pres.slides.push(newSlide(pres, "title"));
  return pres;
}

export function layoutByType(pres: Presentation, type: LayoutType): Layout {
  return pres.layouts.find((l) => l.type === type) ?? pres.layouts.find((l) => l.type === "obj") ?? pres.layouts[0];
}

const FOOTER_TYPES = new Set(["dt", "ftr", "sldNum"]);

/** Düzen yer tutucusundan slayt yer tutucusu üretir. */
export function instantiatePlaceholder(lph: SpShape): SpShape {
  return {
    kind: "sp",
    id: uid(),
    name: lph.name,
    x: lph.x,
    y: lph.y,
    w: lph.w,
    h: lph.h,
    rot: 0,
    geom: "rect",
    fill: { type: "none" },
    line: { fill: { type: "none" }, width: 1, dash: "solid" },
    ph: { ...lph.ph! },
    text: emptyBody(),
  };
}

export function footerShapes(pres: Presentation, layout: Layout, slideIsTitle: boolean): SpShape[] {
  const hf = pres.headerFooter;
  if (slideIsTitle && hf.hideOnTitle) return [];
  const out: SpShape[] = [];
  for (const s of layout.shapes) {
    if (s.kind !== "sp" || !s.ph || !FOOTER_TYPES.has(s.ph.type)) continue;
    if (s.ph.type === "dt" && !hf.dateTime) continue;
    if (s.ph.type === "ftr" && !hf.footer) continue;
    if (s.ph.type === "sldNum" && !hf.slideNumber) continue;
    const sh = instantiatePlaceholder(s);
    if (s.ph.type === "ftr") sh.text = { paragraphs: [{ runs: [{ text: hf.footerText, props: {} }], level: 0, props: {} }] };
    if (s.ph.type === "sldNum") sh.text = { paragraphs: [{ runs: [{ text: "‹#›", props: { field: "slidenum" } }], level: 0, props: {} }] };
    if (s.ph.type === "dt")
      sh.text = { paragraphs: [{ runs: [{ text: hf.dateAuto ? formatToday() : hf.dateFixed, props: hf.dateAuto ? { field: "datetime" } : {} }], level: 0, props: {} }] };
    out.push(sh);
  }
  return out;
}

export function formatToday(): string {
  return new Intl.DateTimeFormat("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(new Date());
}

export function newSlide(pres: Presentation, layoutType: LayoutType | string = "obj"): Slide {
  const layout = pres.layouts.find((l) => l.id === layoutType) ?? layoutByType(pres, layoutType as LayoutType);
  const shapes: Shape[] = [];
  for (const s of layout.shapes) {
    if (s.kind === "sp" && s.ph && !FOOTER_TYPES.has(s.ph.type)) shapes.push(instantiatePlaceholder(s));
  }
  shapes.push(...footerShapes(pres, layout, layout.type === "title"));
  return { id: uid("sl"), layoutId: layout.id, shapes, notes: "", animations: [] };
}

function sameGeom(a: Shape, b: Shape) {
  return Math.abs(a.x - b.x) < 0.6 && Math.abs(a.y - b.y) < 0.6 && Math.abs(a.w - b.w) < 0.6 && Math.abs(a.h - b.h) < 0.6;
}

/**
 * Slaytı yeni düzene geçirir: eşleşen yer tutucular korunur (konumları
 * değiştirilmemişse yeni düzenin konumuna taşınır), eşleşmeyen boş yer
 * tutucular silinir, eksik yer tutucular eklenir.
 */
export function relayoutSlide(_pres: Presentation, slide: Slide, newLayout: Layout, oldLayout?: Layout, reset = false): Slide {
  const used = new Set<SpShape>();
  const shapes: Shape[] = [];
  for (const s of slide.shapes) {
    if (s.kind !== "sp" || !s.ph) {
      shapes.push(s);
      continue;
    }
    if (FOOTER_TYPES.has(s.ph.type)) {
      const target = findPlaceholder(newLayout.shapes, s.ph);
      shapes.push(target ? { ...s, x: target.x, y: target.y, w: target.w, h: target.h } : s);
      continue;
    }
    const target = newLayout.shapes.find(
      (l): l is SpShape => l.kind === "sp" && !!l.ph && !used.has(l as SpShape) && findPlaceholder([l], s.ph!) === l,
    ) as SpShape | undefined;
    const empty = !s.text || s.text.paragraphs.every((p) => p.runs.every((r) => !r.text));
    if (!target) {
      if (!empty) shapes.push(s);
      continue;
    }
    used.add(target);
    const old = oldLayout ? findPlaceholder(oldLayout.shapes, s.ph) : undefined;
    const moveIt = reset || !old || sameGeom(old, s);
    shapes.push({
      ...s,
      ph: { ...target.ph! },
      ...(moveIt ? { x: target.x, y: target.y, w: target.w, h: target.h, rot: 0 } : {}),
      ...(reset && s.text ? { text: { ...s.text, paragraphs: s.text.paragraphs.map((p) => ({ ...p, props: {}, runs: p.runs.map((r) => ({ ...r, props: r.props.field ? { field: r.props.field } : {} })) })) } } : {}),
    });
  }
  for (const l of newLayout.shapes) {
    if (l.kind === "sp" && l.ph && !FOOTER_TYPES.has(l.ph.type) && !used.has(l)) {
      shapes.push(instantiatePlaceholder(l));
    }
  }
  return { ...slide, layoutId: newLayout.id, shapes };
}

/** Temayı (ve çeşitlemeyi) tüm sunuya uygular. */
export function applyTheme(pres: Presentation, themeId: string, variant = 0): Presentation {
  const def = getThemeDef(themeId) ?? THEMES[0];
  const { master, layouts } = buildMasterAndLayouts(pres.width, pres.height, def.style);
  const next: Presentation = { ...pres, theme: themeFromDef(def.id, variant), master, layouts };
  next.slides = pres.slides.map((s) => {
    const old = pres.layouts.find((l) => l.id === s.layoutId);
    const nl = layouts.find((l) => l.type === old?.type) ?? layouts.find((l) => l.type === "obj")!;
    return relayoutSlide(next, s, nl, old);
  });
  return next;
}

/** Yalnız renk çeşitlemesini değiştirir. */
export function applyVariant(pres: Presentation, variant: number): Presentation {
  const def = getThemeDef(pres.theme.id);
  if (!def) return pres;
  return { ...pres, theme: themeFromDef(def.id, variant) };
}

/** Slayt boyutunu değiştirir ve içeriği ölçekler. */
export function resizePresentation(pres: Presentation, width: number, height: number, mode: "scale" | "fit" = "fit"): Presentation {
  const sx = width / pres.width;
  const sy = height / pres.height;
  const s = mode === "fit" ? Math.min(sx, sy) : Math.max(sx, sy);
  const ox = (width - pres.width * s) / 2;
  const oy = (height - pres.height * s) / 2;
  const tf = (sh: Shape): Shape => {
    const base = { ...sh, x: sh.x * s + ox, y: sh.y * s + oy, w: sh.w * s, h: sh.h * s };
    if (sh.kind === "group") return { ...base, children: sh.children.map(tf) } as Shape;
    if (sh.kind === "table") return { ...base, cols: sh.cols.map((c) => c * s), rows: sh.rows.map((r) => ({ ...r, h: r.h * s })) } as Shape;
    return base as Shape;
  };
  const def = getThemeDef(pres.theme.id);
  let master = pres.master;
  let layouts = pres.layouts;
  if (def) {
    const built = buildMasterAndLayouts(width, height, def.style);
    master = built.master;
    layouts = pres.layouts.map((l) => {
      const nl = built.layouts.find((b) => b.type === l.type);
      return nl ? { ...nl, id: l.id } : { ...l, shapes: l.shapes.map(tf) };
    });
  } else {
    master = { ...pres.master, shapes: pres.master.shapes.map(tf) };
    layouts = pres.layouts.map((l) => ({ ...l, shapes: l.shapes.map(tf) }));
  }
  return { ...pres, width, height, master, layouts, slides: pres.slides.map((sl) => ({ ...sl, shapes: sl.shapes.map(tf) })) };
}
