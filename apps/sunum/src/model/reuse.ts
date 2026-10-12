/**
 * Slaytları Yeniden Kullan: başka bir sunudaki slaytı bu sunuya ekler.
 * Varsayılan olarak hedef tema/düzen kullanılır (yer tutucular hedefin
 * biçimini alır); "Kaynak biçimlendirmesini koru" seçilirse tema renkleri,
 * tema yazı tipleri, arka plan ve yer tutucu düzey stilleri kaynaktan
 * sabit değerlere çevrilerek taşınır.
 */
import { resolveHex } from "./color";
import { cloneSlide } from "./ops";
import { layoutByType } from "./presentation";
import { levelResolver, resolveFont } from "./textStyle";
import type { Color, Fill, Layout, LevelStyle, Presentation, Shape, Slide, Theme } from "./types";

/** Nesne ağacındaki tema renklerini (scheme) sabit RGB'ye çevirir. */
export function freezeColors<T>(v: T, theme: Theme): T {
  if (Array.isArray(v)) return v.map((x) => freezeColors(x, theme)) as unknown as T;
  if (!v || typeof v !== "object") return v;
  const o = v as Record<string, unknown>;
  if (typeof o.scheme === "string" && !("rgb" in o && o.rgb)) {
    const c = o as Color;
    const out: Color = { rgb: resolveHex(c, theme) };
    if (c.alpha !== undefined) out.alpha = c.alpha;
    return out as unknown as T;
  }
  const r: Record<string, unknown> = {};
  for (const [k, x] of Object.entries(o)) {
    if (k === "font" && typeof x === "string" && x.startsWith("+")) r[k] = resolveFont(x, theme);
    else r[k] = freezeColors(x, theme);
  }
  return r as T;
}

function scaleShape(sh: Shape, s: number, ox: number, oy: number): Shape {
  const base = { ...sh, x: sh.x * s + ox, y: sh.y * s + oy, w: sh.w * s, h: sh.h * s };
  if (sh.kind === "group") return { ...base, children: sh.children.map((c) => scaleShape(c, s, ox, oy)) } as Shape;
  if (sh.kind === "table") return { ...base, cols: sh.cols.map((c) => c * s), rows: sh.rows.map((r) => ({ ...r, h: r.h * s, minH: r.minH && r.minH * s })) } as Shape;
  return base as Shape;
}

/** Kaynak düzene karşılık gelen hedef düzen (aynı tür, sonra aynı ad, yoksa Başlık ve İçerik). */
export function matchLayout(dst: Presentation, srcLayout: Layout | undefined): Layout {
  if (srcLayout) {
    const byType = srcLayout.type !== "cust" ? dst.layouts.find((l) => l.type === srcLayout.type) : undefined;
    if (byType) return byType;
    const byName = dst.layouts.find((l) => l.name === srcLayout.name);
    if (byName) return byName;
  }
  return layoutByType(dst, "obj");
}

export function reuseSlide(dst: Presentation, src: Presentation, srcSlide: Slide, keepSource: boolean): Slide {
  const srcLayout = src.layouts.find((l) => l.id === srcSlide.layoutId);
  const layout = matchLayout(dst, srcLayout);
  let c = cloneSlide(srcSlide);
  c = { ...c, layoutId: layout.id, section: undefined };
  if (src.width !== dst.width || src.height !== dst.height) {
    const s = Math.min(dst.width / src.width, dst.height / src.height);
    const ox = (dst.width - src.width * s) / 2;
    const oy = (dst.height - src.height * s) / 2;
    c = { ...c, shapes: c.shapes.map((sh) => scaleShape(sh, s, ox, oy)) };
  }
  if (!keepSource) return c;
  const theme = src.theme;
  const tctx = { master: src.master, layout: srcLayout };
  // Yer tutucuların kalıtımla gelen düzey stilleri şekle yazılır.
  const shapes = c.shapes.map((sh) => {
    if (sh.kind !== "sp" || !sh.ph || !sh.text) return sh;
    const srcShape = srcSlide.shapes.find((x) => x.kind === "sp" && x.ph?.type === sh.ph?.type && x.ph?.idx === sh.ph?.idx);
    const lvl = levelResolver(srcShape && srcShape.kind === "sp" ? srcShape : sh, tctx);
    const levels: Partial<LevelStyle>[] = [];
    for (let i = 0; i < 5; i++) {
      const st = lvl(i);
      levels.push({ para: { ...st.para, ...(sh.text.levels?.[i]?.para ?? {}) }, run: { ...st.run, ...(sh.text.levels?.[i]?.run ?? {}) } });
    }
    return { ...sh, text: { ...sh.text, levels } };
  });
  const bg: Fill | undefined = srcSlide.background ?? srcLayout?.background ?? src.master.background;
  return freezeColors({ ...c, shapes, background: bg }, theme);
}
