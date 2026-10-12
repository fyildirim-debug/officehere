/**
 * Seçili Slaytlara Uygula: bir temayı yalnız seçili slaytlara uygular. Tema,
 * kendi düzen kümesiyle (ikincil asıl slayt) gelir; asıl slaydın arka planı,
 * süs şekilleri ve metin stilleri bu düzenlere taşınır ve düzen `theme`
 * alanında temayı taşır. PPTX'e ayrı bir slideMaster + theme olarak yazılır.
 */
import { buildMasterAndLayouts } from "../themes/build";
import { getThemeDef, THEMES } from "../themes/themes";
import { uid } from "./factory";
import { applyTheme, relayoutSlide, themeFromDef } from "./presentation";
import { placeholderKind } from "./textStyle";
import type { Layout, LevelStyle, Master, Presentation, Shape } from "./types";

function mergeLevels(a: Partial<LevelStyle>[] | undefined, b: Partial<LevelStyle>[] | undefined): Partial<LevelStyle>[] | undefined {
  if (!a?.length) return b;
  if (!b?.length) return a;
  return Array.from({ length: Math.max(a.length, b.length) }, (_, i) => ({ para: { ...(a[i]?.para ?? {}), ...(b[i]?.para ?? {}) }, run: { ...(a[i]?.run ?? {}), ...(b[i]?.run ?? {}) } }) as Partial<LevelStyle>);
}

/** Asıl slaydın görünümünü (arka plan, süsler, metin stilleri) düzenlere katar. */
export function foldMasterIntoLayouts(master: Master, layouts: Layout[], theme: Layout["theme"]): Layout[] {
  const deco = master.shapes.filter((s) => !(s.kind === "sp" && s.ph));
  return layouts.map((l) => {
    let shapes: Shape[] = l.shapes.map((sh) => {
      if (sh.kind !== "sp" || !sh.ph) return sh;
      const own = master.textStyles[placeholderKind(sh.ph)] as Partial<LevelStyle>[];
      // Asıl slayttaki aynı türden yer tutucunun konumu/biçimi de düzene geçer.
      const mph = master.shapes.find((m) => m.kind === "sp" && m.ph?.type === sh.ph!.type);
      const levels = mergeLevels(mergeLevels(own, mph?.kind === "sp" ? mph.text?.levels : undefined), sh.text?.levels);
      return { ...sh, text: { ...(sh.text ?? { paragraphs: [{ runs: [], level: 0, props: {} }] }), levels } };
    });
    if (l.showMasterShapes) shapes = [...deco.map((d) => ({ ...d, id: uid("s") })), ...shapes];
    return { ...l, id: uid("L"), shapes, background: l.background ?? master.background, showMasterShapes: false, theme };
  });
}

/** Temayı yalnız verilen slaytlara uygular (tüm slaytlar seçiliyse sunuya uygular). */
export function applyThemeToSlides(pres: Presentation, themeId: string, slideIds: string[]): Presentation {
  const ids = new Set(slideIds);
  if (!ids.size || pres.slides.every((s) => ids.has(s.id))) return applyTheme(pres, themeId, 0);
  const def = getThemeDef(themeId) ?? THEMES[0];
  // Seçili slaytlar sununun kendi temasına dönüyorsa ana düzenler kullanılır.
  const toPrimary = !pres.masterEdited && pres.theme.id === def.id;
  let layouts = pres.layouts;
  let target: Layout[];
  if (toPrimary) target = pres.layouts.filter((l) => !l.theme);
  else {
    const existing = pres.layouts.filter((l) => l.theme?.id === def.id);
    if (existing.length) target = existing;
    else {
      const built = buildMasterAndLayouts(pres.width, pres.height, def.style);
      target = foldMasterIntoLayouts(built.master, built.layouts, themeFromDef(def.id, 0));
      layouts = [...layouts, ...target];
    }
  }
  const next: Presentation = { ...pres, layouts };
  next.slides = pres.slides.map((s) => {
    if (!ids.has(s.id)) return s;
    const old = layouts.find((l) => l.id === s.layoutId);
    const nl = target.find((l) => l.type === old?.type) ?? target.find((l) => l.type === "obj") ?? target[0];
    return nl && nl.id !== old?.id ? relayoutSlide(next, s, nl, old) : s;
  });
  return pruneDesignLayouts(next);
}

/** Hiçbir slaydın kullanmadığı ikincil tema düzenlerini kaldırır. */
export function pruneDesignLayouts(pres: Presentation): Presentation {
  const used = new Set(pres.slides.map((s) => s.layoutId));
  const usedThemes = new Set(pres.layouts.filter((l) => l.theme && used.has(l.id)).map((l) => l.theme!.id ?? l.theme!.name));
  // Yalnız uygulamanın kendi temalarından üretilmiş düzenler kaldırılır; içe aktarılmış asıl slaytlar korunur.
  const layouts = pres.layouts.filter((l) => !l.theme || !getThemeDef(l.theme.id ?? "") || usedThemes.has(l.theme.id ?? l.theme.name));
  return layouts.length === pres.layouts.length ? pres : { ...pres, layouts };
}

/** Slaytın etkin teması. */
export function slideTheme(pres: Presentation, layoutId: string | undefined) {
  return pres.layouts.find((l) => l.id === layoutId)?.theme ?? pres.theme;
}
