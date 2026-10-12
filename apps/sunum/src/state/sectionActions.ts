/** Bölüm komutları (Giriş > Bölüm, küçük resim ve sıralayıcı bölüm başlıkları). */
import { addSectionAt, moveSection, removeAllSections, removeSection, removeSectionWithSlides, renameSection, sectionAt, sectionsOf } from "../model/sections";
import type { Actions } from "./actions";

/** Seçili slayttan başlayan bölüm ekler; yeni bölümün kimliğini döndürür (yeniden adlandırma için). */
export function addSection(a: Actions): string | null {
  const s = a.s;
  if (!s.pres.slides.length) return null;
  const idx = s.slideIdx;
  const r = addSectionAt(s.pres.slides, idx);
  if (!r) {
    // Bu slaytta zaten bir bölüm başlıyor: onu yeniden adlandır.
    return s.pres.slides[idx].section?.id ?? null;
  }
  a.store.commit((p) => ({ ...p, slides: r.slides }));
  return r.id;
}

export function renameSectionCmd(a: Actions, id: string, name: string) {
  const n = name.trim();
  if (!n) return;
  a.store.commit((p) => ({ ...p, slides: renameSection(p.slides, id, n) }));
}

export function removeSectionCmd(a: Actions, id: string, withSlides = false) {
  const before = a.s.pres.slides[a.s.slideIdx]?.id;
  a.store.commit((p) => ({ ...p, slides: withSlides ? removeSectionWithSlides(p.slides, id) : removeSection(p.slides, id) }), { sel: [], editing: null });
  const after = a.s.pres.slides.findIndex((x) => x.id === before);
  a.goToSlide(after >= 0 ? after : Math.min(a.s.slideIdx, a.s.pres.slides.length - 1));
}

export function removeAllSectionsCmd(a: Actions) {
  a.store.commit((p) => ({ ...p, slides: removeAllSections(p.slides) }));
  a.store.set({ collapsedSections: [] });
}

export function moveSectionCmd(a: Actions, id: string, dir: -1 | 1) {
  const cur = a.s.pres.slides[a.s.slideIdx]?.id;
  a.store.commit((p) => ({ ...p, slides: moveSection(p.slides, id, dir) }));
  const i = a.s.pres.slides.findIndex((x) => x.id === cur);
  if (i >= 0) a.store.set({ slideIdx: i });
}

/** Bölümün tüm slaytlarını seçer. */
export function selectSection(a: Actions, id: string) {
  const sec = sectionsOf(a.s.pres.slides).find((s) => s.id === id);
  if (!sec) return;
  a.store.set({ slideIdx: sec.start, slideSel: sec.slideIds, sel: [], editing: null, focus: "slides" });
}

export function toggleSectionCollapsed(a: Actions, id: string) {
  a.store.set((s) => ({ collapsedSections: s.collapsedSections.includes(id) ? s.collapsedSections.filter((x) => x !== id) : [...s.collapsedSections, id] }));
}

export function collapseAllSections(a: Actions, collapse: boolean) {
  a.store.set({ collapsedSections: collapse ? sectionsOf(a.s.pres.slides).map((s) => s.id) : [] });
}

/** Geçerli slaydın bölümü. */
export function currentSection(a: Actions) {
  return sectionAt(a.s.pres.slides, a.s.slideIdx);
}
