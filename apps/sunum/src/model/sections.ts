/**
 * Bölümler: her bölüm, ilk slaydındaki işaretle (Slide.section) başlar ve
 * bir sonraki işarete kadar sürer. Bölüm varsa ilk slayt her zaman bir
 * bölüm başlatır (PowerPoint'teki "Varsayılan Bölüm" davranışı).
 */
import { uid } from "./factory";
import type { Slide } from "./types";

export interface SectionInfo {
  id: string;
  name: string;
  /** İlk slaydın dizini. */
  start: number;
  /** Bitiş (hariç). */
  end: number;
  slideIds: string[];
}

export const DEFAULT_SECTION_NAME = "Varsayılan Bölüm";
export const NEW_SECTION_NAME = "Adsız Bölüm";

export function hasSections(slides: Slide[]): boolean {
  return slides.some((s) => s.section);
}

export function sectionsOf(slides: Slide[]): SectionInfo[] {
  const out: SectionInfo[] = [];
  slides.forEach((s, i) => {
    if (s.section) out.push({ id: s.section.id, name: s.section.name, start: i, end: i + 1, slideIds: [s.id] });
    else if (out.length) {
      const cur = out[out.length - 1];
      cur.end = i + 1;
      cur.slideIds.push(s.id);
    }
  });
  return out;
}

/** Slaydın bulunduğu bölüm. */
export function sectionAt(slides: Slide[], index: number): SectionInfo | undefined {
  return sectionsOf(slides).find((s) => index >= s.start && index < s.end);
}

function withMarker(s: Slide, m: Slide["section"]): Slide {
  if (s.section === m) return s;
  const n = { ...s };
  if (m) n.section = m;
  else delete n.section;
  return n;
}

/** Bölüm varsa ilk slaydın bir bölüm başlatmasını sağlar. */
export function normalizeSections(slides: Slide[]): Slide[] {
  if (!slides.length || slides[0].section) return slides;
  const k = slides.findIndex((s) => s.section);
  if (k < 0) return slides;
  // İlk işaretten önceki slaytlar ilk bölüme aittir: işaret başa taşınır.
  const out = slides.slice();
  out[0] = withMarker(out[0], out[k].section);
  out[k] = withMarker(out[k], undefined);
  return out;
}

/**
 * Silinecek ya da taşınacak slaytların işaretlerini bölümde kalan bir
 * sonraki slayta devreder; işaretli slaytlardan işaret kaldırılır.
 */
export function releaseMarkers(slides: Slide[], ids: Set<string>): Slide[] {
  let carry: Slide["section"];
  let changed = false;
  const out = slides.map((s) => {
    if (ids.has(s.id)) {
      if (s.section) {
        carry ??= s.section;
        changed = true;
        return withMarker(s, undefined);
      }
      return s;
    }
    if (carry) {
      const m = carry;
      carry = undefined;
      if (!s.section) {
        changed = true;
        return withMarker(s, m);
      }
    }
    return s;
  });
  return changed ? out : slides;
}

/** Slaytları siler (bölüm işaretleri korunarak). */
export function removeSlidesKeepSections(slides: Slide[], ids: Set<string>): Slide[] {
  return normalizeSections(releaseMarkers(slides, ids).filter((s) => !ids.has(s.id)));
}

/** Slaytları taşır; taşınan slaytlar bırakıldıkları yerin bölümüne katılır. */
export function moveSlidesKeepSections(slides: Slide[], ids: string[], toIndex: number): Slide[] {
  const set = new Set(ids);
  const released = releaseMarkers(slides, set);
  const moving = released.filter((s) => set.has(s.id));
  const before = released.slice(0, toIndex).filter((s) => !set.has(s.id));
  const after = released.slice(toIndex).filter((s) => !set.has(s.id));
  return normalizeSections([...before, ...moving, ...after]);
}

/** Verilen slayttan başlayan yeni bölüm ekler; ilk slayt bölümsüzse "Varsayılan Bölüm" oluşturulur. */
export function addSectionAt(slides: Slide[], index: number, name = NEW_SECTION_NAME): { slides: Slide[]; id: string } | null {
  if (index < 0 || index >= slides.length || slides[index].section) return null;
  const id = uid("sec");
  const out = slides.slice();
  out[index] = withMarker(out[index], { id, name });
  if (index > 0 && !out[0].section) out[0] = withMarker(out[0], { id: uid("sec"), name: DEFAULT_SECTION_NAME });
  return { slides: out, id };
}

export function renameSection(slides: Slide[], id: string, name: string): Slide[] {
  return slides.map((s) => (s.section?.id === id ? withMarker(s, { ...s.section, name }) : s));
}

/** Bölümü kaldırır; slaytları önceki bölüme (ilk bölümse sonrakine) katılır. */
export function removeSection(slides: Slide[], id: string): Slide[] {
  const secs = sectionsOf(slides);
  const i = secs.findIndex((s) => s.id === id);
  if (i < 0) return slides;
  const out = slides.slice();
  const sec = secs[i];
  if (i === 0 && secs.length > 1) {
    // İlk bölüm kaldırılınca slaytları bir sonraki bölüme katılır.
    const next = secs[1];
    out[sec.start] = withMarker(out[sec.start], out[next.start].section);
    out[next.start] = withMarker(out[next.start], undefined);
    return out;
  }
  out[sec.start] = withMarker(out[sec.start], undefined);
  return out;
}

/** Bölümü ve slaytlarını siler. */
export function removeSectionWithSlides(slides: Slide[], id: string): Slide[] {
  const sec = sectionsOf(slides).find((s) => s.id === id);
  if (!sec) return slides;
  const ids = new Set(sec.slideIds);
  return normalizeSections(slides.filter((s) => !ids.has(s.id)));
}

export function removeAllSections(slides: Slide[]): Slide[] {
  return slides.map((s) => (s.section ? withMarker(s, undefined) : s));
}

/** Bölümü (slaytlarıyla) bir önceki/sonraki bölümle yer değiştirir. */
export function moveSection(slides: Slide[], id: string, dir: -1 | 1): Slide[] {
  const secs = sectionsOf(slides);
  const i = secs.findIndex((s) => s.id === id);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= secs.length) return slides;
  const [a, b] = i < j ? [secs[i], secs[j]] : [secs[j], secs[i]];
  return [...slides.slice(0, a.start), ...slides.slice(b.start, b.end), ...slides.slice(a.start, a.end), ...slides.slice(b.end)];
}
