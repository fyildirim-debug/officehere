/**
 * Köprüler: metin çalıştırmalarında (RunProps.link) ve şekillerde
 * (ShapeBase.link). Dış adresler olduğu gibi tutulur; sunu içi hedefler
 * "#slide:<kimlik>", "#next", "#prev", "#first", "#last" biçimindedir.
 */
import type { Presentation, Shape, Slide, TextBody } from "./types";

export type LinkTarget =
  | { kind: "url"; url: string }
  | { kind: "slide"; slideId: string }
  | { kind: "jump"; jump: "next" | "prev" | "first" | "last" };

const JUMPS = new Set(["next", "prev", "first", "last"]);

export function parseLink(link: string): LinkTarget {
  if (link.startsWith("#slide:")) return { kind: "slide", slideId: link.slice(7) };
  if (link.startsWith("#") && JUMPS.has(link.slice(1))) return { kind: "jump", jump: link.slice(1) as "next" };
  return { kind: "url", url: link };
}

export const JUMP_LABELS: Record<string, string> = {
  first: "İlk Slayt",
  last: "Son Slayt",
  next: "Sonraki Slayt",
  prev: "Önceki Slayt",
};

/** Kullanıcının yazdığı adresi tamamlar (www.… → https://www.…). */
export function normalizeUrl(input: string): string {
  const t = input.trim();
  if (!t) return "";
  if (/^[a-z][a-z0-9+.-]*:/i.test(t) || t.startsWith("#")) return t;
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(t)) return `mailto:${t}`;
  return `https://${t}`;
}

/** Köprünün okunur açıklaması (ipucu, menü). */
export function linkLabel(link: string, pres: Presentation): string {
  const t = parseLink(link);
  if (t.kind === "jump") return JUMP_LABELS[t.jump];
  if (t.kind === "slide") {
    const i = pres.slides.findIndex((s) => s.id === t.slideId);
    return i >= 0 ? `Slayt ${i + 1}` : "Slayt (silinmiş)";
  }
  return t.url.replace(/^mailto:/, "");
}

/**
 * Gösteride köprünün hedef slayt dizini; dış adresse null.
 * Sunu içi hedef bulunamazsa -1 döner.
 */
export function linkSlideIndex(link: string, pres: Presentation, current: number): number | null {
  const t = parseLink(link);
  if (t.kind === "url") return null;
  const n = pres.slides.length;
  if (t.kind === "slide") return pres.slides.findIndex((s) => s.id === t.slideId);
  switch (t.jump) {
    case "first":
      return 0;
    case "last":
      return n - 1;
    case "next":
      return current + 1 < n ? current + 1 : -1;
    case "prev":
      return current > 0 ? current - 1 : -1;
  }
}

function mapBody(b: TextBody | undefined, fn: (link: string) => string | undefined): TextBody | undefined {
  if (!b) return b;
  let changed = false;
  const paragraphs = b.paragraphs.map((p) => {
    let pc = false;
    const runs = p.runs.map((r) => {
      if (!r.props.link) return r;
      const nl = fn(r.props.link);
      if (nl === r.props.link) return r;
      pc = true;
      const props = { ...r.props };
      if (nl) props.link = nl;
      else delete props.link;
      return { ...r, props };
    });
    if (!pc) return p;
    changed = true;
    return { ...p, runs };
  });
  return changed ? { ...b, paragraphs } : b;
}

/** Şekil ağacındaki tüm köprüleri (şekil ve metin) dönüştürür; undefined dönerse köprü kaldırılır. */
export function mapShapeLinks(s: Shape, fn: (link: string) => string | undefined): Shape {
  let out: Shape = s;
  if (s.link) {
    const nl = fn(s.link);
    if (nl !== s.link) {
      out = { ...s };
      if (nl) out.link = nl;
      else delete out.link;
    }
  }
  if (out.hoverLink) {
    const nh = fn(out.hoverLink);
    if (nh !== out.hoverLink) {
      out = { ...out };
      if (nh) out.hoverLink = nh;
      else delete out.hoverLink;
    }
  }
  // Yakınlaştırma şeklinin hedefi köprüsüyle birlikte taşınır.
  if (out.kind === "pic" && out.zoom && out.link?.startsWith("#slide:") && out.zoom.slideId !== out.link.slice(7)) out = { ...out, zoom: { ...out.zoom, slideId: out.link.slice(7) } };
  if (out.kind === "sp" && out.text) {
    const t = mapBody(out.text, fn);
    if (t !== out.text) out = { ...out, text: t };
  } else if (out.kind === "table") {
    let changed = false;
    const rows = out.rows.map((r) => {
      let rc = false;
      const cells = r.cells.map((c) => {
        const t = mapBody(c.text, fn)!;
        if (t === c.text) return c;
        rc = true;
        return { ...c, text: t };
      });
      if (!rc) return r;
      changed = true;
      return { ...r, cells };
    });
    if (changed) out = { ...out, rows };
  } else if (out.kind === "group") {
    const g = out;
    const children = g.children.map((c) => mapShapeLinks(c, fn));
    if (children.some((c, i) => c !== g.children[i])) out = { ...g, children };
  }
  return out;
}

export function mapSlideLinks(slide: Slide, fn: (link: string) => string | undefined): Slide {
  const shapes = slide.shapes.map((s) => mapShapeLinks(s, fn));
  return shapes.some((s, i) => s !== slide.shapes[i]) ? { ...slide, shapes } : slide;
}

/** Şekil ağacındaki köprüleri toplar. */
export function collectLinks(shapes: Shape[], out: string[] = []): string[] {
  for (const s of shapes)
    mapShapeLinks(s, (l) => {
      out.push(l);
      return l;
    });
  return out;
}
