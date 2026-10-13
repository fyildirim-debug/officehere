/** XML yazma/okuma yardımcıları (önek bağımsız, yerel ad tabanlı). */

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  oh: "urn:officehere:sunum:1",
  /** PowerPoint 2010 uzantıları (bölümler, ortam). */
  p14: "http://schemas.microsoft.com/office/powerpoint/2010/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
};

/** p:presentation/p:extLst içindeki bölüm listesi uzantısı. */
export const SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
/** p:nvPr/p:extLst içindeki gömülü ortam (p14:media) uzantısı. */
export const MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

export const REL = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  notesSlide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  notesMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
  presProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
  viewProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
  tableStyles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
  coreProps: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  commentAuthors: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  package: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
  video: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
  audio: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
  /** PowerPoint 2010 gömülü ortam ilişkisi (p14:media). */
  media: "http://schemas.microsoft.com/office/2007/relationships/media",
};

/** Uygulamaya özgü uzantı kimliği (ECMA-376 extLst mekanizması). */
export const OH_EXT_URI = "{0FF1CE00-5A5A-4E7E-9B0B-5A53554E554D}";

const ESC: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" };

export function esc(s: string | number | undefined | null): string {
  if (s === undefined || s === null) return "";
  // XML 1.0'da geçersiz denetim karakterlerini at.
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/[&<>"']/g, (c) => ESC[c]);
}

/** Öznitelik dizgesi: tanımsız değerler atlanır. */
export function attrs(o: Record<string, string | number | boolean | undefined | null>): string {
  let s = "";
  for (const [k, v] of Object.entries(o)) {
    if (v === undefined || v === null || v === false) continue;
    s += ` ${k}="${esc(v === true ? "1" : v)}"`;
  }
  return s;
}

export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function parseXml(text: string): Document {
  return new DOMParser().parseFromString(text, "application/xml");
}

export function kids(el: Element | null | undefined, name?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (let n = el.firstElementChild; n; n = n.nextElementSibling) if (!name || n.localName === name) out.push(n);
  return out;
}

export function kid(el: Element | null | undefined, name: string): Element | undefined {
  if (!el) return undefined;
  for (let n = el.firstElementChild; n; n = n.nextElementSibling) if (n.localName === name) return n;
  return undefined;
}

/** Yol boyunca ilk eşleşen torun: kidPath(el, "spPr", "xfrm"). */
export function kidPath(el: Element | null | undefined, ...names: string[]): Element | undefined {
  let cur: Element | undefined = el ?? undefined;
  for (const n of names) {
    cur = kid(cur, n);
    if (!cur) return undefined;
  }
  return cur;
}

export function attr(el: Element | null | undefined, name: string): string | undefined {
  if (!el) return undefined;
  if (el.hasAttribute(name)) return el.getAttribute(name) ?? undefined;
  // Önekli öznitelik (r:embed vb.)
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return undefined;
}

/** İlişki kimliği (r:id) — aynı adlı "id" özniteliğiyle karışmasın diye ad alanıyla okunur. */
export function relId(el: Element | null | undefined): string | undefined {
  if (!el) return undefined;
  const v = el.getAttributeNS(NS.r, "id");
  if (v) return v;
  for (const a of Array.from(el.attributes)) if (a.localName === "id" && a.prefix) return a.value;
  return undefined;
}

export function numAttr(el: Element | null | undefined, name: string): number | undefined {
  const v = attr(el, name);
  if (v === undefined || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

export function boolAttr(el: Element | null | undefined, name: string): boolean | undefined {
  const v = attr(el, name);
  if (v === undefined) return undefined;
  return v === "1" || v === "true" || v === "on";
}

export function descendants(el: Element | null | undefined, name: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  const walk = (n: Element) => {
    for (let c = n.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === name) out.push(c);
      walk(c);
    }
  };
  walk(el);
  return out;
}

/** Göreli hedef yolunu çözer ("../media/image1.png" → "ppt/media/image1.png"). */
export function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/");
  parts.pop();
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

export function relsPathFor(part: string): string {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
}

export interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export function parseRels(xml: string | undefined, partPath: string): Map<string, Rel> {
  const map = new Map<string, Rel>();
  if (!xml) return map;
  const doc = parseXml(xml);
  for (const r of kids(doc.documentElement, "Relationship")) {
    const id = attr(r, "Id") ?? "";
    const external = attr(r, "TargetMode") === "External";
    const target = attr(r, "Target") ?? "";
    map.set(id, { id, type: attr(r, "Type") ?? "", target: external ? target : resolvePath(partPath, target), external });
  }
  return map;
}

export class RelsBuilder {
  private list: { id: string; type: string; target: string; external?: boolean }[] = [];
  add(type: string, target: string, external = false): string {
    const existing = this.list.find((r) => r.type === type && r.target === target);
    if (existing) return existing.id;
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, external });
    return id;
  }
  xml(): string {
    return (
      XML_HEAD +
      `<Relationships xmlns="${NS.rel}">` +
      this.list.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("") +
      "</Relationships>"
    );
  }
  get size() {
    return this.list.length;
  }
}

export const EMU_PX = 9525;
export const emu = (px: number) => Math.round(px * EMU_PX);
export const fromEmu = (v: number | undefined, def = 0) => (v === undefined ? def : v / EMU_PX);
