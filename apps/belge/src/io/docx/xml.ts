// OOXML yardımcıları: ad alanları, kaçış, DOM erişimi.

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w10: "urn:schemas-microsoft-com:office:word",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
};

export const REL = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  app: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  custom: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
  fontTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
};

export function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** XML 1.0'da geçersiz denetim karakterlerini atar. */
export function cleanText(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "");
}

export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/* ------------------------------------------------------------------ */
/* Okuma tarafı                                                         */
/* ------------------------------------------------------------------ */

export function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error(`XML ayrıştırılamadı: ${err.textContent?.slice(0, 120)}`);
  return doc;
}

/** Öğenin yerel adı (ön ek yok sayılır). */
export const ln = (el: Element) => el.localName;

export function kids(el: Element | null | undefined, name?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c);
  return out;
}

export function kid(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
  return null;
}

/** Yol boyunca ilk alt öğeyi bulur: path(el, "pPr", "spacing"). */
export function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let cur: Element | null | undefined = el;
  for (const n of names) {
    cur = kid(cur, n);
    if (!cur) return null;
  }
  return cur ?? null;
}

/** Öznitelik (yerel adla; w: ön ekli olanlar önceliklidir). */
export function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  const w = el.getAttributeNS(NS.w, name);
  if (w !== null && w !== "") return w;
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
}

export function attrNum(el: Element | null | undefined, name: string): number | null {
  const v = attr(el, name);
  if (v === null || v === "") return null;
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : null;
}

/** w:b, w:i gibi aç/kapa özellikleri: öğe varsa ve val false/0/off değilse açık. */
export function onOff(el: Element | null): boolean | null {
  if (!el) return null;
  const v = attr(el, "val");
  return !(v === "0" || v === "false" || v === "off" || v === "none");
}

/** Tüm alt ağaçta yerel ada göre arama. */
export function descendants(el: Element, name: string): Element[] {
  const out: Element[] = [];
  const all = el.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) out.push(all[i]);
  return out;
}

export function relsTarget(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/");
  parts.pop();
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

export const HIGHLIGHT_NAMES: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#C0C0C0",
  black: "#000000",
  white: "#FFFFFF",
};

export function highlightName(hex: string): string | null {
  const h = hex.toUpperCase();
  for (const [k, v] of Object.entries(HIGHLIGHT_NAMES)) if (v === h) return k;
  return null;
}

/** CSS renk (#RGB, #RRGGBB, rgb()) → OOXML onaltılık (RRGGBB). */
export function hex6(c: string | null | undefined): string | null {
  if (!c) return null;
  const s = c.trim();
  let m = /^#?([0-9a-f]{6})$/i.exec(s);
  if (m) return m[1].toUpperCase();
  m = /^#?([0-9a-f]{3})$/i.exec(s);
  if (m) return m[1].split("").map((x) => x + x).join("").toUpperCase();
  m = /^rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(s);
  if (m) return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("").toUpperCase();
  return null;
}
