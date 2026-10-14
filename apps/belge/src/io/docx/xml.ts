// OOXML yardımcıları: ad alanları, kaçış, DOM erişimi.
import { BelgeOpenError } from "../warnings";

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

/** İç içe öğe düzeyi sınırı: tarayıcıların XML ayrıştırıcıları ~256 düzeyde durur; okuyucular özyinelemelidir. */
export const MAX_XML_DEPTH = 200;

/** Derinlik sınırı aşılınca düzleştirilen belgeler (okuyucu uyarı verir). */
export const flattenedDocs = new WeakSet<Document>();

const TOKEN = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE(?:[^>[]|\[[^\]]*\])*>|<(\/?)[^\s>/!?][^>]*?(\/?)>/g;

/**
 * Sınırdan derin öğelerin etiketleri atılır (içerikleri sınırdaki öğeye kalır); böylece aşırı iç içe
 * (bozuk ya da kötü niyetli) belge ayrıştırıcıyı ve özyinelemeli okuyucuları kilitlemez. Derinlik
 * aşılmıyorsa metin olduğu gibi döner.
 */
export function limitDepth(text: string, max = MAX_XML_DEPTH): { text: string; flattened: boolean } {
  let depth = 0;
  let over = false;
  TOKEN.lastIndex = 0;
  for (let m = TOKEN.exec(text); m; m = TOKEN.exec(text)) {
    if (m[1] === undefined) continue;
    if (m[1] === "/") depth--;
    else if (m[2] !== "/" && ++depth > max) {
      over = true;
      break;
    }
  }
  if (!over) return { text, flattened: false };
  const out: string[] = [];
  let last = 0;
  depth = 0;
  // Sınırın ötesinde açılan etiketlerin kapanışları da atılır (yığın yerine düzey sayacı yeterli).
  TOKEN.lastIndex = 0;
  for (let m = TOKEN.exec(text); m; m = TOKEN.exec(text)) {
    if (m[1] === undefined) continue;
    let drop = false;
    if (m[1] === "/") {
      drop = depth > max;
      depth--;
    } else if (m[2] !== "/") {
      depth++;
      drop = depth > max;
    } else drop = depth >= max;
    if (drop) {
      out.push(text.slice(last, m.index));
      last = m.index + m[0].length;
    }
  }
  out.push(text.slice(last));
  return { text: out.join(""), flattened: true };
}

function tryParse(text: string): { doc: Document; error: string | null } {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  return { doc, error: err ? (err.textContent ?? "").slice(0, 160) : null };
}

/**
 * Bildirilmemiş ad alanı ön ekleri (bazı üreticilerin ve bozuk dosyaların hatası) kök öğeye yer tutucu
 * ad alanıyla eklenir; böylece bilinmeyen öğeler yok sayılır, belgenin geri kalanı okunur.
 */
function declareMissingPrefixes(text: string): string | null {
  const declared = new Set(["xml", "xmlns"]);
  for (const m of text.matchAll(/xmlns:([A-Za-z_][\w.-]*)\s*=/g)) declared.add(m[1]);
  const missing = new Set<string>();
  for (const m of text.matchAll(/<\/?([A-Za-z_][\w.-]*):[A-Za-z_]/g)) if (!declared.has(m[1])) missing.add(m[1]);
  for (const m of text.matchAll(/\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=\s*["']/g)) if (!declared.has(m[1])) missing.add(m[1]);
  if (!missing.size || missing.size > 64) return null;
  const root = /<([A-Za-z_][\w.:-]*)/.exec(text.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>/g, (s) => " ".repeat(s.length)));
  if (!root) return null;
  const at = root.index + root[0].length;
  const decl = [...missing].map((p) => ` xmlns:${p}="urn:officehere:bilinmeyen:${p}"`).join("");
  return text.slice(0, at) + decl + text.slice(at);
}

export function parseXml(source: string): Document {
  const { text, flattened } = limitDepth(source);
  const done = (doc: Document) => {
    if (flattened) flattenedDocs.add(doc);
    return doc;
  };
  const r = tryParse(text);
  if (r.error) {
    const fixed = declareMissingPrefixes(text);
    if (fixed) {
      const again = tryParse(fixed);
      if (!again.error) return done(again.doc);
    }
    throw new Error(`XML ayrıştırılamadı: ${r.error.slice(0, 120)}`);
  }
  return done(r.doc);
}

/** XML parçasının izin verilen en büyük açılmış boyutu (zip bombasına karşı). */
export const MAX_XML_PART = 512 * 1024 * 1024;
/** Bu boyutun üstünde 200 kattan fazla sıkıştırılmış XML parçası zip bombası sayılır (gerçek belgeler ~5-30 kat). */
export const SUSPICIOUS_PART = 64 * 1024 * 1024;

interface ZipEntry {
  async(type: "uint8array"): Promise<Uint8Array>;
  name?: string;
}

/** Ortam dosyası (resim vb.) zip bombası ya da olağandışı büyük mü (açılmadan atlanır). */
export function oversizedEntry(f: unknown): boolean {
  const size = entrySize(f);
  const packed = (f as { _data?: { compressedSize?: number } } | null)?._data?.compressedSize ?? null;
  if (size === null) return false;
  return size > MAX_XML_PART || (packed !== null && size > SUSPICIOUS_PART && size > packed * 200);
}

/** Paket girdisinin açılmış boyutu (JSZip iç verisinden; bilinmiyorsa null). */
export function entrySize(f: unknown): number | null {
  const n = (f as { _data?: { uncompressedSize?: number } } | null)?._data?.uncompressedSize;
  return typeof n === "number" && Number.isFinite(n) ? n : null;
}

/**
 * Paketteki XML parçasını metne çevirir: yerel TextDecoder (JSZip'in metin çözücüsünden çok daha hızlı),
 * sondaki boşluklar atılır; olağandışı büyük parça anlaşılır hatayla reddedilir.
 */
export async function zipText(f: ZipEntry): Promise<string> {
  const size = entrySize(f);
  const packed = (f as { _data?: { compressedSize?: number } })._data?.compressedSize ?? null;
  // Zip bombası: çok büyük ve olağandışı yüksek oranda sıkıştırılmış parça açılmadan reddedilir.
  const bomb = size !== null && packed !== null && size > SUSPICIOUS_PART && size > packed * 200;
  if (size !== null && (size > MAX_XML_PART || bomb)) throw new BelgeOpenError(`Dosya açılamadı: paketteki bir bölüm olağandışı büyük (${Math.round(size / 1048576)} MB${bomb ? `, ${Math.round(size / Math.max(1, packed!))} kat sıkıştırılmış` : ""}); dosya bozuk ya da zararlı (zip bombası) olabilir.`);
  const b = await f.async("uint8array");
  let end = b.length;
  while (end > 0 && (b[end - 1] === 0x20 || b[end - 1] === 0x0a || b[end - 1] === 0x0d || b[end - 1] === 0x09 || b[end - 1] === 0)) end--;
  if (b[0] === 0xff && b[1] === 0xfe) return new TextDecoder("utf-16le").decode(b.subarray(2, end));
  if (b[0] === 0xfe && b[1] === 0xff) return new TextDecoder("utf-16be").decode(b.subarray(2, end));
  return new TextDecoder("utf-8").decode(b.subarray(0, end));
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
