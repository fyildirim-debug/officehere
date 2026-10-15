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

/* ------------------------------------------------------------------ */
/* Yerleşik ayrıştırıcı (Worker)                                         */
/* ------------------------------------------------------------------ */

/**
 * DOMParser olmayan ortamda (Worker) düğümlerin kurulacağı boş belgeyi üreten işlev (io/xmlDom.ts kurar).
 * `force`: DOMParser olsa da yerleşik ayrıştırıcı kullanılır (Worker ile süreç içi okuyucuyu karşılaştıran testler).
 */
let builtinDom: { create: () => Document; force: boolean } | null = null;

export function setBuiltinXml(create: (() => Document) | null, opts: { force?: boolean } = {}): void {
  builtinDom = create ? { create, force: !!opts.force } : null;
}

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

class XmlSyntaxError extends Error {}

// XML 1.0 (5. baskı) ad karakterleri.
const NAME_START = ":A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}";
const NAME_RE = new RegExp(`[${NAME_START}][${NAME_START}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F\\u2040]*`, "uy");
const INVALID_CHAR = new RegExp("[^\\t\\n\\u0020-\\uD7FF\\uE000-\\uFFFD\\u{10000}-\\u{10FFFF}]", "u");
const REF_RE = /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|([^\s&;<]+))?(;?)/g;
const PREDEFINED: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };
const XML_DECL = /<\?xml[ \t\n]+version[ \t\n]*=[ \t\n]*(?:"1\.[0-9]+"|'1\.[0-9]+')(?:[ \t\n]+encoding[ \t\n]*=[ \t\n]*(?:"[A-Za-z][\w.-]*"|'[A-Za-z][\w.-]*'))?(?:[ \t\n]+standalone[ \t\n]*=[ \t\n]*(?:"(?:yes|no)"|'(?:yes|no)'))?[ \t\n]*\?>/y;

const isWs = (c: number) => c === 0x20 || c === 0x0a || c === 0x09 || c === 0x0d;
const validChar = (cp: number) => cp === 0x09 || cp === 0x0a || cp === 0x0d || (cp >= 0x20 && cp <= 0xd7ff) || (cp >= 0xe000 && cp <= 0xfffd) || (cp >= 0x10000 && cp <= 0x10ffff);

/** Ad alanı kapsamı: ön ek → ad alanı ("" varsayılan); üst kapsamdan kalıtılır. */
type Scope = Record<string, string>;
const ROOT_SCOPE: Scope = Object.assign(Object.create(null) as Scope, { xml: XML_NS, xmlns: XMLNS_NS });

/**
 * Tarayıcının XML ayrıştırıcısıyla aynı katılıkta (iyi biçimlilik ve ad alanı kuralları) ayrıştırır; düğümler
 * verilen belgede kurulur. Bozuk girdide XmlSyntaxError (tarayıcıdaki parsererror karşılığı). Tanımsız varlıklar
 * (DTD iç kümesi) desteklenmez: OOXML/ODF parçalarında DTD bulunmaz.
 */
function buildXml(source: string, doc: Document): Document {
  const s = source.includes("\r") ? source.replace(/\r\n?/g, "\n") : source;
  const fail = (msg: string, at: number): never => {
    const line = s.slice(0, at).split("\n").length;
    throw new XmlSyntaxError(`${msg} (satır ${line})`);
  };
  const bad = INVALID_CHAR.exec(s);
  if (bad) fail(`geçersiz karakter U+${bad[0].codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}`, bad.index); // yerel-ayar: ascii
  const decode = (raw: string, at: number) =>
    raw.replace(REF_RE, (m: string, hex?: string, dec?: string, name?: string, semi?: string) => {
      if (!semi || (!hex && !dec && !name)) return fail("çıplak & ya da eksik ';'", at);
      if (name) return PREDEFINED[name] ?? fail(`tanımsız varlık &${name};`, at);
      const cp = parseInt(hex ?? dec!, hex ? 16 : 10);
      return validChar(cp) ? String.fromCodePoint(cp) : fail(`geçersiz karakter başvurusu ${m}`, at);
    });
  const skipWs = (j: number) => {
    while (j < s.length && isWs(s.charCodeAt(j))) j++;
    return j;
  };
  const nameAt = (j: number): string | null => {
    NAME_RE.lastIndex = j;
    const m = NAME_RE.exec(s);
    return m ? m[0] : null;
  };
  /** QName'in ön eki ("" yok); ad alanı kurallarına uymayan ad (iki iki nokta, boş parça) hata. */
  const prefixOf = (qname: string, at: number): string => {
    const c = qname.indexOf(":");
    if (c < 0) return "";
    if (c === 0 || c === qname.length - 1 || qname.indexOf(":", c + 1) >= 0) fail(`geçersiz nitelikli ad "${qname}"`, at);
    return qname.slice(0, c);
  };

  let i = s.charCodeAt(0) === 0xfeff ? 1 : 0;
  if (s.startsWith("<?xml", i) && isWs(s.charCodeAt(i + 5))) {
    XML_DECL.lastIndex = i;
    if (!XML_DECL.test(s)) fail("geçersiz XML bildirimi", i);
    i = XML_DECL.lastIndex;
  }
  const stack: { qname: string; parent: Node; scope: Scope }[] = [];
  let parent: Node = doc;
  let scope = ROOT_SCOPE;
  let root: Element | null = null;
  let doctype = false;
  while (i < s.length) {
    const lt = s.indexOf("<", i);
    const end = lt < 0 ? s.length : lt;
    if (end > i) {
      const raw = s.slice(i, end);
      if (!stack.length) {
        for (let j = 0; j < raw.length; j++) if (!isWs(raw.charCodeAt(j))) fail("kök öğenin dışında metin", i + j);
      } else {
        if (raw.includes("]]>")) fail("metinde ']]>'", i);
        parent.appendChild(doc.createTextNode(raw.includes("&") ? decode(raw, i) : raw));
      }
      i = end;
      if (lt < 0) break;
    }
    const c = s.charCodeAt(i + 1);
    if (c === 0x21) {
      // <!-- … -->, <![CDATA[ … ]]>, <!DOCTYPE … >
      if (s.startsWith("<!--", i)) {
        const e = s.indexOf("-->", i + 4);
        if (e < 0) fail("açıklama kapanmamış", i);
        const body = s.slice(i + 4, e);
        if (body.includes("--") || body.endsWith("-")) fail("açıklamada '--'", i);
        parent.appendChild(doc.createComment(body));
        i = e + 3;
      } else if (s.startsWith("<![CDATA[", i)) {
        if (!stack.length) fail("kök öğenin dışında CDATA", i);
        const e = s.indexOf("]]>", i + 9);
        if (e < 0) fail("CDATA kapanmamış", i);
        parent.appendChild(doc.createCDATASection(s.slice(i + 9, e)));
        i = e + 3;
      } else if (s.startsWith("<!DOCTYPE", i) && !doctype && !root) {
        doctype = true;
        let j = i + 9;
        let depth = 0;
        for (; j < s.length; j++) {
          const ch = s[j];
          if (ch === '"' || ch === "'") {
            j = s.indexOf(ch, j + 1);
            if (j < 0) break;
          } else if (ch === "[") depth++;
          else if (ch === "]") depth--;
          else if (ch === ">" && depth <= 0) break;
        }
        if (j < 0 || j >= s.length) fail("DOCTYPE kapanmamış", i);
        i = j + 1;
      } else fail("geçersiz biçimlendirme", i);
      continue;
    }
    if (c === 0x3f) {
      // İşleme yönergesi
      const target = nameAt(i + 2);
      if (!target || target.includes(":") || target.toLowerCase() === "xml") fail("geçersiz işleme yönergesi", i); // yerel-ayar: ascii
      const j = i + 2 + target!.length;
      const e = s.indexOf("?>", j);
      if (e < 0) fail("işleme yönergesi kapanmamış", i);
      if (e > j && !isWs(s.charCodeAt(j))) fail("geçersiz işleme yönergesi", i);
      parent.appendChild(doc.createProcessingInstruction(target!, s.slice(skipWs(j), e)));
      i = e + 2;
      continue;
    }
    if (c === 0x2f) {
      // Kapanış etiketi
      const name = nameAt(i + 2);
      const j = name ? skipWs(i + 2 + name.length) : -1;
      const top = stack.pop();
      if (!name || s.charCodeAt(j) !== 0x3e || !top || top.qname !== name) fail(`kapanış etiketi eşleşmiyor: ${name ?? "?"}`, i);
      parent = top!.parent;
      scope = top!.scope;
      i = j + 1;
      continue;
    }
    // Açılış etiketi
    const qname = nameAt(i + 1);
    if (!qname) fail("geçersiz etiket", i);
    let j = i + 1 + qname!.length;
    const attrs: string[] = [];
    let selfClose = false;
    for (;;) {
      const k = skipWs(j);
      const ch = s.charCodeAt(k);
      if (ch === 0x3e) {
        j = k + 1;
        break;
      }
      if (ch === 0x2f && s.charCodeAt(k + 1) === 0x3e) {
        j = k + 2;
        selfClose = true;
        break;
      }
      const an = k > j ? nameAt(k) : null;
      if (!an) fail(k >= s.length ? "etiket kapanmamış" : "geçersiz öznitelik", k);
      let v = skipWs(k + an!.length);
      if (s.charCodeAt(v) !== 0x3d) fail(`öznitelikte '=' yok: ${an}`, v);
      v = skipWs(v + 1);
      const q = s[v];
      if (q !== '"' && q !== "'") fail(`öznitelik değeri tırnaksız: ${an}`, v);
      const e = s.indexOf(q, v + 1);
      if (e < 0) fail("öznitelik değeri kapanmamış", v);
      let raw = s.slice(v + 1, e);
      if (raw.includes("<")) fail("öznitelik değerinde '<'", v);
      // Değerdeki satır sonu ve sekme boşluğa döner; karakter başvuruları (&#10;) korunur.
      if (/[\t\n]/.test(raw)) raw = raw.replace(/[\t\n]/g, " ");
      attrs.push(an!, raw.includes("&") ? decode(raw, v) : raw);
      j = e + 1;
    }
    // Ad alanları
    let inner = scope;
    for (let k = 0; k < attrs.length; k += 2) {
      const an = attrs[k]!;
      if (an !== "xmlns" && !an.startsWith("xmlns:")) continue;
      const p = an === "xmlns" ? "" : an.slice(6);
      const uri = attrs[k + 1]!;
      if (p) {
        prefixOf(an, i);
        if (!uri || p === "xmlns" || (p === "xml") !== (uri === XML_NS) || uri === XMLNS_NS) fail(`geçersiz ad alanı bildirimi ${an}`, i);
      } else if (uri === XML_NS || uri === XMLNS_NS) fail("geçersiz varsayılan ad alanı", i);
      if (inner === scope) inner = Object.create(scope) as Scope;
      inner[p] = uri;
    }
    const ep = prefixOf(qname!, i);
    if (ep === "xmlns") fail("öğe adında xmlns ön eki", i);
    const ens = inner[ep] || null;
    if (ep && !ens) fail(`bildirilmemiş ad alanı ön eki: ${ep}`, i);
    const el = doc.createElementNS(ens, qname!);
    const seen = new Set<string>();
    for (let k = 0; k < attrs.length; k += 2) {
      const an = attrs[k]!;
      const p = an === "xmlns" ? "xmlns" : prefixOf(an, i);
      const ans = p ? (inner[p] ?? fail(`bildirilmemiş ad alanı ön eki: ${p}`, i)) : null;
      const key = `${ans ?? ""} ${p ? an.slice(an.indexOf(":") + 1) : an}`;
      if (seen.has(key)) fail(`yinelenen öznitelik: ${an}`, i);
      seen.add(key);
      el.setAttributeNS(ans, an, attrs[k + 1]!);
    }
    if (!stack.length) {
      if (root) fail("birden çok kök öğe", i);
      root = el;
    }
    parent.appendChild(el);
    if (!selfClose) {
      stack.push({ qname: qname!, parent, scope });
      parent = el;
      scope = inner;
    }
    i = j;
  }
  if (stack.length) fail(`kapanmamış öğe: ${stack[stack.length - 1]!.qname}`, s.length);
  if (!root) fail("kök öğe yok", s.length);
  return doc;
}

/** Yerleşik ayrıştırıcıyla (verilen DOM'da); `parseXml` ile aynı uyarlama ve hatalar. */
function builtinActive(): boolean {
  return !!builtinDom && (builtinDom.force || typeof DOMParser === "undefined");
}

function tryParse(text: string): { doc: Document; error: string | null } {
  if (builtinActive()) {
    try {
      return { doc: buildXml(text, builtinDom!.create()), error: null };
    } catch (err) {
      if (err instanceof XmlSyntaxError) return { doc: null as unknown as Document, error: err.message.slice(0, 160) };
      throw err;
    }
  }
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
