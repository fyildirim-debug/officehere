/**
 * Küçük, DOM'suz XML okuyucusu (ODF parçaları için). Tarayıcıda ve Node'da
 * aynı çalışır; büyük content.xml dosyalarında ağaç kurmadan olay akışıyla
 * da kullanılabilir. Ad alanı önekleri olduğu gibi korunur ("table:table").
 */

export interface XmlEl {
  name: string;
  attrs: Record<string, string>;
  children: (XmlEl | string)[];
}

export interface XmlHandlers {
  open(name: string, attrs: Record<string, string>, selfClosing: boolean): void;
  close(name: string): void;
  text(text: string): void;
}

const ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

/** XML varlıklarını çözer (&amp; &#x41; …). */
export function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e: string) => {
    if (e[0] === "#") {
      const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(n) && n >= 0 && n <= 0x10ffff ? String.fromCodePoint(n) : "";
    }
    return ENTITIES[e] ?? m;
  });
}

/** Metin ve öznitelik değerleri için kaçış. */
export function esc(s: string): string {
  return s.replace(/[&<>"]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;")).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "");
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseAttrs(src: string): Record<string, string> {
  const out: Record<string, string> = {};
  if (!src) return out;
  ATTR_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ATTR_RE.exec(src))) out[m[1]] = decodeEntities(m[3] ?? m[4] ?? "");
  return out;
}

/** Olay akışıyla ayrıştırır; bozuk işaretlemede Error fırlatır. */
export function scanXml(xml: string, h: XmlHandlers): void {
  let i = 0;
  const n = xml.length;
  const stack: string[] = [];
  while (i < n) {
    const lt = xml.indexOf("<", i);
    if (lt < 0) {
      if (stack.length) h.text(decodeEntities(xml.slice(i)));
      break;
    }
    if (lt > i && stack.length) h.text(decodeEntities(xml.slice(i, lt)));
    if (xml.startsWith("<!--", lt)) {
      const e = xml.indexOf("-->", lt + 4);
      if (e < 0) throw new Error("XML açıklaması kapanmamış");
      i = e + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const e = xml.indexOf("]]>", lt + 9);
      if (e < 0) throw new Error("CDATA kapanmamış");
      if (stack.length) h.text(xml.slice(lt + 9, e));
      i = e + 3;
      continue;
    }
    if (xml.startsWith("<?", lt) || xml.startsWith("<!", lt)) {
      const e = xml.indexOf(">", lt + 2);
      if (e < 0) throw new Error("XML bildirimi kapanmamış");
      i = e + 1;
      continue;
    }
    // Etiketin sonu (tırnak içindeki ">" atlanır)
    let j = lt + 1;
    let q = 0;
    for (; j < n; j++) {
      const c = xml.charCodeAt(j);
      if (q) {
        if (c === q) q = 0;
      } else if (c === 34 || c === 39) q = c;
      else if (c === 62) break;
    }
    if (j >= n) throw new Error("XML etiketi kapanmamış");
    const body = xml.slice(lt + 1, j);
    i = j + 1;
    if (body[0] === "/") {
      const name = body.slice(1).trim();
      const top = stack.pop();
      if (top !== name) throw new Error(`XML yapısı bozuk: </${name}>`);
      h.close(name);
      continue;
    }
    const self = body.endsWith("/");
    const inner = self ? body.slice(0, -1) : body;
    const sp = inner.search(/\s/);
    const name = sp < 0 ? inner : inner.slice(0, sp);
    const attrs = parseAttrs(sp < 0 ? "" : inner.slice(sp + 1));
    h.open(name, attrs, self);
    if (self) h.close(name);
    else stack.push(name);
  }
  if (stack.length) throw new Error("XML belgesi eksik (kapanmamış öğeler)");
}

/** Bütün belgeyi ağaca çevirir (küçük parçalar için). */
export function parseXml(xml: string): XmlEl {
  const root: XmlEl = { name: "#document", attrs: {}, children: [] };
  const stack: XmlEl[] = [root];
  scanXml(xml, {
    open(name, attrs) {
      const el: XmlEl = { name, attrs, children: [] };
      stack[stack.length - 1].children.push(el);
      stack.push(el);
    },
    close() {
      stack.pop();
    },
    text(t) {
      stack[stack.length - 1].children.push(t);
    },
  });
  return root;
}

/** Olay akışından alt ağaç kurmak için yardımcı (açılış olayından başlatılır). */
export class TreeBuilder {
  private stack: XmlEl[];
  readonly root: XmlEl;
  constructor(name: string, attrs: Record<string, string>) {
    this.root = { name, attrs, children: [] };
    this.stack = [this.root];
  }
  open(name: string, attrs: Record<string, string>): void {
    const el: XmlEl = { name, attrs, children: [] };
    this.stack[this.stack.length - 1].children.push(el);
    this.stack.push(el);
  }
  /** Kökün kapanışında true döner. */
  close(): boolean {
    this.stack.pop();
    return this.stack.length === 0;
  }
  text(t: string): void {
    this.stack[this.stack.length - 1]?.children.push(t);
  }
}

export function kids(el: XmlEl | undefined, name?: string): XmlEl[] {
  if (!el) return [];
  return el.children.filter((c): c is XmlEl => typeof c !== "string" && (!name || c.name === name));
}

export function child(el: XmlEl | undefined, name: string): XmlEl | undefined {
  return el?.children.find((c): c is XmlEl => typeof c !== "string" && c.name === name);
}

/** Derinlemesine ilk eşleşen öğe. */
export function find(el: XmlEl | undefined, name: string): XmlEl | undefined {
  if (!el) return undefined;
  for (const c of el.children) {
    if (typeof c === "string") continue;
    if (c.name === name) return c;
    const f = find(c, name);
    if (f) return f;
  }
  return undefined;
}

/** Öğenin düz metni (alt öğeler dahil). */
export function textOf(el: XmlEl): string {
  return el.children.map((c) => (typeof c === "string" ? c : textOf(c))).join("");
}
