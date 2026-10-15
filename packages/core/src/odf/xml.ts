// ODF XML parçacığı yardımcıları: okurken taşınacak çizim öğesinin (draw:frame) kendi ad alanı bildirimleriyle
// tek başına yazılabilir XML'i, andığı otomatik stiller; yazarken stil adlarının ve dosya başvurularının
// yeniden yazımı. Parçacıklar dize olarak işlenir (DOM gerekmez); DOM'dan okuyanlar için `elementXml` var.

export const XLINK_NS = "http://www.w3.org/1999/xlink";

/** İlk açılış etiketinin bitişi (tırnak içindeki ">" atlanır); bulunamazsa -1. */
export function openTagEnd(xml: string, from = 0): number {
  const lt = xml.indexOf("<", from);
  if (lt < 0) return -1;
  let q = "";
  for (let i = lt + 1; i < xml.length; i++) {
    const ch = xml[i];
    if (q) {
      if (ch === q) q = "";
    } else if (ch === '"' || ch === "'") q = ch;
    else if (ch === ">") return i;
  }
  return -1;
}

/** Kök (ilk öğe) etiketinin başı ve bitişi; XML bildirimi ve açıklamalar atlanır. */
function rootTag(xml: string): { start: number; end: number } | null {
  let from = 0;
  for (;;) {
    const lt = xml.indexOf("<", from);
    if (lt < 0) return null;
    if (xml.startsWith("<?", lt) || xml.startsWith("<!", lt)) {
      const e = xml.startsWith("<!--", lt) ? xml.indexOf("-->", lt) : xml.indexOf(">", lt);
      if (e < 0) return null;
      from = e + 1;
      continue;
    }
    const end = openTagEnd(xml, lt);
    return end < 0 ? null : { start: lt, end };
  }
}

/** Kök etiketin ad alanı bildirimleri (ön ek → adres; öneksiz varsayılan "" anahtarında). */
export function rootNamespaces(xml: string): Record<string, string> {
  const t = rootTag(xml);
  const out: Record<string, string> = {};
  if (!t) return out;
  for (const m of xml.slice(t.start, t.end).matchAll(/\sxmlns(?::([\w.-]+))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) out[m[1] ?? ""] = m[2] ?? m[3] ?? "";
  return out;
}

const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Parçacığın kök etiketine, içinde geçen (öğe/öznitelik adı ya da `chart:bar` gibi değer olarak) ve kökte
 * bildirilmemiş öneklerin bildirimlerini ekler. `decls`: kaynak belgenin kökündeki bildirimler.
 */
export function withNamespaces(snippet: string, decls: Record<string, string>): string {
  const t = rootTag(snippet);
  if (!t) return snippet;
  const head = snippet.slice(t.start, t.end);
  const own = new Set([...head.matchAll(/\sxmlns:([\w.-]+)\s*=/g)].map((m) => m[1]));
  let add = "";
  for (const [p, uri] of Object.entries(decls)) {
    if (!p || p === "xml" || p === "xmlns" || own.has(p)) continue;
    if (!new RegExp(`(^|[^\\w.:-])${escRe(p)}:[^\\s"'<>]`).test(snippet.replace(/\sxmlns:[\w.-]+\s*=\s*"[^"]*"/g, ""))) continue;
    add += ` xmlns:${p}="${uri.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"`;
  }
  if (!add) return snippet;
  const close = snippet[t.end - 1] === "/" ? t.end - 1 : t.end;
  return snippet.slice(0, close) + add + snippet.slice(close);
}

/** Parçacıktaki `…style-name` ve `…class-names` özniteliklerinde anılan stil adları. */
export function styleRefs(snippet: string): Set<string> {
  const out = new Set<string>();
  for (const m of snippet.matchAll(/\s[\w.-]+:([\w-]*style-name|class-names)\s*=\s*"([^"]*)"/g)) {
    for (const v of m[1] === "class-names" ? m[2].split(/\s+/) : [m[2]]) if (v) out.add(v);
  }
  return out;
}

/** Stil öğesinin adı ve ailesi (kök etiketten). */
export function styleIdentity(styleXml: string): { name: string; family: string } | null {
  const t = rootTag(styleXml);
  if (!t) return null;
  const head = styleXml.slice(t.start, t.end);
  const name = /\s[\w.-]+:name\s*=\s*"([^"]*)"/.exec(head)?.[1];
  if (!name) return null;
  return { name, family: /\s[\w.-]+:family\s*=\s*"([^"]*)"/.exec(head)?.[1] ?? "" };
}

/** Öznitelik yerel adından andığı stil ailesi (aynı adlı iki stilde seçim için). */
function familyOfAttr(prefix: string, local: string): string[] {
  if (local === "text-style-name") return ["paragraph"];
  if (prefix === "draw" && local === "style-name") return ["graphic"];
  if (prefix === "presentation" && local === "style-name") return ["presentation"];
  if (prefix === "text" && local === "style-name") return ["paragraph", "text"];
  if (prefix === "table" && local === "style-name") return ["table", "table-row", "table-column", "table-cell"];
  if (prefix === "chart" && local === "style-name") return ["chart"];
  return [];
}

/**
 * Parçacıktaki stil başvurularını yeniden adlandırır. `rename`: ad → [aile, yeni ad] listesi (aynı adlı birden çok
 * aile olabilir). Eşlenmeyen başvurular olduğu gibi kalır.
 */
export function renameStyleRefs(snippet: string, rename: Map<string, [string, string][]>): string {
  if (!rename.size) return snippet;
  const pick = (prefix: string, local: string, v: string): string => {
    const c = rename.get(v);
    if (!c?.length) return v;
    if (c.length === 1) return c[0][1];
    const fams = familyOfAttr(prefix, local);
    return (c.find(([f]) => fams.includes(f)) ?? c[0])[1];
  };
  return snippet.replace(/(\s([\w.-]+):([\w-]*style-name|class-names)\s*=\s*")([^"]*)"/g, (_m, lead: string, prefix: string, local: string, v: string) => {
    const nv = local === "class-names" ? v.split(/(\s+)/).map((x) => (/^\s*$/.test(x) ? x : pick(prefix, local, x))).join("") : pick(prefix, local, v);
    return `${lead}${nv.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"`;
  });
}

/** Stil öğesinin kendi adını değiştirir (style:name; görünen ad korunur). */
export function renameStyle(styleXml: string, name: string): string {
  const t = rootTag(styleXml);
  if (!t) return styleXml;
  const head = styleXml.slice(t.start, t.end).replace(/(\s[\w.-]+:name\s*=\s*")[^"]*"/, (_m, a: string) => `${a}${name}"`);
  return styleXml.slice(0, t.start) + head + styleXml.slice(t.end);
}

/**
 * Kök etiketin öznitelikleri: `set` verilenleri yazar ya da değiştirir, `remove` verilenleri siler (nitelikli adla:
 * "svg:x"). Değerler kaçışlanır.
 */
export function setRootAttrs(xml: string, set: Record<string, string | undefined>, remove: string[] = []): string {
  const t = rootTag(xml);
  if (!t) return xml;
  let head = xml.slice(t.start, t.end);
  const selfClose = head.endsWith("/");
  if (selfClose) head = head.slice(0, -1);
  for (const name of [...remove, ...Object.keys(set)]) head = head.replace(new RegExp(`\\s${escRe(name)}\\s*=\\s*(?:"[^"]*"|'[^']*')`, "g"), "");
  for (const [k, v] of Object.entries(set)) if (v !== undefined) head += ` ${k}="${v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;")}"`;
  return xml.slice(0, t.start) + head + (selfClose ? "/" : "") + xml.slice(t.end);
}

/** Kök etiketin öznitelik değeri (nitelikli adla). */
export function rootAttr(xml: string, name: string): string | undefined {
  const t = rootTag(xml);
  if (!t) return undefined;
  const m = new RegExp(`\\s${escRe(name)}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(xml.slice(t.start, t.end));
  return m ? (m[1] ?? m[2] ?? "").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&apos;/g, "'").replace(/&amp;/g, "&") : undefined;
}

const escText = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escValue = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;");

/**
 * DOM öğesinin tek başına yazılabilir XML'i: nitelikli adlarla yazılır, alt ağaçta ve atalarda bildirilen ad
 * alanlarından parçacıkta geçenler köke konur. Seri hâle getirici kullanılmaz: tarayıcı, jsdom ve Worker'daki
 * xmldom aynı metni üretir (açıklama ve işleme yönergeleri atılır, CDATA metin olur).
 */
export function elementXml(el: Element): string {
  const decls: Record<string, string> = {};
  const declare = (e: Element, override: boolean) => {
    for (const a of Array.from(e.attributes)) {
      const m = /^xmlns:(.+)$/.exec(a.name);
      if (m && (override || !(m[1] in decls))) decls[m[1]] = a.value;
    }
  };
  for (let p = el.parentNode; p && p.nodeType === 1; p = p.parentNode) declare(p as Element, false);
  const write = (n: Node): string => {
    if (n.nodeType === 3 || n.nodeType === 4) return escText(n.nodeValue ?? "");
    if (n.nodeType !== 1) return "";
    const e = n as Element;
    declare(e, true);
    let s = `<${e.tagName}`;
    for (const a of Array.from(e.attributes)) if (a.name !== "xmlns" && !a.name.startsWith("xmlns:")) s += ` ${a.name}="${escValue(a.value)}"`;
    let inner = "";
    for (let c = e.firstChild; c; c = c.nextSibling) inner += write(c);
    return inner ? `${s}>${inner}</${e.tagName}>` : `${s}/>`;
  };
  return withNamespaces(write(el), decls);
}

/** Kök etiketteki, parçacıkta hiç geçmeyen öneklerin (ve yinelenen) bildirimlerini atar. */
export function pruneNamespaces(xml: string): string {
  const t = rootTag(xml);
  if (!t) return xml;
  const head = xml.slice(t.start, t.end);
  const rest = xml.slice(0, t.start) + head.replace(/\sxmlns:[\w.-]+\s*=\s*(?:"[^"]*"|'[^']*')/g, "") + xml.slice(t.end);
  const seen = new Set<string>();
  const pruned = head.replace(/\sxmlns:([\w.-]+)\s*=\s*(?:"[^"]*"|'[^']*')/g, (m, p: string) => {
    if (seen.has(p)) return "";
    seen.add(p);
    return new RegExp(`(^|[^\\w.:-])${escRe(p)}:[^\\s"'<>]`).test(rest) ? m : "";
  });
  return xml.slice(0, t.start) + pruned + xml.slice(t.end);
}
