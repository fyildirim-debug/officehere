// Markdown satır içi ayrıştırıcısı (CommonMark çekirdeği + GFM üstü çizili ve otomatik bağlantılar):
// vurgu ayraç yığını, bağlantı/resim köşeli ayraçları, kod parçaları, kaçışlar, varlıklar, satır içi HTML.

export type Inl =
  | { k: "text"; s: string }
  | { k: "code"; s: string }
  | { k: "br" }
  | { k: "soft" }
  | { k: "em" | "strong" | "del"; c: Inl[] }
  | { k: "link"; href: string; c: Inl[] }
  | { k: "img"; src: string; alt: string }
  | { k: "html"; tag: string; close: boolean }
  | { k: "fn"; label: string };

export interface InlineCtx {
  refs: Map<string, { href: string; title?: string }>;
  notes: Set<string>;
}

export const normLabel = (s: string) => s.trim().replace(/\s+/g, " ").toLocaleLowerCase("tr-TR");

const PUNCT = /[!-/:-@[-`{-~\p{P}\p{S}]/u;
const isWs = (c: string | undefined) => c === undefined || /\s/u.test(c);
const isPunct = (c: string | undefined) => !!c && PUNCT.test(c);

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
  copy: "©",
  reg: "®",
  trade: "™",
  hellip: "…",
  mdash: "—",
  ndash: "–",
  laquo: "«",
  raquo: "»",
  ldquo: "“",
  rdquo: "”",
  lsquo: "‘",
  rsquo: "’",
  euro: "€",
  deg: "°",
  times: "×",
  divide: "÷",
  middot: "·",
  bull: "•",
  sect: "§",
  para: "¶",
  shy: "­",
};

/** "&amp;" / "&#305;" / "&#x131;" çözümü; tanınmayan varlık null. */
export function decodeEntity(s: string): { text: string; len: number } | null {
  const m = /^&(?:#(\d{1,7})|#[xX]([0-9a-fA-F]{1,6})|([a-zA-Z][a-zA-Z0-9]{1,31}));/.exec(s);
  if (!m) return null;
  if (m[1] || m[2]) {
    const cp = m[1] ? parseInt(m[1], 10) : parseInt(m[2], 16);
    return { text: cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�", len: m[0].length };
  }
  const t = ENTITIES[m[3]];
  return t === undefined ? null : { text: t, len: m[0].length };
}

/** Kaçışları ve varlıkları çözer (bağlantı adresi, başlık, kod bilgi dizgesi). */
export function unescapeMd(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === "\\" && i + 1 < s.length && /[!-/:-@[-`{-~]/.test(s[i + 1])) {
      out += s[++i];
    } else if (c === "&") {
      const e = decodeEntity(s.slice(i));
      if (e) {
        out += e.text;
        i += e.len - 1;
      } else out += c;
    } else out += c;
  }
  return out;
}

interface Delim {
  node: { k: "text"; s: string };
  ch: string;
  n: number;
  orig: number;
  open: boolean;
  close: boolean;
}

interface Bracket {
  node: { k: "text"; s: string };
  image: boolean;
  active: boolean;
  delimBottom: number;
  /** Köşeli ayracın kaynaktaki konumu (etiket metni için). */
  pos: number;
}

/** Bağlantı hedefi ve isteğe bağlı başlık: "(adres "başlık")". */
function linkTail(s: string, at: number): { href: string; end: number } | null {
  let i = at + 1;
  const skipWs = () => {
    while (i < s.length && /[ \t\n]/.test(s[i])) i++;
  };
  skipWs();
  let href = "";
  if (s[i] === "<") {
    const close = s.indexOf(">", i);
    if (close < 0 || /[\n<]/.test(s.slice(i + 1, close))) return null;
    href = s.slice(i + 1, close);
    i = close + 1;
  } else {
    let depth = 0;
    const start = i;
    while (i < s.length) {
      const c = s[i];
      if (c === "\\" && i + 1 < s.length) {
        i += 2;
        continue;
      }
      if (/[\s\x00-\x1f]/.test(c)) break;
      if (c === "(") depth++;
      else if (c === ")") {
        if (depth === 0) break;
        depth--;
      }
      i++;
    }
    href = s.slice(start, i);
  }
  skipWs();
  if (s[i] === '"' || s[i] === "'" || s[i] === "(") {
    const closeCh = s[i] === "(" ? ")" : s[i];
    let j = i + 1;
    while (j < s.length && s[j] !== closeCh) j += s[j] === "\\" ? 2 : 1;
    if (j >= s.length) return null;
    i = j + 1;
    skipWs();
  }
  if (s[i] !== ")") return null;
  return { href: unescapeMd(href), end: i + 1 };
}

const INLINE_TAGS = /^<(\/?)([a-zA-Z][a-zA-Z0-9-]*)((?:\s+[a-zA-Z_:][\w.:-]*(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s"'=<>`]+))?)*)\s*(\/?)>/;
const AUTOLINK = /^<([a-zA-Z][a-zA-Z0-9+.-]{1,31}:[^\s<>]*)>/;
const EMAIL = /^<([a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)>/;
const BARE_URL = /^(?:https?:\/\/|www\.)[^\s<]*[^\s<?!.,:*_~'")\]]/;

/** Paragraf metnini satır içi düğümlere ayrıştırır. */
export function parseInline(src: string, ctx: InlineCtx): Inl[] {
  const nodes: Inl[] = [];
  const delims: Delim[] = [];
  const brackets: Bracket[] = [];
  let text = "";
  const flush = () => {
    if (text) nodes.push({ k: "text", s: text });
    text = "";
  };
  let i = 0;
  const s = src;
  while (i < s.length) {
    const c = s[i];
    // Kaçış ve ters bölüyle satır sonu
    if (c === "\\") {
      const nx = s[i + 1];
      if (nx === "\n") {
        flush();
        nodes.push({ k: "br" });
        i += 2;
        while (s[i] === " ") i++;
        continue;
      }
      if (nx !== undefined && /[!-/:-@[-`{-~]/.test(nx)) {
        text += nx;
        i += 2;
        continue;
      }
      text += c;
      i++;
      continue;
    }
    if (c === "\n") {
      // İki ya da daha çok boşlukla biten satır: zorunlu satır sonu.
      const hard = / {2,}$/.test(text);
      text = text.replace(/ +$/, "");
      flush();
      nodes.push(hard ? { k: "br" } : { k: "soft" });
      i++;
      while (s[i] === " ") i++;
      continue;
    }
    if (c === "`") {
      let n = 0;
      while (s[i + n] === "`") n++;
      const fence = "`".repeat(n);
      let j = i + n;
      let found = -1;
      while (j < s.length) {
        const k = s.indexOf(fence, j);
        if (k < 0) break;
        let m = 0;
        while (s[k + m] === "`") m++;
        if (m === n) {
          found = k;
          break;
        }
        j = k + m;
      }
      if (found < 0) {
        text += fence;
        i += n;
        continue;
      }
      let code = s.slice(i + n, found).replace(/\n/g, " ");
      if (code.length > 2 && code.startsWith(" ") && code.endsWith(" ") && code.trim()) code = code.slice(1, -1);
      flush();
      nodes.push({ k: "code", s: code });
      i = found + n;
      continue;
    }
    if (c === "&") {
      const e = decodeEntity(s.slice(i));
      if (e) {
        text += e.text;
        i += e.len;
        continue;
      }
    }
    if (c === "<") {
      const rest = s.slice(i);
      let m = AUTOLINK.exec(rest);
      if (m) {
        flush();
        nodes.push({ k: "link", href: m[1], c: [{ k: "text", s: m[1] }] });
        i += m[0].length;
        continue;
      }
      m = EMAIL.exec(rest);
      if (m) {
        flush();
        nodes.push({ k: "link", href: `mailto:${m[1]}`, c: [{ k: "text", s: m[1] }] });
        i += m[0].length;
        continue;
      }
      if (rest.startsWith("<!--")) {
        const end = rest.indexOf("-->", 4);
        if (end >= 0) {
          i += end + 3;
          continue;
        }
      }
      m = INLINE_TAGS.exec(rest);
      if (m) {
        flush();
        nodes.push({ k: "html", tag: m[2].toLowerCase(), close: m[1] === "/" });
        i += m[0].length;
        continue;
      }
    }
    // GFM genişletilmiş otomatik bağlantı (sözcük başında).
    if ((c === "h" || c === "w") && (i === 0 || /[\s*_~(]/.test(s[i - 1]))) {
      const m = BARE_URL.exec(s.slice(i));
      if (m && !brackets.some((b) => b.active && !b.image)) {
        let url = m[0];
        // Dengesiz kapanış parantezi bağlantıya dahil edilmez.
        while (url.endsWith(")") && (url.match(/\(/g)?.length ?? 0) < (url.match(/\)/g)?.length ?? 0)) url = url.slice(0, -1);
        flush();
        nodes.push({ k: "link", href: url.startsWith("www.") ? `http://${url}` : url, c: [{ k: "text", s: url }] });
        i += url.length;
        continue;
      }
    }
    if (c === "*" || c === "_" || c === "~") {
      let n = 0;
      while (s[i + n] === c) n++;
      if (c === "~" && n > 2) {
        text += s.slice(i, i + n);
        i += n;
        continue;
      }
      const before = i > 0 ? s[i - 1] : undefined;
      const after = s[i + n];
      const left = !isWs(after) && (!isPunct(after) || isWs(before) || isPunct(before));
      const right = !isWs(before) && (!isPunct(before) || isWs(after) || isPunct(after));
      let open = left;
      let close = right;
      if (c === "_") {
        open = left && (!right || isPunct(before));
        close = right && (!left || isPunct(after));
      }
      flush();
      const node = { k: "text" as const, s: s.slice(i, i + n) };
      nodes.push(node);
      if (open || close) delims.push({ node, ch: c, n, orig: n, open, close });
      i += n;
      continue;
    }
    if (c === "[" && s[i + 1] === "^") {
      const m = /^\[\^([^\]\s]+)\]/.exec(s.slice(i));
      if (m && ctx.notes.has(normLabel(m[1])) && s[i + m[0].length] !== ":") {
        flush();
        nodes.push({ k: "fn", label: normLabel(m[1]) });
        i += m[0].length;
        continue;
      }
    }
    if (c === "[" || (c === "!" && s[i + 1] === "[")) {
      flush();
      const image = c === "!";
      const node = { k: "text" as const, s: image ? "![" : "[" };
      nodes.push(node);
      brackets.push({ node, image, active: true, delimBottom: delims.length, pos: i + (image ? 2 : 1) });
      i += image ? 2 : 1;
      continue;
    }
    if (c === "]") {
      flush();
      const b = brackets.pop();
      if (!b) {
        text += "]";
        i++;
        continue;
      }
      if (!b.active) {
        text += "]";
        i++;
        continue;
      }
      const labelText = s.slice(b.pos, i);
      let href: string | null = null;
      let end = i + 1;
      if (s[i + 1] === "(") {
        const t = linkTail(s, i + 1);
        if (t) {
          href = t.href;
          end = t.end;
        }
      }
      if (href === null) {
        // Başvuru bağlantıları: [metin][etiket], [metin][], [metin]
        const m = /^\[([^\]]*)\]/.exec(s.slice(i + 1));
        const label = m && m[1].trim() ? m[1] : labelText;
        const ref = ctx.refs.get(normLabel(label));
        if (ref) {
          href = ref.href;
          end = m ? i + 1 + m[0].length : i + 1;
        }
      }
      if (href === null) {
        text += "]";
        i++;
        continue;
      }
      processEmphasis(nodes, delims, b.delimBottom);
      const at = nodes.indexOf(b.node);
      const inner = nodes.splice(at + 1);
      nodes.pop();
      if (b.image) nodes.push({ k: "img", src: href, alt: plainOf(inner) });
      else {
        nodes.push({ k: "link", href, c: inner });
        // Bağlantı bağlantı içeremez: önceki açık köşeli ayraçlar etkisizleşir.
        for (const o of brackets) if (!o.image) o.active = false;
      }
      i = end;
      continue;
    }
    text += c;
    i++;
  }
  flush();
  processEmphasis(nodes, delims, 0);
  return nodes;
}

/** CommonMark "vurguyu işle" algoritması (üçün katı kuralıyla). */
function processEmphasis(nodes: Inl[], delims: Delim[], bottom: number) {
  let ci = bottom;
  while (ci < delims.length) {
    const closer = delims[ci];
    if (!closer.close) {
      ci++;
      continue;
    }
    let oi = -1;
    for (let j = ci - 1; j >= bottom; j--) {
      const o = delims[j];
      if (o.ch !== closer.ch || !o.open) continue;
      if (o.ch === "~") {
        if (o.n === closer.n) {
          oi = j;
          break;
        }
        continue;
      }
      const odd = (o.close || closer.open) && (o.orig + closer.orig) % 3 === 0 && !(o.orig % 3 === 0 && closer.orig % 3 === 0);
      if (!odd) {
        oi = j;
        break;
      }
    }
    if (oi < 0) {
      if (!closer.open) delims.splice(ci, 1);
      else ci++;
      continue;
    }
    const o = delims[oi];
    const use = o.ch === "~" ? o.n : closer.n >= 2 && o.n >= 2 ? 2 : 1;
    const a = nodes.indexOf(o.node);
    const b = nodes.indexOf(closer.node);
    const inner = nodes.slice(a + 1, b);
    const wrap: Inl = { k: o.ch === "~" ? "del" : use === 2 ? "strong" : "em", c: inner };
    nodes.splice(a + 1, b - a - 1, wrap);
    o.n -= use;
    o.node.s = o.node.s.slice(0, o.n);
    closer.n -= use;
    closer.node.s = closer.node.s.slice(0, closer.n);
    delims.splice(oi + 1, ci - oi - 1);
    ci = oi + 1;
    if (o.n === 0) {
      nodes.splice(nodes.indexOf(o.node), 1);
      delims.splice(oi, 1);
      ci--;
    }
    if (closer.n === 0) {
      nodes.splice(nodes.indexOf(closer.node), 1);
      delims.splice(ci, 1);
    }
  }
  delims.length = bottom;
}

/** Düğümlerin düz metni (resim alternatif metni için). */
export function plainOf(nodes: Inl[]): string {
  return nodes
    .map((n) => {
      switch (n.k) {
        case "text":
        case "code":
          return n.s;
        case "soft":
        case "br":
          return " ";
        case "img":
          return n.alt;
        case "em":
        case "strong":
        case "del":
        case "link":
          return plainOf(n.c);
        default:
          return "";
      }
    })
    .join("");
}
