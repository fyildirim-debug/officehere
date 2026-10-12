/** Sayı biçimi kodu ayrıştırıcısı: bölümler, belirteçler, renkler, koşullar. */
export const COLOR_NAMES: Record<string, string> = {
  black: "#000000",
  blue: "#0000FF",
  cyan: "#00FFFF",
  green: "#00FF00",
  magenta: "#FF00FF",
  red: "#FF0000",
  white: "#FFFFFF",
  yellow: "#FFFF00",
  siyah: "#000000",
  mavi: "#0000FF",
  camgöbeği: "#00FFFF",
  yeşil: "#00FF00",
  macenta: "#FF00FF",
  kırmızı: "#FF0000",
  beyaz: "#FFFFFF",
  sarı: "#FFFF00",
};

export const COLOR_TR: Record<string, string> = {
  Black: "Siyah",
  Blue: "Mavi",
  Cyan: "Camgöbeği",
  Green: "Yeşil",
  Magenta: "Macenta",
  Red: "Kırmızı",
  White: "Beyaz",
  Yellow: "Sarı",
};
export const COLOR_EN: Record<string, string> = Object.fromEntries(Object.entries(COLOR_TR).map(([k, v]) => [v.toLocaleLowerCase("tr-TR"), k]));

export type Tok =
  | { t: "lit"; v: string }
  | { t: "digit"; v: "0" | "#" | "?" }
  | { t: "point" }
  | { t: "comma" }
  | { t: "percent" }
  | { t: "exp"; sign: "+" | "-" }
  | { t: "slash" }
  | { t: "text" }
  | { t: "date"; v: string }
  | { t: "ampm"; v: string }
  | { t: "elapsed"; v: "h" | "m" | "s"; n: number }
  | { t: "fracsec"; n: number }
  | { t: "general" }
  | { t: "fill"; v: string };

export interface Section {
  toks: Tok[];
  color?: string;
  cond?: { op: string; v: number };
  kind: "number" | "date" | "text" | "general";
  hasAmPm: boolean;
  /** Yerel ayar etiketi ([$-tr-TR], [$-409]…): ay/gün adlarının dili. */
  lang?: "tr" | "en";
  /** tr/en dışındaki yerel ayar ([$-de-DE], [$-407]): adlar Intl ile üretilir. */
  locale?: string;
}

/** Sık kullanılan Windows yerel ayar kimlikleri (onaltılık LCID → BCP 47). */
const LCID: Record<string, string> = {
  "407": "de-DE", "807": "de-CH", "c07": "de-AT", "40c": "fr-FR", "80c": "fr-BE", "c0c": "fr-CA", "100c": "fr-CH", "40a": "es-ES", "c0a": "es-ES", "80a": "es-MX",
  "410": "it-IT", "413": "nl-NL", "813": "nl-BE", "416": "pt-BR", "816": "pt-PT", "419": "ru-RU", "422": "uk-UA", "415": "pl-PL", "405": "cs-CZ", "40e": "hu-HU",
  "418": "ro-RO", "402": "bg-BG", "41a": "hr-HR", "408": "el-GR", "41d": "sv-SE", "406": "da-DK", "414": "nb-NO", "40b": "fi-FI", "401": "ar-SA", "40d": "he-IL",
  "429": "fa-IR", "411": "ja-JP", "412": "ko-KR", "804": "zh-CN", "404": "zh-TW", "42c": "az-Latn-AZ", "43f": "kk-KZ", "443": "uz-Latn-UZ",
};

/** Etiketteki yerel ayarın BCP 47 adı (tr/en dışı; tanınmazsa undefined). */
export function localeTagName(tag: string): string | undefined {
  const t = tag.trim();
  const hex = t.toLowerCase().replace(/^0+/, "");
  if (/^[0-9a-f]{3,8}$/.test(hex)) return LCID[hex.length > 4 ? hex.slice(-4).replace(/^0+/, "") : hex];
  if (/^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/i.test(t)) {
    try {
      return Intl.getCanonicalLocales(t)[0];
    } catch {
      return undefined;
    }
  }
  return undefined;
}

/** [$sembol-yerel] etiketindeki yerel ayarın dili. */
export function localeTagLang(lcid: string): "tr" | "en" | undefined {
  const l = lcid.trim().toLowerCase();
  if (/^(41f|tr)/.test(l)) return "tr";
  if (/^(409|809|c09|1009|1409|1809|2009|4009|en)/.test(l) || /^(f800|f400)$/.test(l)) return "en";
  return undefined;
}

export interface Parsed {
  sections: Section[];
}

const cache = new Map<string, Parsed>();

export function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = "";
  let inQ = false;
  let inB = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"' && !inB) inQ = !inQ;
    else if (c === "[" && !inQ) inB = true;
    else if (c === "]" && !inQ) inB = false;
    else if (c === "\\" && !inQ && !inB) {
      cur += c + (code[i + 1] ?? "");
      i++;
      continue;
    }
    if (c === ";" && !inQ && !inB) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  out.push(cur);
  return out;
}

export function parseSection(src: string): Section {
  const toks: Tok[] = [];
  const sec: Section = { toks, kind: "number", hasAmPm: false };
  let i = 0;
  const lower = src.toLowerCase();
  while (i < src.length) {
    const c = src[i];
    const lc = lower[i];
    if (c === '"') {
      let j = i + 1;
      let s = "";
      while (j < src.length && src[j] !== '"') s += src[j++];
      toks.push({ t: "lit", v: s });
      i = j + 1;
      continue;
    }
    if (c === "\\") {
      toks.push({ t: "lit", v: src[i + 1] ?? "" });
      i += 2;
      continue;
    }
    if (c === "_") {
      toks.push({ t: "lit", v: " " });
      i += 2;
      continue;
    }
    if (c === "*") {
      toks.push({ t: "fill", v: src[i + 1] ?? " " });
      i += 2;
      continue;
    }
    if (c === "[") {
      const j = src.indexOf("]", i);
      const inner = j < 0 ? src.slice(i + 1) : src.slice(i + 1, j);
      i = j < 0 ? src.length : j + 1;
      const il = inner.toLocaleLowerCase("tr-TR");
      if (COLOR_NAMES[il]) sec.color = COLOR_NAMES[il];
      else if (/^color\s*\d+$/i.test(inner)) sec.color = paletteColor(parseInt(inner.replace(/\D/g, ""), 10));
      else if (/^(<=|>=|<>|<|>|=)\s*-?[\d.]+$/.test(inner)) {
        const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner)!;
        sec.cond = { op: m[1], v: parseFloat(m[2]) };
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        toks.push({ t: "elapsed", v: inner[0].toLowerCase() as "h" | "m" | "s", n: inner.length });
      } else if (inner.startsWith("$")) {
        const body = inner.slice(1);
        const dash = body.indexOf("-");
        const sym = dash < 0 ? body : body.slice(0, dash);
        if (sym) toks.push({ t: "lit", v: sym });
        if (dash >= 0) {
          const tag = body.slice(dash + 1);
          const l = localeTagLang(tag);
          if (l) {
            sec.lang = l;
            sec.locale = undefined;
          } else sec.locale = localeTagName(tag) ?? sec.locale;
        }
      }
      continue;
    }
    if (lower.startsWith("general", i)) {
      toks.push({ t: "general" });
      i += 7;
      continue;
    }
    if (lower.startsWith("am/pm", i)) {
      toks.push({ t: "ampm", v: "AM/PM" });
      sec.hasAmPm = true;
      i += 5;
      continue;
    }
    if (lower.startsWith("a/p", i)) {
      toks.push({ t: "ampm", v: "A/P" });
      sec.hasAmPm = true;
      i += 3;
      continue;
    }
    if (c === "0" || c === "#" || c === "?") {
      toks.push({ t: "digit", v: c });
      i++;
      continue;
    }
    if (c === ".") {
      // Saniye kesirleri: ss.00
      const prev = toks[toks.length - 1];
      if (prev && (prev.t === "date" && prev.v[0] === "s" || prev.t === "elapsed" && prev.v === "s") && src[i + 1] === "0") {
        let n = 0;
        let j = i + 1;
        while (src[j] === "0") {
          n++;
          j++;
        }
        toks.push({ t: "fracsec", n });
        i = j;
        continue;
      }
      toks.push({ t: "point" });
      i++;
      continue;
    }
    if (c === ",") {
      toks.push({ t: "comma" });
      i++;
      continue;
    }
    if (c === "%") {
      toks.push({ t: "percent" });
      i++;
      continue;
    }
    if ((lc === "e") && (src[i + 1] === "+" || src[i + 1] === "-")) {
      toks.push({ t: "exp", sign: src[i + 1] as "+" | "-" });
      i += 2;
      continue;
    }
    if (c === "/") {
      toks.push({ t: "slash" });
      i++;
      continue;
    }
    if (c === "@") {
      toks.push({ t: "text" });
      i++;
      continue;
    }
    if ("ymdhs".includes(lc)) {
      let j = i;
      while (j < src.length && lower[j] === lc) j++;
      toks.push({ t: "date", v: lc.repeat(j - i) });
      i = j;
      continue;
    }
    if (lc === "b" && /^b{2,4}/.test(lower.slice(i))) {
      // Budist takvimi yılı: yıl olarak kabul et.
      let j = i;
      while (lower[j] === "b") j++;
      toks.push({ t: "date", v: j - i > 2 ? "yyyy" : "yy" });
      i = j;
      continue;
    }
    toks.push({ t: "lit", v: c });
    i++;
  }
  if (toks.some((t) => t.t === "date" || t.t === "elapsed" || t.t === "ampm")) sec.kind = "date";
  else if (toks.some((t) => t.t === "text")) sec.kind = "text";
  else if (toks.some((t) => t.t === "general")) sec.kind = "general";
  // Dakika/ay ayrımı: h'den sonra ya da s'den önce gelen m dakikadır.
  if (sec.kind === "date") {
    const dateIdx = toks.map((t, k) => (t.t === "date" || t.t === "elapsed" ? k : -1)).filter((k) => k >= 0);
    for (let n = 0; n < dateIdx.length; n++) {
      const tk = toks[dateIdx[n]] as { t: "date"; v: string };
      if (tk.t !== "date" || tk.v[0] !== "m" || tk.v.length > 2) continue;
      const prev = n > 0 ? toks[dateIdx[n - 1]] : undefined;
      const next = n < dateIdx.length - 1 ? toks[dateIdx[n + 1]] : undefined;
      const prevH = prev && ((prev.t === "date" && prev.v[0] === "h") || (prev.t === "elapsed" && prev.v === "h"));
      const nextS = next && ((next.t === "date" && next.v[0] === "s") || (next.t === "elapsed" && next.v === "s"));
      if (prevH || nextS) tk.v = tk.v.length === 1 ? "M" : "MM";
    }
  }
  return sec;
}

export function paletteColor(n: number): string {
  const base = ["#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080"];
  return base[(n - 1) % base.length] ?? "#000000";
}

export function getParsed(code: string): Parsed {
  let p = cache.get(code);
  if (!p) {
    const parts = splitSections(code || "General");
    p = { sections: parts.map(parseSection) };
    if (cache.size > 2000) cache.clear();
    cache.set(code, p);
  }
  return p;
}

/** Biçim bir tarih/saat biçimi mi? */
export function isDateFormat(code: string | undefined): boolean {
  if (!code) return false;
  const p = getParsed(code);
  return p.sections[0]?.kind === "date";
}

/** Biçim yalnız saat mi (tarih bileşeni yok)? */
export function isTimeOnlyFormat(code: string | undefined): boolean {
  if (!code || !isDateFormat(code)) return false;
  const s = getParsed(code).sections[0];
  return !s.toks.some((t) => t.t === "date" && /^[ydm]/.test(t.v));
}

export function isTextFormat(code: string | undefined): boolean {
  return code === "@";
}

/** Biçimdeki yüzde işareti (giriş ayrıştırma için). */
export function isPercentFormat(code: string | undefined): boolean {
  if (!code) return false;
  return getParsed(code).sections[0]?.toks.some((t) => t.t === "percent") ?? false;
}
