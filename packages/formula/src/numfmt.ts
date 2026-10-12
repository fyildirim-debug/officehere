/**
 * Sayı biçimi kodları (ECMA-376 §18.8.31 numFmt) için biçimlendirici.
 * Kodlar kanonik (İngilizce) tutulur; çıktı tr-TR ayırıcılarıyla üretilir.
 * Türkçe kodlar ("#.##0,00", "gg.aa.yyyy", "[Kırmızı]") formatFromLocal ile çevrilir.
 */
import { DAYS_EN, DAYS_TR, DAYS_TR_SHORT, MONTHS_EN, MONTHS_TR, MONTHS_TR_SHORT, serialToDate, serialToTime, weekdayOf } from "./dates";

export interface FormatOptions {
  decimal?: string;
  thousands?: string;
  /** Ay/gün adları dili. */
  lang?: "tr" | "en";
}

export interface FormatResult {
  text: string;
  color?: string;
  /** "*x" dolgu: hücre genişliğine kadar bu konumda tekrarlanacak karakter. */
  fill?: { index: number; char: string };
}

const TR_OPTS: Required<FormatOptions> = { decimal: ",", thousands: ".", lang: "tr" };

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

const COLOR_TR: Record<string, string> = {
  Black: "Siyah",
  Blue: "Mavi",
  Cyan: "Camgöbeği",
  Green: "Yeşil",
  Magenta: "Macenta",
  Red: "Kırmızı",
  White: "Beyaz",
  Yellow: "Sarı",
};
const COLOR_EN: Record<string, string> = Object.fromEntries(Object.entries(COLOR_TR).map(([k, v]) => [v.toLocaleLowerCase("tr-TR"), k]));

type Tok =
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

interface Section {
  toks: Tok[];
  color?: string;
  cond?: { op: string; v: number };
  kind: "number" | "date" | "text" | "general";
  hasAmPm: boolean;
}

interface Parsed {
  sections: Section[];
}

const cache = new Map<string, Parsed>();

function splitSections(code: string): string[] {
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

function parseSection(src: string): Section {
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
        const sym = inner.slice(1).split("-")[0];
        if (sym) toks.push({ t: "lit", v: sym });
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

function paletteColor(n: number): string {
  const base = ["#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080"];
  return base[(n - 1) % base.length] ?? "#000000";
}

function getParsed(code: string): Parsed {
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

function condMatch(c: { op: string; v: number }, x: number): boolean {
  switch (c.op) {
    case "<":
      return x < c.v;
    case "<=":
      return x <= c.v;
    case ">":
      return x > c.v;
    case ">=":
      return x >= c.v;
    case "=":
      return x === c.v;
    case "<>":
      return x !== c.v;
  }
  return false;
}

/** General biçimi: en çok 11 karakter. */
export function formatGeneral(v: number, decimal = ","): string {
  if (!Number.isFinite(v)) return "#NUM!";
  if (v === 0) return "0";
  const abs = Math.abs(v);
  if (Number.isInteger(v) && abs < 1e11) return String(v);
  if (abs >= 1e11 || abs < 1e-9) {
    for (let p = 6; p >= 1; p--) {
      const s = expString(v, p - 1, decimal);
      if (s.length <= 11) return s;
    }
    return expString(v, 0, decimal);
  }
  for (let p = 10; p >= 1; p--) {
    let s = Number(v.toPrecision(p)).toString();
    if (s.includes("e")) s = Number(v.toPrecision(p)).toFixed(Math.max(0, 10)).replace(/0+$/, "").replace(/\.$/, "");
    if (s.length <= 11 || p === 1) return s.replace(".", decimal);
  }
  return String(v).replace(".", decimal);
}

function expString(v: number, decimals: number, decimal: string): string {
  const s = v.toExponential(decimals); // "1.23e+11"
  const m = /^(-?)(\d)(?:\.(\d+))?e([+-])(\d+)$/.exec(s);
  if (!m) return s;
  let mant = m[3] ? `${m[2]}.${m[3]}`.replace(/0+$/, "").replace(/\.$/, "") : m[2];
  mant = mant.replace(".", decimal);
  return `${m[1]}${mant}E${m[4]}${m[5].padStart(2, "0")}`;
}

/** Formüllerde sayı → metin dönüşümü (15 anlamlı basamak). */
export function numberToText(v: number, decimal = ","): string {
  if (!Number.isFinite(v)) return "#NUM!";
  if (v === 0) return "0";
  const abs = Math.abs(v);
  let s: string;
  if (abs >= 1e21 || abs < 1e-19) {
    s = v.toExponential(14);
    const m = /^(-?)(\d)(?:\.(\d+))?e([+-])(\d+)$/.exec(s)!;
    const frac = (m[3] ?? "").replace(/0+$/, "");
    return `${m[1]}${m[2]}${frac ? decimal + frac : ""}E${m[4]}${m[5].padStart(2, "0")}`;
  }
  s = Number(v.toPrecision(15)).toString();
  if (s.includes("e")) {
    // Küçük sayılar için üstel olmayan gösterim.
    s = Number(v.toPrecision(15)).toFixed(20).replace(/0+$/, "").replace(/\.$/, "");
  }
  return s.replace(".", decimal);
}

function groupDigits(int: string, sep: string): string {
  let out = "";
  for (let i = 0; i < int.length; i++) {
    if (i > 0 && (int.length - i) % 3 === 0) out += sep;
    out += int[i];
  }
  return out;
}

/** Sayı bölümünü biçimlendirir. */
function formatNumberSection(value: number, sec: Section, o: Required<FormatOptions>, prefixMinus: boolean): string {
  const toks = sec.toks;
  let v = value;
  const pct = toks.filter((t) => t.t === "percent").length;
  if (pct) v *= Math.pow(100, pct);

  // Kesir biçimi (# ?/?)
  const slashIdx = toks.findIndex((t) => t.t === "slash");
  if (slashIdx > 0 && toks.slice(0, slashIdx).some((t) => t.t === "digit") && toks.slice(slashIdx + 1).some((t) => t.t === "digit" || (t.t === "lit" && /\d/.test(t.v)))) {
    const r = formatFraction(Math.abs(v), toks, slashIdx, o);
    return (prefixMinus && v < 0 && r.trim() !== "0" ? "-" : "") + r;
  }

  const expIdx = toks.findIndex((t) => t.t === "exp");
  const mainEnd = expIdx >= 0 ? expIdx : toks.length;
  const pointIdx = toks.slice(0, mainEnd).findIndex((t) => t.t === "point");
  const intEnd = pointIdx >= 0 ? pointIdx : mainEnd;

  // Ölçekleme ve binlik gruplama
  let lastIntDigit = -1;
  let firstIntDigit = -1;
  for (let k = 0; k < intEnd; k++) {
    if (toks[k].t === "digit") {
      if (firstIntDigit < 0) firstIntDigit = k;
      lastIntDigit = k;
    }
  }
  let scale = 0;
  let grouping = false;
  for (let k = 0; k < intEnd; k++) {
    if (toks[k].t !== "comma") continue;
    if (k > lastIntDigit && (pointIdx < 0 || true)) {
      // son basamaktan sonra gelen virgüller (ondalıktan önce) ölçekler
      if (lastIntDigit >= 0) scale++;
    } else if (firstIntDigit >= 0 && k > firstIntDigit) grouping = true;
  }
  if (pointIdx >= 0) {
    let lastF = -1;
    for (let k = pointIdx + 1; k < mainEnd; k++) if (toks[k].t === "digit") lastF = k;
    for (let k = pointIdx + 1; k < mainEnd; k++) if (toks[k].t === "comma" && k > lastF) scale++;
  }
  if (scale) v /= Math.pow(1000, scale);

  const intPh = toks.slice(0, intEnd).filter((t) => t.t === "digit") as { t: "digit"; v: string }[];
  const fracPh = pointIdx >= 0 ? (toks.slice(pointIdx + 1, mainEnd).filter((t) => t.t === "digit") as { t: "digit"; v: string }[]) : [];
  const neg = v < 0;
  let abs = Math.abs(v);
  let exponent = 0;
  if (expIdx >= 0 && abs !== 0) {
    const step = intPh.length > 1 && intPh[0].v === "#" ? intPh.length : 1;
    exponent = Math.floor(Math.log10(abs));
    exponent = Math.floor(exponent / step) * step;
    if (step === 1 && intPh.length > 1) exponent -= intPh.length - 1;
    abs = abs / Math.pow(10, exponent);
    if (roundTo(abs, fracPh.length) >= Math.pow(10, Math.max(1, step === 1 ? Math.max(1, intPh.length) : step))) {
      abs /= 10;
      exponent += 1;
    }
  }
  const rounded = roundTo(abs, fracPh.length);
  const intVal = Math.floor(rounded + 1e-9 * (fracPh.length ? 0 : 0));
  let fracStr = "";
  if (fracPh.length) {
    const fixed = rounded.toFixed(fracPh.length);
    const dot = fixed.indexOf(".");
    fracStr = dot >= 0 ? fixed.slice(dot + 1) : "".padEnd(fracPh.length, "0");
  }
  const intFixed = fracPh.length ? rounded.toFixed(fracPh.length).split(".")[0] : Math.round(rounded).toString();
  let D = intFixed === "0" ? "" : intFixed;
  if (D.includes("e")) D = BigInt(Math.round(intVal)).toString();

  // Sondaki '#' sıfırları atılır, '?' boşluk olur
  const fracChars = fracStr.split("");
  for (let k = fracChars.length - 1; k >= 0; k--) {
    if (fracChars[k] !== "0") break;
    if (fracPh[k].v === "#") fracChars[k] = "";
    else if (fracPh[k].v === "?") fracChars[k] = " ";
    else break;
  }

  let out = "";
  let phIdx = 0;
  let fracIdx = 0;
  const k = intPh.length;
  let groupedDone = false;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (i < intEnd) {
      if (t.t === "digit") {
        if (grouping) {
          if (!groupedDone) {
            const minDigits = intPh.reduce((n, d, j) => (d.v === "0" ? Math.max(n, k - j) : n), 0);
            const digits = D.padStart(minDigits, "0");
            out += digits ? groupDigits(digits, o.thousands) : "";
            groupedDone = true;
          }
        } else {
          const idx = D.length - k + phIdx;
          if (phIdx === 0 && idx > 0) out += D.slice(0, idx + 1);
          else if (idx >= 0) out += D[idx];
          else out += t.v === "0" ? "0" : t.v === "?" ? " " : "";
        }
        phIdx++;
      } else if (t.t === "lit") out += t.v;
      else if (t.t === "percent") out += "%";
      else if (t.t === "slash") out += "/";
      else if (t.t === "fill") out += "\u0000" + t.v;
      continue;
    }
    if (i === pointIdx) {
      if (k === 0 && D) out += D;
      out += o.decimal;
      continue;
    }
    if (i < mainEnd) {
      if (t.t === "digit") out += fracChars[fracIdx++] ?? "";
      else if (t.t === "lit") out += t.v;
      else if (t.t === "percent") out += "%";
      else if (t.t === "fill") out += "\u0000" + t.v;
      continue;
    }
    if (i === expIdx) {
      if (k === 0 && pointIdx < 0 && D) out += D;
      const ed = toks.slice(expIdx + 1).filter((x) => x.t === "digit").length;
      const sign = exponent < 0 ? "-" : (t as { sign: string }).sign === "+" ? "+" : "";
      out += "E" + sign + Math.abs(exponent).toString().padStart(ed, "0");
      continue;
    }
    if (i > expIdx) {
      if (t.t === "lit") out += t.v;
      else if (t.t === "percent") out += "%";
      else if (t.t === "fill") out += "\u0000" + t.v;
    }
  }
  if (k === 0 && pointIdx < 0 && expIdx < 0 && D && !toks.some((t) => t.t === "digit")) out = out;
  const isZero = rounded === 0;
  return (neg && prefixMinus && !isZero ? "-" : "") + out;
}

function roundTo(x: number, d: number): number {
  if (d > 15) d = 15;
  const f = Math.pow(10, d);
  const r = Math.round((x + Number.EPSILON * x) * f) / f;
  return r;
}

function gcd(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}

function formatFraction(abs: number, toks: Tok[], slashIdx: number, o: Required<FormatOptions>): string {
  const left = toks.slice(0, slashIdx);
  const right = toks.slice(slashIdx + 1);
  // Paydanın sabit sayı olup olmadığı
  const rightLits = right.filter((t) => t.t === "lit" || t.t === "digit");
  const fixedDen = /^\d+$/.test(rightLits.map((t) => (t.t === "lit" ? t.v : t.v)).join("")) && rightLits.every((t) => t.t === "lit" || t.v === "0") && right.some((t) => t.t === "lit" && /\d/.test(t.v));
  const denDigits = right.filter((t) => t.t === "digit").length || 1;
  // Tam kısım var mı? (sol tarafta boşlukla ayrılmış iki basamak grubu)
  let wholeToks: Tok[] = [];
  let numToks: Tok[] = left;
  const spaceIdx = left.findIndex((t, k) => t.t === "lit" && t.v === " " && left.slice(0, k).some((x) => x.t === "digit"));
  if (spaceIdx >= 0) {
    wholeToks = left.slice(0, spaceIdx);
    numToks = left.slice(spaceIdx + 1);
  }
  let whole = wholeToks.length ? Math.floor(abs) : 0;
  const frac = wholeToks.length ? abs - whole : abs;
  let num: number;
  let den: number;
  if (fixedDen) {
    den = parseInt(rightLits.map((t) => (t.t === "lit" ? t.v : t.v)).join(""), 10);
    num = Math.round(frac * den);
  } else {
    const maxDen = Math.pow(10, denDigits) - 1;
    [num, den] = bestFraction(frac, maxDen);
  }
  if (num === den && wholeToks.length) {
    whole += 1;
    num = 0;
  }
  const g = fixedDen ? 1 : gcd(num, den) || 1;
  num /= g;
  den /= g;
  const pad = (s: string, ph: Tok[]) => {
    const n = ph.filter((t) => t.t === "digit").length;
    return s.padStart(n, " ");
  };
  let out = "";
  if (wholeToks.length) {
    out += whole === 0 && num !== 0 ? pad("", wholeToks) : String(whole);
    if (num === 0) return out.trim() === "" ? "0" : out + " ".repeat(numToks.length + 1 + right.length);
    out += " ";
  }
  if (!wholeToks.length && num === 0 && abs === 0) return "0";
  out += pad(String(num), numToks) + "/" + (fixedDen ? String(den) : String(den).padEnd(denDigits, " "));
  void o;
  return out;
}

function bestFraction(x: number, maxDen: number): [number, number] {
  if (x === 0) return [0, 1];
  let bestN = 0;
  let bestD = 1;
  let bestErr = Infinity;
  // Sürekli kesir yaklaşımı
  let h1 = 1,
    h0 = 0,
    k1 = 0,
    k0 = 1;
  let b = x;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(b);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDen) break;
    const err = Math.abs(x - h2 / k2);
    if (err < bestErr) {
      bestErr = err;
      bestN = h2;
      bestD = k2;
    }
    h0 = h1;
    h1 = h2;
    k0 = k1;
    k1 = k2;
    if (b - a < 1e-12) break;
    b = 1 / (b - a);
  }
  // Ara yaklaşımlar için kaba tarama (küçük paydalar)
  if (maxDen <= 999) {
    for (let d = 1; d <= maxDen; d++) {
      const n = Math.round(x * d);
      const err = Math.abs(x - n / d);
      if (err < bestErr - 1e-15) {
        bestErr = err;
        bestN = n;
        bestD = d;
      }
    }
  }
  return [bestN, bestD];
}

function formatDateSection(serial: number, sec: Section, o: Required<FormatOptions>): string {
  if (serial < 0) return "#".repeat(10);
  const fracSec = sec.toks.find((t) => t.t === "fracsec") as { n: number } | undefined;
  const precision = fracSec ? Math.pow(10, fracSec.n) : 1;
  // Gösterilen hassasiyete yuvarla
  let s = Math.round(serial * 86400 * precision) / (86400 * precision);
  const hasDateTok = sec.toks.some((t) => t.t === "date" && /^[ydm]/.test(t.v));
  if (!hasDateTok && !sec.toks.some((t) => t.t === "elapsed")) {
    // yalnız saat: gün kısmı önemsiz
  }
  if (s < 0) s = 0;
  const dp = serialToDate(s);
  const tp = serialToTime(s);
  const totalSeconds = s * 86400;
  const months = o.lang === "tr" ? MONTHS_TR : MONTHS_EN;
  const monthsShort = o.lang === "tr" ? MONTHS_TR_SHORT : MONTHS_EN.map((m) => m.slice(0, 3));
  const days = o.lang === "tr" ? DAYS_TR : DAYS_EN;
  const daysShort = o.lang === "tr" ? DAYS_TR_SHORT : DAYS_EN.map((d) => d.slice(0, 3));
  let out = "";
  for (const t of sec.toks) {
    switch (t.t) {
      case "lit":
        out += t.v;
        break;
      case "date": {
        const v = t.v;
        if (v[0] === "y") out += v.length <= 2 ? String(dp.y % 100).padStart(2, "0") : String(dp.y);
        else if (v === "m") out += String(dp.m);
        else if (v === "mm") out += String(dp.m).padStart(2, "0");
        else if (v === "mmm") out += monthsShort[dp.m - 1];
        else if (v === "mmmmm") out += months[dp.m - 1][0];
        else if (v.startsWith("mmmm")) out += months[dp.m - 1];
        else if (v === "M") out += String(tp.mi);
        else if (v === "MM") out += String(tp.mi).padStart(2, "0");
        else if (v === "d") out += String(dp.d);
        else if (v === "dd") out += String(dp.d).padStart(2, "0");
        else if (v === "ddd") out += daysShort[weekdayOf(s)];
        else if (v.startsWith("dddd")) out += days[weekdayOf(s)];
        else if (v[0] === "h") {
          let h = tp.h;
          if (sec.hasAmPm) {
            h = h % 12;
            if (h === 0) h = 12;
          }
          out += v.length >= 2 ? String(h).padStart(2, "0") : String(h);
        } else if (v[0] === "s") out += v.length >= 2 ? String(tp.s).padStart(2, "0") : String(tp.s);
        break;
      }
      case "elapsed": {
        let n: number;
        if (t.v === "h") n = Math.floor(totalSeconds / 3600 + 1e-9);
        else if (t.v === "m") n = Math.floor(totalSeconds / 60 + 1e-9);
        else n = Math.floor(totalSeconds + 1e-9);
        out += String(n).padStart(t.n, "0");
        break;
      }
      case "fracsec":
        out += o.decimal + String(Math.round(tp.ms / Math.pow(10, 3 - t.n))).padStart(t.n, "0");
        break;
      case "ampm": {
        const pm = tp.h >= 12;
        if (o.lang === "tr") out += t.v === "A/P" ? (pm ? "S" : "Ö") : pm ? "ÖS" : "ÖÖ";
        else out += t.v === "A/P" ? (pm ? "P" : "A") : pm ? "PM" : "AM";
        break;
      }
      case "digit":
        out += t.v === "0" ? "0" : "";
        break;
      case "point":
        out += ".";
        break;
      case "comma":
        out += ",";
        break;
      case "slash":
        out += "/";
        break;
      case "fill":
        out += "\u0000" + t.v;
        break;
      case "percent":
        out += "%";
        break;
      default:
        break;
    }
  }
  return out;
}

function formatTextSection(text: string, sec: Section): string {
  let out = "";
  for (const t of sec.toks) {
    if (t.t === "text") out += text;
    else if (t.t === "lit") out += t.v;
    else if (t.t === "fill") out += "\u0000" + t.v;
  }
  return out;
}

/** Dolgu işaretini ayıklar. */
function extractFill(res: FormatResult): FormatResult {
  const i = res.text.indexOf("\u0000");
  if (i < 0) return res;
  const char = res.text[i + 1] ?? " ";
  const text = res.text.replace(/\u0000./gs, "");
  return { ...res, text, fill: { index: i, char } };
}

/** Sayıyı biçim koduyla biçimlendirir. */
export function formatNumber(value: number, code: string | undefined, opts: FormatOptions = {}): FormatResult {
  const o = { ...TR_OPTS, ...opts };
  if (!Number.isFinite(value)) return { text: "#NUM!" };
  const parsed = getParsed(code || "General");
  const secs = parsed.sections;
  let sec: Section;
  let prefixMinus = true;
  const numericSecs = secs.filter((s, i) => i < 3 && !(s.kind === "text" && i === secs.length - 1 && secs.length === 4));
  const hasConds = secs.some((s) => s.cond);
  if (hasConds) {
    const matched = secs.slice(0, 3).find((s) => s.cond && condMatch(s.cond, value));
    if (matched) {
      sec = matched;
      prefixMinus = !(value < 0 && matched.cond && matched.cond.v <= 0 && (matched.cond.op === "<" || matched.cond.op === "<="));
    } else {
      sec = secs.slice(0, 3).find((s) => !s.cond) ?? secs[0];
      prefixMinus = true;
    }
  } else if (value > 0 || numericSecs.length === 1) {
    sec = secs[0];
  } else if (value < 0) {
    sec = numericSecs[1] ?? secs[0];
    prefixMinus = numericSecs.length < 2;
  } else {
    sec = numericSecs[2] ?? secs[0];
  }
  if (sec.toks.length === 0 && secs.length > 1) return { text: "", color: sec.color };
  let text: string;
  switch (sec.kind) {
    case "general": {
      const g = formatGeneral(prefixMinus ? value : Math.abs(value), o.decimal);
      text = "";
      for (const t of sec.toks) text += t.t === "general" ? g : t.t === "lit" ? t.v : "";
      break;
    }
    case "date":
      text = formatDateSection(value, sec, o);
      break;
    case "text":
      text = formatTextSection(formatGeneral(value, o.decimal), sec);
      break;
    default:
      text = formatNumberSection(prefixMinus ? value : Math.abs(value), sec, o, true);
  }
  return extractFill({ text, color: sec.color });
}

/** Herhangi bir hücre değerini biçimlendirir. */
export function formatValue(value: unknown, code: string | undefined, opts: FormatOptions = {}): FormatResult {
  if (value === null || value === undefined) return { text: "" };
  if (typeof value === "number") return formatNumber(value, code, opts);
  if (typeof value === "boolean") return { text: (opts.lang ?? "tr") === "tr" ? (value ? "DOĞRU" : "YANLIŞ") : value ? "TRUE" : "FALSE" };
  if (typeof value === "string") {
    if (!code || code === "General") return { text: value };
    const parsed = getParsed(code);
    const textSec = parsed.sections.length === 4 ? parsed.sections[3] : parsed.sections.find((s) => s.kind === "text");
    if (textSec) return extractFill({ text: formatTextSection(value, textSec), color: textSec.color });
    return { text: value };
  }
  return { text: String(value) };
}

/* ------------------------------------------------------------------ */
/* Türkçe biçim kodu dönüşümü                                          */
/* ------------------------------------------------------------------ */

/** Kod içindeki tırnaksız/parantezsiz kısımları dönüştürür. */
function mapUnquoted(code: string, fn: (seg: string) => string, bracket?: (inner: string) => string): string {
  let out = "";
  let seg = "";
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') {
      out += fn(seg);
      seg = "";
      const j = code.indexOf('"', i + 1);
      const end = j < 0 ? code.length : j + 1;
      out += code.slice(i, end);
      i = end - 1;
    } else if (c === "[") {
      out += fn(seg);
      seg = "";
      const j = code.indexOf("]", i);
      const end = j < 0 ? code.length : j + 1;
      const inner = code.slice(i + 1, end - 1);
      out += "[" + (bracket ? bracket(inner) : inner) + "]";
      i = end - 1;
    } else if (c === "\\" || c === "_" || c === "*") {
      out += fn(seg);
      seg = "";
      out += code.slice(i, i + 2);
      i++;
    } else seg += c;
  }
  out += fn(seg);
  return out;
}

/** Kanonik kod → Türkçe gösterim ("#,##0.00" → "#.##0,00", "dd.mm.yyyy" → "gg.aa.yyyy"). */
export function formatToLocal(code: string): string {
  if (code === "General") return "Genel";
  return splitSections(code)
    .map((section) => {
      const sec = parseSection(section);
      const isDate = sec.kind === "date";
      return mapUnquoted(
        section,
        (seg) => {
          let r = seg.replace(/General/gi, "Genel");
          if (isDate) {
            r = r.replace(/AM\/PM/gi, "ÖÖ/ÖS").replace(/A\/P/gi, "Ö/S");
            // Dakika/ay ayrımı için parse edilmiş belirteçleri kullanmak yerine basit sıra kuralı
            let out = "";
            let lastWasH = false;
            for (let i = 0; i < r.length; ) {
              const c = r[i].toLowerCase();
              if ("ymdhs".includes(c)) {
                let j = i;
                while (j < r.length && r[j].toLowerCase() === c) j++;
                const run = j - i;
                let rest = r.slice(j);
                const nextIsS = /^[^a-z]*s/i.test(rest);
                let rep: string;
                if (c === "y") rep = "y".repeat(run);
                else if (c === "d") rep = "g".repeat(run);
                else if (c === "h") rep = "s".repeat(run);
                else if (c === "s") rep = "n".repeat(run);
                else rep = (lastWasH || nextIsS) && run <= 2 ? "d".repeat(run) : "a".repeat(run);
                lastWasH = c === "h";
                out += rep;
                i = j;
                rest = "";
                void rest;
              } else {
                out += r[i];
                i++;
              }
            }
            r = out;
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ".");
          } else {
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ",");
          }
          return r;
        },
        (inner) => {
          const key = Object.keys(COLOR_TR).find((k) => k.toLowerCase() === inner.toLowerCase());
          if (key) return COLOR_TR[key];
          if (/^(h+|m+|s+)$/i.test(inner)) return inner.toLowerCase().replace(/h/g, "s").replace(/m/g, "d").replace(/s(?=s*$)/g, (x) => x);
          return inner.replace(/^(<=|>=|<>|<|>|=)(\s*-?)(\d+)\.(\d+)$/, "$1$2$3,$4");
        },
      );
    })
    .join(";");
}

/** Kodun Türkçe yazımla mı girildiğini tahmin eder. */
export function looksLocalFormat(code: string): boolean {
  let letters = "";
  mapUnquoted(code, (seg) => {
    letters += seg.replace(/genel|general|am\/pm|a\/p|öö\/ös|ö\/s|e[+-]/gi, "");
    return seg;
  });
  const l = letters.toLocaleLowerCase("tr-TR");
  if (/[mh]/.test(l)) return false;
  if (/[gan]/.test(l)) return true;
  if (/s/.test(l) && /d/.test(l)) return true;
  if (/[dsy]/.test(l)) return false;
  // Yalnız sayı biçimi: ayırıcıların sırasına bak
  const code2 = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");
  const iDot = code2.indexOf(".");
  const iCom = code2.indexOf(",");
  if (iDot >= 0 && iCom >= 0) return iDot < iCom;
  if (iDot >= 0) return /\.(##0|###)/.test(code2);
  if (iCom >= 0) return !/,(##0|###)/.test(code2) && !/[0#],+(\s|$|;|%)/.test(code2);
  return true;
}

/** Türkçe kod → kanonik ("#.##0,00" → "#,##0.00", "gg.aa.yyyy" → "dd.mm.yyyy"). */
export function formatFromLocal(code: string): string {
  const t = code.trim();
  if (/^genel$/i.test(t)) return "General";
  return splitSections(t)
    .map((section) => {
      const lower = section.toLocaleLowerCase("tr-TR");
      const unq = lower.replace(/"[^"]*"|\[[^\]]*\]/g, "").replace(/öö\/ös|ö\/s|genel/g, "");
      const isDate = /[gaysdn]/.test(unq.replace(/e[+-]/g, ""));
      return mapUnquoted(
        section,
        (seg) => {
          let r = seg.replace(/Genel/gi, "General");
          if (isDate) {
            r = r.replace(/ÖÖ\/ÖS/gi, "AM/PM").replace(/Ö\/S/gi, "A/P");
            let out = "";
            for (let i = 0; i < r.length; ) {
              const c = r[i].toLocaleLowerCase("tr-TR");
              if ("gaysdn".includes(c)) {
                let j = i;
                while (j < r.length && r[j].toLocaleLowerCase("tr-TR") === c) j++;
                const run = j - i;
                const map: Record<string, string> = { g: "d", a: "m", y: "y", s: "h", d: "m", n: "s" };
                out += map[c].repeat(run);
                i = j;
              } else {
                out += r[i];
                i++;
              }
            }
            r = out;
          } else {
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ",");
          }
          return r;
        },
        (inner) => {
          const en = COLOR_EN[inner.toLocaleLowerCase("tr-TR")];
          if (en) return en;
          if (/^(s+|d+|n+)$/i.test(inner)) return inner.toLowerCase().replace(/s/g, "h").replace(/d/g, "m").replace(/n/g, "s");
          return inner.replace(/^(<=|>=|<>|<|>|=)(\s*-?)(\d+),(\d+)$/, "$1$2$3.$4");
        },
      );
    })
    .join(";");
}
