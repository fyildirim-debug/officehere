/**
 * Excel tür zorlama kuralları: sayıya/metne/mantıksala çevirme,
 * karşılaştırma sıralaması (sayı < metin < mantıksal) ve ölçüt eşleme.
 */
import { parseDateTime } from "./dates";
import { numberToText } from "./numfmt";
import { Errors, FormulaError, isError, type Scalar } from "./values";

export interface LocaleInfo {
  decimal: "." | ",";
  lang: "tr" | "en";
}

export const TR_LOCALE: LocaleInfo = { decimal: ",", lang: "tr" };
export const EN_LOCALE: LocaleInfo = { decimal: ".", lang: "en" };

const collatorTr = new Intl.Collator("tr", { sensitivity: "accent" });
const collatorEn = new Intl.Collator("en", { sensitivity: "accent" });

/**
 * Metni sayıya çevirmeyi dener: "1,5", "1.234,5", "1.5", "1e3", "%50", "50%",
 * "₺10", "10 TL", "(5)", tarih ve saat metinleri. Başarısızsa null.
 */
export function parseNumberText(text: string, locale: LocaleInfo = TR_LOCALE): number | null {
  let s = text.trim();
  if (s === "") return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1).trim();
  }
  let pct = false;
  if (s.endsWith("%")) {
    pct = true;
    s = s.slice(0, -1).trim();
  } else if (s.startsWith("%")) {
    pct = true;
    s = s.slice(1).trim();
  }
  s = s.replace(/^[₺$€£]\s*/, "").replace(/\s*(₺|TL|\$|€|£)$/i, "");
  let sign = 1;
  if (s.startsWith("-")) {
    sign = -1;
    s = s.slice(1).trim();
  } else if (s.startsWith("+")) s = s.slice(1).trim();
  s = s.replace(/^[₺$€£]\s*/, "");
  let n: number | null = null;
  const dec = locale.decimal;
  const thou = dec === "," ? "." : ",";
  const reLocal = new RegExp(`^(\\d{1,3}(\\${thou}\\d{3})+|\\d+)(\\${dec}\\d*)?([eE][+-]?\\d+)?$`);
  const reLocalFrac = new RegExp(`^\\${dec}\\d+([eE][+-]?\\d+)?$`);
  if (reLocal.test(s) || reLocalFrac.test(s)) {
    n = parseFloat(s.split(thou).join("").replace(dec, "."));
  } else if (/^\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s) || /^\.\d+$/.test(s)) {
    // Ayırıcıdan bağımsız: "1.5" her yerelde 1,5 kabul edilir
    n = parseFloat(s);
  } else if (/^\d+,\d+$/.test(s)) {
    n = parseFloat(s.replace(",", "."));
  }
  if (n === null && !pct && sign === 1 && !neg) {
    const dt = parseDateTime(text);
    if (dt) return dt.serial;
    return null;
  }
  if (n === null || !Number.isFinite(n)) return null;
  n *= sign;
  if (neg) n = -n;
  if (pct) n /= 100;
  return n;
}

export function toNumber(v: Scalar, locale: LocaleInfo = TR_LOCALE): number | FormulaError {
  if (typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isError(v)) return v;
  const n = parseNumberText(v, locale);
  return n === null ? Errors.VALUE : n;
}

export function boolText(b: boolean, locale: LocaleInfo = TR_LOCALE): string {
  return locale.lang === "tr" ? (b ? "DOĞRU" : "YANLIŞ") : b ? "TRUE" : "FALSE";
}

export function toText(v: Scalar, locale: LocaleInfo = TR_LOCALE): string | FormulaError {
  if (typeof v === "string") return v;
  if (v === null) return "";
  if (typeof v === "number") return numberToText(v, locale.decimal);
  if (typeof v === "boolean") return boolText(v, locale);
  return v;
}

export function toBool(v: Scalar): boolean | FormulaError {
  if (typeof v === "boolean") return v;
  if (v === null) return false;
  if (typeof v === "number") return v !== 0;
  if (isError(v)) return v;
  const u = v.trim().toLocaleUpperCase("tr-TR");
  if (u === "TRUE" || u === "DOĞRU") return true;
  if (u === "FALSE" || u === "YANLIŞ") return false;
  return Errors.VALUE;
}

function typeRank(v: Scalar): number {
  if (typeof v === "number") return 0;
  if (typeof v === "string") return 1;
  if (typeof v === "boolean") return 2;
  return 3;
}

export function textEquals(a: string, b: string): boolean {
  if (a === b) return true;
  if (a.length !== b.length && Math.abs(a.length - b.length) > 2) return false;
  return a.toLocaleUpperCase("tr-TR") === b.toLocaleUpperCase("tr-TR") || a.toUpperCase() === b.toUpperCase();
}

export function compareText(a: string, b: string, locale: LocaleInfo = TR_LOCALE): number {
  if (textEquals(a, b)) return 0;
  const c = (locale.lang === "tr" ? collatorTr : collatorEn).compare(a, b);
  return c < 0 ? -1 : c > 0 ? 1 : 0;
}

/** Sayı eşitliği (Excel 15 anlamlı basamak toleransı). */
export function numEquals(a: number, b: number): boolean {
  if (a === b) return true;
  const m = Math.max(Math.abs(a), Math.abs(b));
  return Math.abs(a - b) <= m * 1e-15;
}

/**
 * Excel karşılaştırması: -1, 0, 1. Boş hücre karşı tarafın türünde
 * sıfır değeri gibi davranır. Hatalar çağıran tarafından ele alınmalıdır.
 */
export function compareScalars(a: Scalar, b: Scalar, locale: LocaleInfo = TR_LOCALE): number {
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === "number") {
    const bn = b as number;
    if (numEquals(a, bn)) return 0;
    return a < bn ? -1 : 1;
  }
  if (typeof a === "string") return compareText(a, b as string, locale);
  if (typeof a === "boolean") return a === b ? 0 : a ? 1 : -1;
  return 0;
}

/* ------------------------------------------------------------------ */
/* Ölçütler (EĞERSAY, ETOPLA …)                                         */
/* ------------------------------------------------------------------ */

export type Criteria = (v: Scalar) => boolean;

/** Joker karakterli deseni (* ? ~) düzenli ifadeye çevirir. */
export function wildcardToRegExp(pattern: string, whole = true): RegExp {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "~" && i + 1 < pattern.length && "*?~".includes(pattern[i + 1])) {
      re += "\\" + pattern[++i];
    } else if (c === "*") re += "[\\s\\S]*";
    else if (c === "?") re += "[\\s\\S]";
    else re += c.replace(/[.+^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(whole ? `^${re}$` : re, "iu");
}

export function hasWildcards(s: string): boolean {
  return /[*?]/.test(s.replace(/~[*?~]/g, ""));
}

function trUpper(s: string): string {
  return s.toLocaleUpperCase("tr-TR");
}

/** Ölçüt değerinden eşleme işlevi üretir: 5, ">5", "<>a", "=b*", "". */
export function makeCriteria(crit: Scalar, locale: LocaleInfo = TR_LOCALE): Criteria {
  if (isError(crit)) {
    const code = crit.code;
    return (v) => isError(v) && v.code === code;
  }
  if (typeof crit === "number") return (v) => (typeof v === "number" ? numEquals(v, crit) : typeof v === "string" ? parseNumberText(v, locale) === crit : false);
  if (typeof crit === "boolean") return (v) => v === crit;
  if (crit === null) return (v) => v === 0;
  const s = crit;
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(s)!;
  const op = m[1] ?? "=";
  const rhs = m[2];
  // Sağ taraf sayı mı?
  const rn = rhs.trim() === "" ? null : parseNumberText(rhs, locale);
  const ub = trUpper(rhs.trim());
  const rb = ub === "TRUE" || ub === "DOĞRU" ? true : ub === "FALSE" || ub === "YANLIŞ" ? false : null;
  if (op === "=" || op === "<>") {
    const neg = op === "<>";
    let test: Criteria;
    if (rhs === "") {
      // "=" boş hücreler; "<>" boş olmayanlar
      test = (v) => v === null || v === "";
      if (neg) return (v) => !(v === null);
      if (m[1] === "=") return (v) => v === null;
      return test;
    } else if (rn !== null) {
      test = (v) => (typeof v === "number" && numEquals(v, rn)) || (typeof v === "string" && parseNumberText(v, locale) === rn);
    } else if (rb !== null) {
      test = (v) => v === rb;
    } else if (rhs.startsWith("#")) {
      test = (v) => isError(v) && v.code === rhs.toUpperCase();
    } else if (hasWildcards(rhs)) {
      const re = wildcardToRegExp(rhs);
      test = (v) => typeof v === "string" && (re.test(v) || re.test(trUpper(v)));
    } else {
      const lit = rhs.replace(/~([*?~])/g, "$1");
      test = (v) => typeof v === "string" && textEquals(v, lit);
    }
    return neg ? (v) => !test(v) : test;
  }
  // < > <= >=
  const cmp = (c: number) => (op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0);
  if (rn !== null) return (v) => typeof v === "number" && cmp(v < rn && !numEquals(v, rn) ? -1 : numEquals(v, rn) ? 0 : 1);
  if (rb !== null) return (v) => typeof v === "boolean" && cmp(compareScalars(v, rb, locale));
  return (v) => typeof v === "string" && cmp(compareText(v, rhs, locale));
}
