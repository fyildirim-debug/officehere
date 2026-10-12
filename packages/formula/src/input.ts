/**
 * Kullanıcının hücreye yazdığı metni değere çevirir (tr-TR kuralları):
 * sayılar (1.234,56), yüzdeler, para birimi, tarih/saat, mantıksal değerler,
 * hata değerleri, formüller ve ' ile başlayan metinler.
 */
import { parseDateTime, dateToSerial } from "./dates";
import { fromLocalFormula, canonicalErrorName, type FormulaLocale } from "./locale";
import { errorFromCode, type Scalar } from "./values";

export interface ParsedInput {
  kind: "empty" | "value" | "formula";
  value?: Scalar;
  /** Kanonik formül ("=" ile). */
  formula?: string;
  /** Girişten çıkarılan sayı biçimi (yalnız hücrenin biçimi Genel ise uygulanır). */
  numFmt?: string;
}

export interface InputOptions {
  locale?: FormulaLocale;
  /** Hücrenin mevcut sayı biçimi. */
  numFmt?: string;
}

function parseLocalNumber(s: string, dec: string): number | null {
  const thou = dec === "," ? "." : ",";
  const re = new RegExp(`^[-+]?(\\d{1,3}(\\${thou}\\d{3})+|\\d+)?(\\${dec}\\d+)?([eE][-+]?\\d+)?$`);
  if (!re.test(s) || !/\d/.test(s)) return null;
  // "1,5E+3" gibi
  const n = parseFloat(s.split(thou).join("").replace(dec, "."));
  return Number.isFinite(n) ? n : null;
}

export function parseCellInput(raw: string, opts: InputOptions = {}): ParsedInput {
  const locale = opts.locale ?? "tr";
  const dec = locale === "tr" ? "," : ".";
  if (raw === "") return { kind: "empty" };
  if (raw.startsWith("'")) return { kind: "value", value: raw.slice(1) };
  if (opts.numFmt === "@") return { kind: "value", value: raw };
  if (raw.startsWith("=") && raw.length > 1) return { kind: "formula", formula: fromLocalFormula(raw, locale) };
  const t = raw.trim();
  if (t === "") return { kind: "value", value: raw };
  const up = t.toLocaleUpperCase("tr-TR");
  if (up === "DOĞRU" || up === "TRUE") return { kind: "value", value: true };
  if (up === "YANLIŞ" || up === "FALSE") return { kind: "value", value: false };
  if (t.startsWith("#")) {
    const code = canonicalErrorName(t);
    if (code) return { kind: "value", value: errorFromCode(code) ?? t };
  }
  // Yüzde
  let m = /^([-+]?)%\s*(.+)$/.exec(t) ?? /^([-+]?)(.+?)\s*%$/.exec(t);
  if (m) {
    const n = parseLocalNumber(m[2], dec) ?? (/^\d+(\.\d+)?$/.test(m[2]) ? parseFloat(m[2]) : null);
    if (n !== null) {
      const v = (m[1] === "-" ? -n : n) / 100;
      const decimals = (m[2].split(dec)[1] ?? "").length;
      return { kind: "value", value: v, numFmt: decimals ? `0.${"0".repeat(decimals)}%` : "0%" };
    }
  }
  // Para birimi
  m = /^([-+]?)\s*₺\s*(.+)$/.exec(t) ?? /^([-+]?)(.+?)\s*(₺|TL)$/i.exec(t);
  if (m) {
    const n = parseLocalNumber(m[2].trim(), dec);
    if (n !== null) {
      const v = m[1] === "-" ? -n : n;
      const hasDec = m[2].includes(dec);
      return { kind: "value", value: v, numFmt: hasDec ? '#,##0.00 "₺"' : '#,##0 "₺"' };
    }
  }
  // Parantezli negatif
  m = /^\((.+)\)$/.exec(t);
  if (m) {
    const n = parseLocalNumber(m[1], dec);
    if (n !== null) return { kind: "value", value: -n };
  }
  // Sayı
  const n = parseLocalNumber(t, dec);
  if (n !== null) {
    const thou = dec === "," ? "." : ",";
    // tr: "1.5" (binlik değil) → gün.ay tarihi
    if (locale === "tr" && /^\d{1,2}\.\d{1,2}$/.test(t)) {
      const [d, mo] = t.split(".").map(Number);
      if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) {
        const y = new Date().getFullYear();
        return { kind: "value", value: dateToSerial(y, mo, d), numFmt: "d mmm" };
      }
    }
    const grouped = t.includes(thou) && new RegExp(`\\d\\${thou}\\d{3}`).test(t);
    if (grouped) {
      const decimals = (t.split(dec)[1] ?? "").length;
      return { kind: "value", value: n, numFmt: decimals ? `#,##0.${"0".repeat(decimals)}` : "#,##0" };
    }
    if (/[eE]/.test(t)) return { kind: "value", value: n, numFmt: "0.00E+00" };
    return { kind: "value", value: n };
  }
  // tr dışında da "1.5" gibi nokta ondalıklı yazım
  if (/^[-+]?\d*\.\d+$/.test(t) && locale === "tr") {
    // "3.14" gibi geçersiz tarih: sayı kabul et
    const [d, mo] = t.replace(/^[-+]/, "").split(".").map(Number);
    if (!(d >= 1 && d <= 31 && mo >= 1 && mo <= 12)) return { kind: "value", value: parseFloat(t) };
  }
  // Kesir: "0 1/2", "3 3/4"
  m = /^([-+]?)(\d+)\s+(\d+)\/(\d+)$/.exec(t);
  if (m && +m[4] !== 0) {
    const v = +m[2] + +m[3] / +m[4];
    return { kind: "value", value: m[1] === "-" ? -v : v, numFmt: +m[4] < 10 ? "# ?/?" : "# ??/??" };
  }
  // Tarih / saat
  const dt = parseDateTime(t);
  if (dt) {
    let fmt = dt.format;
    if (/^dd\.mm\.yyyy/.test(fmt) || /^yyyy-mm-dd/.test(fmt)) fmt = fmt.replace(/^(dd\.mm\.yyyy|yyyy-mm-dd)/, "dd.mm.yyyy");
    return { kind: "value", value: dt.serial, numFmt: fmt };
  }
  // "1/2" → gün/ay
  m = /^(\d{1,2})\/(\d{1,2})$/.exec(t);
  if (m) {
    const d = +m[1];
    const mo = +m[2];
    if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) return { kind: "value", value: dateToSerial(new Date().getFullYear(), mo, d), numFmt: "d mmm" };
  }
  return { kind: "value", value: raw };
}
