/**
 * Düzenli ifade işlevleri (REGEXTEST, REGEXEXTRACT, REGEXREPLACE) ve dizi/değer
 * metinleştirme (DİZİDENMETNE, DEĞERDENMETNE). Desenler PCRE uyumlu temel sözdizimiyle
 * çalışır; büyük/küçük harf duyarlılığı 0 (duyarlı, varsayılan) ya da 1 (duyarsız).
 */
import { boolText, toText } from "../coerce";
import type { FnContext } from "../context";
import { Errors, isArray, isError, isRef, makeArray, type Scalar, type Value } from "../values";
import { matrix, optNum, scalar, text } from "./helpers";
import { fn, type FnDef } from "./registry";

const cache = new Map<string, RegExp>();

/** Deseni derler: önce Unicode kipinde, olmazsa (ör. "\-" kaçışı) klasik kipte; geçersizse #DEĞER!. */
function compile(pattern: string, flags: string): RegExp {
  const key = flags + "\u0000" + pattern;
  const hit = cache.get(key);
  if (hit) return hit;
  let re: RegExp;
  try {
    re = new RegExp(pattern, flags + "u");
  } catch {
    try {
      re = new RegExp(pattern, flags);
    } catch {
      throw Errors.VALUE;
    }
  }
  if (cache.size > 200) cache.clear();
  cache.set(key, re);
  return re;
}

function caseFlag(v: Value | undefined, c: FnContext): string {
  const n = Math.trunc(optNum(v, 0, c));
  if (n !== 0 && n !== 1) throw Errors.VALUE;
  return n === 1 ? "i" : "";
}

/** Metin bağımsız değişkeni dizi/aralıksa işlev öğe öğe uygulanır. */
function perText(v: Value, c: FnContext, f: (s: string) => Scalar): Value {
  if (isArray(v) || (isRef(v) && (v.r1 !== v.r2 || v.c1 !== v.c2))) {
    return makeArray(matrix(v, c).map((r) => r.map((x) => (isError(x) ? x : (() => {
      try {
        return f(toTextStrict(x, c));
      } catch (e) {
        if (isError(e)) return e;
        throw e;
      }
    })()))));
  }
  return f(text(v, c));
}

function toTextStrict(x: Scalar, c: FnContext): string {
  const t = toText(x, c.locale);
  if (isError(t)) throw t;
  return t;
}

/** Değerin metni: kısa biçimde metin tırnaksız, katı biçimde tırnaklı (formülde yazıldığı gibi). */
function valueText(v: Scalar, strict: boolean, c: FnContext): string {
  if (v === null) return "";
  if (typeof v === "string") return strict ? `"${v.replace(/"/g, '""')}"` : v;
  if (typeof v === "boolean") return boolText(v, c.locale);
  if (isError(v)) return v.code;
  return toTextStrict(v, c);
}

function formatArg(v: Value | undefined, c: FnContext): boolean {
  const n = Math.trunc(optNum(v, 0, c));
  if (n !== 0 && n !== 1) throw Errors.VALUE;
  return n === 1;
}

export const regexFunctions: FnDef[] = [
  fn("REGEXTEST", "REGEXTEST", "text", "metin; desen; [büyük_küçük_harf_duyarlılığı]", "Metnin bir kısmının desenle eşleşip eşleşmediğini sınar.", 2, 3, (a, c) => {
    const re = compile(text(a[1], c), caseFlag(a[2], c));
    return perText(a[0], c, (s) => {
      re.lastIndex = 0;
      return re.test(s);
    });
  }, { xlfn: true }),
  fn("REGEXEXTRACT", "REGEXEXTRACT", "text", "metin; desen; [dönüş_modu]; [büyük_küçük_harf_duyarlılığı]", "Metinden desenle eşleşen bölümleri (ilk eşleşme, tüm eşleşmeler ya da yakalama grupları) ayıklar.", 2, 4, (a, c) => {
    const mode = Math.trunc(optNum(a[2], 0, c));
    if (mode < 0 || mode > 2) throw Errors.VALUE;
    const flags = caseFlag(a[3], c);
    const re = compile(text(a[1], c), flags);
    if (mode === 0) {
      return perText(a[0], c, (s) => {
        const m = re.exec(s);
        if (!m) throw Errors.NA;
        return m[0];
      });
    }
    const s = text(a[0], c);
    if (mode === 1) {
      // Tüm eşleşmeler alt alta
      const all = [...s.matchAll(compile(re.source, flags + "g"))].map((m) => [m[0] as Scalar]);
      if (!all.length) throw Errors.NA;
      return all.length === 1 ? all[0][0] : makeArray(all);
    }
    // İlk eşleşmenin yakalama grupları yan yana
    const m = re.exec(s);
    if (!m) throw Errors.NA;
    const groups = m.slice(1).map((g) => (g === undefined ? "" : g) as Scalar);
    if (!groups.length) throw Errors.NA;
    return groups.length === 1 ? groups[0] : makeArray([groups]);
  }, { xlfn: true }),
  fn("REGEXREPLACE", "REGEXREPLACE", "text", "metin; desen; değiştirme; [geçtiği_yer]; [büyük_küçük_harf_duyarlılığı]", "Metindeki desen eşleşmelerini değiştirme metniyle değiştirir ($1 grup başvurusu kullanılabilir).", 3, 5, (a, c) => {
    const flags = caseFlag(a[4], c);
    const re = compile(text(a[1], c), flags + "g");
    const rep = text(a[2], c);
    const occ = Math.trunc(optNum(a[3], 0, c));
    return perText(a[0], c, (s) => {
      if (occ === 0) return s.replace(re, rep);
      const matches = [...s.matchAll(re)];
      const idx = occ > 0 ? occ - 1 : matches.length + occ;
      const m = matches[idx];
      if (!m || m.index === undefined) return s;
      // Yalnız seçilen eşleşme değişir; grup başvuruları o eşleşmeye göre çözülür
      const one = m[0].replace(compile(re.source, flags), rep);
      return s.slice(0, m.index) + one + s.slice(m.index + m[0].length);
    });
  }, { xlfn: true }),
  fn("ARRAYTOTEXT", "DİZİDENMETNE", "text", "dizi; [biçim]", "Bir dizinin değerlerini metin olarak verir (0: kısa, 1: katı biçim).", 1, 2, (a, c) => {
    const strict = formatArg(a[1], c);
    const m = matrix(a[0], c);
    const list = c.locale.decimal === "," ? ";" : ",";
    if (!strict) return m.flat().map((v) => valueText(v, false, c)).join(list + " ");
    const colSep = c.locale.decimal === "," ? "\\" : ",";
    const rowSep = ";";
    return "{" + m.map((r) => r.map((v) => valueText(v, true, c)).join(colSep)).join(rowSep) + "}";
  }, { xlfn: true }),
  fn("VALUETOTEXT", "DEĞERDENMETNE", "text", "değer; [biçim]", "Değeri metin olarak verir (0: kısa, 1: katı biçimde metin tırnaklı).", 1, 2, (a, c) => {
    const strict = formatArg(a[1], c);
    const v = a[0];
    if (isArray(v) || (isRef(v) && (v.r1 !== v.r2 || v.c1 !== v.c2))) return makeArray(matrix(v, c).map((r) => r.map((x) => valueText(x, strict, c))));
    return valueText(scalar(v, c), strict, c);
  }, { xlfn: true }),
];

