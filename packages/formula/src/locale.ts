/**
 * Formül yerelleştirme: kanonik (İngilizce, "," ayırıcı, "." ondalık)
 * ⇄ Türkçe (TOPLA, ";" ayırıcı, "," ondalık). Girişte iki yazım da kabul edilir.
 */
import "./functions";
import { BOOL_NAMES, EN_SYNTAX, ERROR_FROM_TR, ERROR_TO_TR, TR_SYNTAX, tokenize, type SyntaxLocale, type Token } from "./lexer";
import { parse as parseRaw, stripFunctionPrefix } from "./parser";
import { FUNCTIONS, resolveFunctionName, trNameOf } from "./functions/registry";
import { renderRefToken } from "./transform";
import { renderStructRef } from "./structref";
import type { Node } from "./ast";

export type FormulaLocale = "tr" | "en";

export function syntaxOf(locale: FormulaLocale): SyntaxLocale {
  return locale === "tr" ? TR_SYNTAX : EN_SYNTAX;
}

function convertTokens(tokens: Token[], from: SyntaxLocale, to: SyntaxLocale, toLocale: FormulaLocale, fromLocale: FormulaLocale): string {
  let out = "";
  for (const t of tokens) {
    switch (t.type) {
      case "func": {
        const raw = stripFunctionPrefix(t.value as string);
        const canonical = resolveFunctionName(raw, fromLocale === "tr");
        if (!canonical) {
          // Bilinmeyen ad (kullanıcı tanımlı işlev, LAMBDA adı) yazıldığı gibi kalır: Türkçe büyük harf
          // kuralı "i"yi "İ" yapıp dosyadaki adı bozmasın (LibreOffice hata 142293), İngilizce kural da
          // "işle"yi "IŞLE" yapıp Türkçe adlı LAMBDA'yı koparmasın. Değerlendirici adları kendisi eşler.
          out += raw;
          break;
        }
        out += toLocale === "tr" ? trNameOf(canonical) : canonical;
        break;
      }
      case "number": {
        let s = t.text;
        if (from.decimal !== to.decimal) s = s.replace(from.decimal, to.decimal);
        out += s;
        break;
      }
      case "bool": {
        const v = t.value as boolean;
        out += toLocale === "tr" ? (v ? "DOĞRU" : "YANLIŞ") : v ? "TRUE" : "FALSE";
        break;
      }
      case "error": {
        const canonical = t.value as string;
        out += toLocale === "tr" ? (ERROR_TO_TR[canonical] ?? canonical) : canonical;
        break;
      }
      case "ref":
        if (t.ref!.refError) out += renderRefToken(t.ref!, toLocale);
        else out += t.text;
        break;
      case "structref":
        out += renderStructRef(t.sref!, toLocale);
        break;
      case "sep":
        out += to.argSep;
        break;
      case "arrayCol":
        out += to.arrayCol;
        break;
      case "arrayRow":
        out += to.arrayRow;
        break;
      default:
        out += t.text;
    }
  }
  return out;
}

/** Kanonik formülü kullanıcı yerel ayarına çevirir ("=SUM(A1,1.5)" → "=TOPLA(A1;1,5)"). */
export function toLocalFormula(canonical: string, locale: FormulaLocale = "tr"): string {
  const hasEq = canonical.startsWith("=");
  const body = hasEq ? canonical.slice(1) : canonical;
  if (locale === "en") return canonical;
  const tokens = tokenize(body, EN_SYNTAX);
  return (hasEq ? "=" : "") + convertTokens(tokens, EN_SYNTAX, TR_SYNTAX, "tr", "en");
}

/** Kullanıcının yazdığı formülü kanonik biçime çevirir; Türkçe ve İngilizce adlar kabul edilir. */
export function fromLocalFormula(local: string, locale: FormulaLocale = "tr"): string {
  const hasEq = local.startsWith("=");
  const body = hasEq ? local.slice(1) : local;
  const from = syntaxOf(locale);
  const tokens = tokenize(body, from);
  return (hasEq ? "=" : "") + convertTokens(tokens, from, EN_SYNTAX, "en", locale);
}

/** Kanonik formülü ayrıştırır. */
export function parseFormula(canonical: string): Node {
  return parseRaw(canonical, { syntax: EN_SYNTAX, resolveFunction: (n) => resolveFunctionName(n, false) });
}

/** Yerel formülü doğrudan ayrıştırır (doğrulama için). */
export function parseLocalFormula(local: string, locale: FormulaLocale = "tr"): Node {
  return parseRaw(local, { syntax: syntaxOf(locale), resolveFunction: (n) => resolveFunctionName(n, locale === "tr") });
}

/** Fonksiyon adı listesi (otomatik tamamlama): yerel adlar. */
export function functionNames(locale: FormulaLocale = "tr"): { name: string; canonical: string }[] {
  return [...FUNCTIONS.values()].map((d) => ({ name: locale === "tr" ? d.tr : d.name, canonical: d.name })).sort((a, b) => a.name.localeCompare(b.name, locale));
}

/** Hata kodunu yerel adına çevirir (#DIV/0! → #SAYI/0!). */
export function localErrorName(code: string, locale: FormulaLocale = "tr"): string {
  return locale === "tr" ? (ERROR_TO_TR[code] ?? code) : code;
}

export function canonicalErrorName(text: string): string | undefined {
  const u = text.toLocaleUpperCase("tr-TR");
  if (ERROR_FROM_TR[u]) return ERROR_FROM_TR[u];
  if (Object.prototype.hasOwnProperty.call(ERROR_TO_TR, text.toUpperCase())) return text.toUpperCase();
  return undefined;
}

export function localBoolName(b: boolean, locale: FormulaLocale = "tr"): string {
  return locale === "tr" ? (b ? "DOĞRU" : "YANLIŞ") : b ? "TRUE" : "FALSE";
}

export { BOOL_NAMES };
