/** Dosya formülü dönüşümleri: _xlfn. önekleri, tarih seri numarası. */
import { EN_SYNTAX, FUNCTIONS, tokenize } from "@officehere/formula";

export function utcSerial(d: Date): number {
  return 25569 + d.getTime() / 86400000;
}

/** Dosyaya yazarken yeni fonksiyonlara "_xlfn." öneki eklenir (metin sabitleri atlanır). */
export function addXlfnPrefixes(formula: string): string {
  return tokenize(formula, EN_SYNTAX)
    .map((t) => {
      if (t.type !== "func") return t.text;
      const def = FUNCTIONS.get(String(t.value).toUpperCase());
      return def?.xlfn ? `_xlfn.${t.text}` : t.text;
    })
    .join("");
}

/** Okurken önekleri kaldırır. */
export function normalizeFileFormula(f: string): string {
  return f.replace(/_xlfn\.|_xlws\.|_xludf\./g, "").replace(/_xlpm\./g, "");
}
