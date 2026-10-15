/** Yazıcının XML yardımcıları: kaçış, bildirim, ad alanları. */

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
export const NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
export const NS_XM = "http://schemas.microsoft.com/office/excel/2006/main";

const SPECIAL = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;

/**
 * Metin ve öznitelik kaçışı: & < > " ' varlık olur, XML 1.0'da yasak denetim karakterleri (sekme,
 * satır sonu ve satır başı dışındakiler, DEL) atılır.
 */
export function esc(text: string): string {
  if (!SPECIAL.test(text)) return text;
  let out = "";
  let last = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    let rep: string | null;
    switch (c) {
      case 34:
        rep = "&quot;";
        break;
      case 38:
        rep = "&amp;";
        break;
      case 39:
        rep = "&apos;";
        break;
      case 60:
        rep = "&lt;";
        break;
      case 62:
        rep = "&gt;";
        break;
      case 127:
        rep = "";
        break;
      default:
        rep = c <= 31 && c !== 9 && c !== 10 && c !== 13 ? "" : null;
    }
    if (rep === null) continue;
    if (last !== i) out += text.slice(last, i);
    out += rep;
    last = i + 1;
  }
  return last === 0 ? out : out + text.slice(last);
}

/** Tanımlı (undefined olmayan) öznitelikleri sırayla yazar. */
export function attrs(list: [string, string | number | undefined][]): string {
  let out = "";
  for (const [k, v] of list) if (v !== undefined) out += ` ${k}="${esc(String(v))}"`;
  return out;
}
