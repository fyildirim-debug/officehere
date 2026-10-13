/** Dosya formülü dönüşümleri: _xlfn. / _xlws. / _xlpm. önekleri, tarih seri numarası. */
import { EN_SYNTAX, FUNCTIONS, tokenize, type Token } from "@officehere/formula";

export function utcSerial(d: Date): number {
  return 25569 + d.getTime() / 86400000;
}

/** Dosyada "_xlfn._xlws." önekiyle yazılan çalışma sayfası işlevleri. */
const XLWS = new Set(["FILTER", "SORT"]);

/** Bağımsız değişkenleri ad bildiren işlevler: LAMBDA (sonuncusu dışında hepsi), LET (çift sıradakiler). */
function declaresName(fn: string, argIndex: number): boolean {
  return fn === "LAMBDA" || (fn === "LET" && argIndex % 2 === 0);
}

interface Frame {
  fn: string | null;
  arg: number;
  names: Set<string>;
}

/**
 * Dosyaya yazarken yeni işlevlere "_xlfn." (FILTER/SORT için "_xlfn._xlws.") öneki,
 * LET/LAMBDA değişkenlerine ve onlara yapılan başvurulara "_xlpm." öneki eklenir.
 * Önekler yazılmazsa masaüstü hesap tabloları formülü #AD? olarak gösterir.
 * Metin sabitleri olduğu gibi kalır.
 */
export function addXlfnPrefixes(formula: string): string {
  const toks = tokenize(formula, EN_SYNTAX);
  const stack: Frame[] = [];
  const inScope = (name: string) => stack.some((f) => f.names.has(name.toUpperCase()));
  const significant = (i: number, dir: 1 | -1): Token | undefined => {
    for (let k = i + dir; k >= 0 && k < toks.length; k += dir) if (toks[k].type !== "ws") return toks[k];
    return undefined;
  };
  let out = "";
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.type) {
      case "func": {
        const name = String(t.value).toUpperCase();
        const def = FUNCTIONS.get(name);
        if (def?.xlfn) out += (XLWS.has(name) ? "_xlfn._xlws." : "_xlfn.") + t.text;
        else if (!def && inScope(name)) out += "_xlpm." + t.text;
        else out += t.text;
        // Hemen ardından gelen "(" bu işlevin çerçevesini açar
        stack.push({ fn: name, arg: 0, names: new Set() });
        const next = toks[i + 1];
        if (next?.type === "lparen") {
          out += next.text;
          i++;
        }
        break;
      }
      case "lparen":
        stack.push({ fn: null, arg: 0, names: new Set() });
        out += t.text;
        break;
      case "rparen":
        stack.pop();
        out += t.text;
        break;
      case "sep":
        if (stack.length) stack[stack.length - 1].arg++;
        out += t.text;
        break;
      case "name": {
        const name = String(t.value ?? t.text).toUpperCase();
        const top = stack[stack.length - 1];
        const prev = significant(i, -1);
        const next = significant(i, 1);
        // Bağımsız değişkenin tamamı olan ad, ardından "," geliyorsa bildirimdir
        const whole = (prev?.type === "lparen" || prev?.type === "sep") && next?.type === "sep";
        if (top?.fn && whole && declaresName(top.fn, top.arg)) {
          top.names.add(name);
          out += "_xlpm." + t.text;
        } else out += inScope(name) ? "_xlpm." + t.text : t.text;
        break;
      }
      default:
        out += t.text;
    }
  }
  return out;
}

/** Okurken önekleri kaldırır. */
export function normalizeFileFormula(f: string): string {
  return f.replace(/_xlfn\.|_xlws\.|_xludf\./g, "").replace(/_xlpm\./g, "");
}
