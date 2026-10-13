/**
 * OpenFormula (ODF 1.2+ "of:" formülleri) ⇄ motorun kanonik A1 sözdizimi.
 * [.A1] → A1, [$Sayfa2.A1:.B3] → Sayfa2!A1:B3, ";" → ",", {1;2|3;4} → {1,2;3,4}.
 * Çevrilemeyen formülde undefined döner; okuyucu önbellekteki değeri kullanır.
 */
import { EN_SYNTAX, FUNCTIONS, colToName, nameToCol, parse, quoteSheetName, tokenize, type RefPart, type RefToken } from "@officehere/formula";

/** OpenFormula'da farklı adlandırılan işlevler (okuma yönü). */
const OF_TO_XL: Record<string, string> = {
  "ERRORTYPE": "ERROR.TYPE",
  "USDOLLAR": "DOLLAR",
  "CEILING.XCL": "CEILING",
  "FLOOR.XCL": "FLOOR",
  "CHISQDIST": "CHIDIST",
};

/** Yazarken önek isteyen eski/uyumluluk işlevleri. */
const LEGACY = new Set(["NORMSDIST", "NORMSINV", "CHIDIST", "CHIINV", "FDIST", "FINV", "TDIST", "TINV", "CHITEST", "FTEST"]);
/** Yazarken "COM.MICROSOFT." öneki: anlamı OpenFormula'dakinden farklı olanlar. */
const MS_PREFIX = new Set(["CEILING", "FLOOR"]);

/** Dosyadaki sayfa adı → modeldeki (düzeltilmiş) ad. */
export type SheetMap = (raw: string) => string;

function cellA1(colAbs: boolean, col: number, rowAbs: boolean, row: number): string {
  return `${colAbs ? "$" : ""}${colToName(col)}${rowAbs ? "$" : ""}${row + 1}`;
}

interface OfPart {
  sheet?: string;
  col?: number;
  row?: number;
  colAbs: boolean;
  rowAbs: boolean;
}

/** "[$Sayfa.$A$1" iç parçası: [$]['Ad'|Ad].[$]COL[$]ROW ya da yalnız sütun/satır. */
function parseOfPart(s: string): OfPart | null {
  let rest = s;
  let sheet: string | undefined;
  const dot = (() => {
    // Sayfa adı tırnaklı olabilir; tırnak dışındaki son nokta ayırıcıdır
    let q = false;
    let last = -1;
    for (let i = 0; i < rest.length; i++) {
      const c = rest[i];
      if (c === "'") {
        if (q && rest[i + 1] === "'") {
          i++;
          continue;
        }
        q = !q;
      } else if (c === "." && !q) last = i;
    }
    return last;
  })();
  if (dot < 0) return null;
  const sh = rest.slice(0, dot).replace(/^\$/, "");
  rest = rest.slice(dot + 1);
  if (sh) {
    if (sh.startsWith("'")) {
      if (!sh.endsWith("'")) return null;
      sheet = sh.slice(1, -1).replace(/''/g, "'");
    } else sheet = sh;
    if (sheet.includes("#")) return null; // dış dosya ya da hata
  }
  const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d{1,7})?$/.exec(rest);
  if (!m || (!m[2] && !m[4])) return null;
  const part: OfPart = { sheet, colAbs: m[1] === "$", rowAbs: m[3] === "$" };
  if (m[2]) part.col = nameToCol(m[2].toUpperCase());
  if (m[4]) part.row = Number(m[4]) - 1;
  if (!m[2] && m[1] === "$") part.rowAbs = true;
  return part;
}

/** [ ... ] başvurusunu A1 metnine çevirir. */
export function ofRefToA1(inner: string, mapSheet: SheetMap = (s) => s): string | null {
  if (/#REF!/i.test(inner)) return "#REF!";
  if (inner.includes("#") || inner.includes("'file:")) return null;
  // Aralık ayırıcısı: tırnak dışındaki ":"
  let q = false;
  let colon = -1;
  for (let i = 0; i < inner.length; i++) {
    if (inner[i] === "'") q = !q;
    else if (inner[i] === ":" && !q) {
      colon = i;
      break;
    }
  }
  const a = parseOfPart(colon < 0 ? inner : inner.slice(0, colon));
  if (!a) return null;
  const b = colon < 0 ? null : parseOfPart(inner.slice(colon + 1));
  if (colon >= 0 && !b) return null;
  // Üç boyutlu (sayfalar arası) aralık desteklenmez
  if (b?.sheet && a.sheet && b.sheet !== a.sheet) return null;
  const prefix = a.sheet ? quoteSheetName(mapSheet(a.sheet)) + "!" : "";
  const one = (p: OfPart) => {
    if (p.col !== undefined && p.row !== undefined) return cellA1(p.colAbs, p.col, p.rowAbs, p.row);
    if (p.col !== undefined) return `${p.colAbs ? "$" : ""}${colToName(p.col)}`;
    return `${p.rowAbs ? "$" : ""}${(p.row ?? 0) + 1}`;
  };
  if (!b) {
    if (a.col === undefined || a.row === undefined) return null;
    return prefix + one(a);
  }
  // Sütun/satır aralıkları her iki uçta aynı türde olmalı
  const kind = (p: OfPart) => (p.col !== undefined && p.row !== undefined ? "cell" : p.col !== undefined ? "col" : "row");
  if (kind(a) !== kind(b)) return null;
  return `${prefix}${one(a)}:${one(b)}`;
}

/** İşlev adını motorun kanonik adına çevirir; bilinmiyorsa undefined. */
export function ofFunctionName(raw: string): string | undefined {
  let n = raw.toUpperCase().replace(/^(COM\.MICROSOFT\.|_XLFN\.|_XLWS\.|LEGACY\.|ORG\.OPENOFFICE\.|ORG\.LIBREOFFICE\.)+/, "");
  n = OF_TO_XL[n] ?? n;
  return FUNCTIONS.has(n) ? n : undefined;
}

/**
 * "of:=SUM([.A1:.B2])" → "=SUM(A1:B2)". Desteklenmeyen öğede (birleşim "~",
 * dış başvuru, bilinmeyen işlev, 3B aralık) undefined döner.
 */
export function odfFormulaToA1(src: string, mapSheet?: SheetMap): string | undefined {
  let f = src.trim();
  const ns = /^([a-zA-Z]+):(?!\w*\])/.exec(f);
  if (ns) {
    const n = ns[1].toLowerCase();
    if (n === "msoxl") {
      const body = f.slice(ns[0].length).replace(/^=/, "");
      return validate("=" + body.replace(/_xlfn\.|_xlws\./g, ""));
    }
    if (n !== "of" && n !== "oooc" && n !== "ooow") return undefined;
    f = f.slice(ns[0].length);
  }
  f = f.replace(/^=/, "");
  let out = "";
  let i = 0;
  // Parantez/küme yığını: "f" işlev çağrısı, "p" parantez, "a" dizi sabiti
  const stack: string[] = [];
  while (i < f.length) {
    const c = f[i];
    if (c === '"') {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === '"') {
          if (f[j + 1] === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      if (j >= f.length) return undefined;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === "[") {
      let j = i + 1;
      let q = false;
      while (j < f.length && (f[j] !== "]" || q)) {
        if (f[j] === "'") q = !q;
        j++;
      }
      if (j >= f.length) return undefined;
      const a1 = ofRefToA1(f.slice(i + 1, j), mapSheet);
      if (a1 === null) return undefined;
      out += a1;
      i = j + 1;
      continue;
    }
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === "~") {
      // Birleşim yalnız açık parantez içinde virgülle yazılabilir: (A1,B2)
      if (stack[stack.length - 1] !== "p") return undefined;
      out += ",";
      i++;
      continue;
    }
    if (c === "!") {
      out += " ";
      i++;
      continue;
    }
    if (c === ";") {
      out += ",";
      i++;
      continue;
    }
    if (c === "|") {
      if (stack[stack.length - 1] !== "a") return undefined;
      out += ";";
      i++;
      continue;
    }
    if (c === "{") {
      stack.push("a");
      out += c;
      i++;
      continue;
    }
    if (c === "}") {
      stack.pop();
      out += c;
      i++;
      continue;
    }
    if (c === "(") {
      stack.push("p");
      out += c;
      i++;
      continue;
    }
    if (c === ")") {
      stack.pop();
      out += c;
      i++;
      continue;
    }
    if (c === "#") {
      const m = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)/i.exec(f.slice(i));
      if (!m) return undefined;
      out += m[0].toUpperCase();
      i += m[0].length;
      continue;
    }
    const id = /^[\p{L}_][\p{L}\p{N}_.]*/u.exec(f.slice(i));
    if (id) {
      const name = id[0];
      let j = i + name.length;
      while (j < f.length && /\s/.test(f[j])) j++;
      if (f[j] === "(") {
        const fnName = ofFunctionName(name);
        const up = name.toUpperCase();
        // TRUE()/FALSE() OpenFormula'da işlevdir; motorda da işlev olarak vardır
        if (!fnName && up !== "TRUE" && up !== "FALSE") return undefined;
        out += (fnName ?? up) + "(";
        stack.push("f");
        i = j + 1;
        continue;
      }
      // Ad (adlandırılmış aralık) ya da mantıksal sabit
      out += name;
      i += name.length;
      continue;
    }
    const num = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(f.slice(i));
    if (num) {
      out += num[0];
      i += num[0].length;
      continue;
    }
    if ("+-*/^&%=<>:".includes(c)) {
      out += c;
      i++;
      continue;
    }
    return undefined;
  }
  return validate("=" + out);
}

function validate(f: string): string | undefined {
  try {
    parse(f.slice(1));
    // Bilinmeyen işlev adlarını da reddet
    for (const t of tokenize(f.slice(1), EN_SYNTAX)) if (t.type === "func" && !FUNCTIONS.has(String(t.value ?? t.text).toUpperCase().replace(/^_XLFN\./, ""))) return undefined;
    return f;
  } catch {
    return undefined;
  }
}

/* ------------------------------ Yazma ------------------------------ */

function ofSheet(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function ofPart(p: RefPart, kind: RefToken["kind"]): string {
  if (kind === "cols") return `${p.colAbs ? "$" : ""}${colToName(p.col)}`;
  if (kind === "rows") return `${p.rowAbs ? "$" : ""}${p.row + 1}`;
  return cellA1(p.colAbs, p.col, p.rowAbs, p.row);
}

function refToOf(ref: RefToken): string {
  if (ref.refError) return "[.#REF!]";
  const sh = ref.sheet ? "$" + ofSheet(ref.sheet) : "";
  const a = `${sh}.${ofPart(ref.a, ref.kind)}`;
  if (!ref.b) return `[${a}]`;
  return `[${a}:.${ofPart(ref.b, ref.kind)}]`;
}

/** OpenFormula'daki işlev adı. */
function functionToOf(name: string): string {
  const up = name.toUpperCase().replace(/^_XLFN\./, "");
  if (LEGACY.has(up)) return "LEGACY." + up;
  if (MS_PREFIX.has(up) || FUNCTIONS.get(up)?.xlfn) return "COM.MICROSOFT." + up;
  return up;
}

/**
 * Kanonik formül ("=SUM(A1:B2)") → "of:=SUM([.A1:.B2])". Yapısal (tablo)
 * başvurular OpenFormula'da yazılamaz: undefined döner.
 */
export function a1FormulaToOdf(formula: string): string | undefined {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const toks = tokenize(src, EN_SYNTAX);
  let out = "";
  const stack: ("f" | "p" | "a")[] = [];
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    switch (t.type) {
      case "ref":
        out += refToOf(t.ref!);
        break;
      case "structref":
      case "unknown":
        return undefined;
      case "func":
        out += functionToOf(String(t.value ?? t.text));
        break;
      case "lparen":
        stack.push(toks[k - 1]?.type === "func" ? "f" : "p");
        out += "(";
        break;
      case "rparen":
        stack.pop();
        out += ")";
        break;
      case "lbrace":
        stack.push("a");
        out += "{";
        break;
      case "rbrace":
        stack.pop();
        out += "}";
        break;
      case "sep":
        // İşlev dışındaki parantezde virgül başvuru birleşimidir
        out += stack[stack.length - 1] === "p" ? "~" : ";";
        break;
      case "arrayCol":
        out += ";";
        break;
      case "arrayRow":
        out += "|";
        break;
      case "bool":
        out += t.value ? "TRUE()" : "FALSE()";
        break;
      case "ws": {
        // İki başvuru arasındaki boşluk kesişim işlecidir
        const prev = toks.slice(0, k).reverse().find((x) => x.type !== "ws");
        const next = toks.slice(k + 1).find((x) => x.type !== "ws");
        if (prev && next && ["ref", "name", "rparen"].includes(prev.type) && ["ref", "name", "func", "lparen"].includes(next.type)) out += "!";
        break;
      }
      case "op":
        // Örtük kesişim (@) OpenFormula'da yok
        if (t.text !== "@") out += t.text;
        break;
      default:
        out += t.text;
    }
  }
  return "of:=" + out;
}
