/**
 * Formül sözcük çözümleyicisi. Belirteçler kaynak metni (text) korur;
 * böylece başvuru kaydırma ve yerelleştirme boşlukları bozmadan
 * belirteç düzeyinde yapılabilir.
 */
import { MAX_COLS, MAX_ROWS, nameToCol } from "./references";

export interface SyntaxLocale {
  /** Ondalık ayırıcı. */
  decimal: "." | ",";
  /** Bağımsız değişken ayırıcı. */
  argSep: "," | ";";
  /** Dizi sabitinde sütun ayırıcı. */
  arrayCol: string;
  /** Dizi sabitinde satır ayırıcı. */
  arrayRow: string;
}

export const EN_SYNTAX: SyntaxLocale = { decimal: ".", argSep: ",", arrayCol: ",", arrayRow: ";" };
export const TR_SYNTAX: SyntaxLocale = { decimal: ",", argSep: ";", arrayCol: "\\", arrayRow: ";" };

export type TokenType =
  | "number"
  | "string"
  | "bool"
  | "error"
  | "ref"
  | "func"
  | "name"
  | "op"
  | "lparen"
  | "rparen"
  | "lbrace"
  | "rbrace"
  | "sep"
  | "arrayCol"
  | "arrayRow"
  | "ws"
  | "unknown";

export interface RefPart {
  /** -1: yalnız sütun (A:A). */
  row: number;
  /** -1: yalnız satır (1:3). */
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export interface RefToken {
  sheet?: string;
  kind: "cell" | "range" | "cols" | "rows";
  a: RefPart;
  b?: RefPart;
  /** Silinmiş başvuru: "Sayfa1!#REF!". */
  refError?: boolean;
}

export interface Token {
  type: TokenType;
  text: string;
  pos: number;
  value?: number | string | boolean;
  ref?: RefToken;
}

const TR_ERROR_NAMES: Record<string, string> = {
  "#NULL!": "#BOŞ!",
  "#DIV/0!": "#SAYI/0!",
  "#VALUE!": "#DEĞER!",
  "#REF!": "#BAŞV!",
  "#NAME?": "#AD?",
  "#NUM!": "#SAYI!",
  "#N/A": "#YOK",
  "#SPILL!": "#TAŞMA!",
  "#CALC!": "#HESAPLA!",
  "#GETTING_DATA": "#VERİ_ALINIYOR",
};

export const ERROR_TO_TR = TR_ERROR_NAMES;
export const ERROR_FROM_TR: Record<string, string> = Object.fromEntries(Object.entries(TR_ERROR_NAMES).map(([k, v]) => [v, k]));

const ALL_ERRORS = [...Object.keys(TR_ERROR_NAMES), ...Object.values(TR_ERROR_NAMES)].sort((a, b) => b.length - a.length);

const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const PLAIN_SHEET = /([\p{L}_][\p{L}\p{N}_.]*)!/uy;
const CELL_RANGE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/y;
const COL_RANGE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const ROW_RANGE = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;
const IDENT = /[\p{L}_\\][\p{L}\p{N}_.?]*/uy;
const NUM_EN = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
const NUM_TR = /(?:\d+(?:,\d*)?|,\d+)(?:[eE][+-]?\d+)?/y;
const IDENT_CHAR = /[\p{L}\p{N}_.(]/u;

export const BOOL_NAMES: Record<string, boolean> = { TRUE: true, FALSE: false, DOĞRU: true, YANLIŞ: false };

function upperBoth(s: string): string[] {
  return [s.toUpperCase(), s.toLocaleUpperCase("tr-TR")];
}

function stickyMatch(re: RegExp, src: string, pos: number): RegExpExecArray | null {
  re.lastIndex = pos;
  return re.exec(src);
}

function cellPart(colAbs: string, col: string, rowAbs: string, row: string): RefPart | null {
  const c = nameToCol(col);
  const r = parseInt(row, 10) - 1;
  if (c < 0 || r < 0 || r >= MAX_ROWS) return null;
  return { row: r, col: c, rowAbs: rowAbs === "$", colAbs: colAbs === "$" };
}

/** Konumdaki başvuru gövdesini (sayfa öneki olmadan) okur. */
function matchRefBody(src: string, pos: number): { ref: RefToken; end: number } | null {
  const after = (end: number) => end >= src.length || !IDENT_CHAR.test(src[end]);
  let m = stickyMatch(CELL_RANGE, src, pos);
  if (m && after(pos + m[0].length)) {
    const a = cellPart(m[1], m[2], m[3], m[4]);
    if (a) {
      if (m[6] !== undefined) {
        const b = cellPart(m[5], m[6], m[7], m[8]);
        if (b) return { ref: { kind: "range", a, b }, end: pos + m[0].length };
      } else {
        return { ref: { kind: "cell", a }, end: pos + m[0].length };
      }
    }
  }
  // "A1:B" gibi eksik eşleşme yerine yalnız "A1" kısmını dene.
  if (m && m[6] !== undefined) {
    const single = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
    const s = stickyMatch(single, src, pos);
    if (s && src[pos + s[0].length] === ":") {
      const a = cellPart(s[1], s[2], s[3], s[4]);
      if (a) return { ref: { kind: "cell", a }, end: pos + s[0].length };
    }
  }
  m = stickyMatch(COL_RANGE, src, pos);
  if (m && after(pos + m[0].length)) {
    const c1 = nameToCol(m[2]);
    const c2 = nameToCol(m[4]);
    if (c1 >= 0 && c2 >= 0 && c1 < MAX_COLS && c2 < MAX_COLS) {
      return {
        ref: { kind: "cols", a: { row: -1, col: c1, rowAbs: false, colAbs: m[1] === "$" }, b: { row: -1, col: c2, rowAbs: false, colAbs: m[3] === "$" } },
        end: pos + m[0].length,
      };
    }
  }
  m = stickyMatch(ROW_RANGE, src, pos);
  if (m && after(pos + m[0].length)) {
    const r1 = parseInt(m[2], 10) - 1;
    const r2 = parseInt(m[4], 10) - 1;
    if (r1 >= 0 && r2 >= 0 && r1 < MAX_ROWS && r2 < MAX_ROWS) {
      return {
        ref: { kind: "rows", a: { row: r1, col: -1, rowAbs: m[1] === "$", colAbs: false }, b: { row: r2, col: -1, rowAbs: m[3] === "$", colAbs: false } },
        end: pos + m[0].length,
      };
    }
  }
  return null;
}

export function tokenize(src: string, syntax: SyntaxLocale = EN_SYNTAX): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  let braceDepth = 0;
  const push = (type: TokenType, start: number, end: number, extra?: Partial<Token>) => {
    tokens.push({ type, text: src.slice(start, end), pos: start, ...extra });
    i = end;
  };
  const numRe = syntax.decimal === "," ? NUM_TR : NUM_EN;

  while (i < src.length) {
    const ch = src[i];
    // Boşluk
    if (/\s/.test(ch)) {
      let j = i + 1;
      while (j < src.length && /\s/.test(src[j])) j++;
      push("ws", i, j);
      continue;
    }
    // Metin
    if (ch === '"') {
      let j = i + 1;
      let s = "";
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      if (j >= src.length) {
        push("unknown", i, src.length);
        continue;
      }
      push("string", i, j + 1, { value: s });
      continue;
    }
    // Dizi sabiti ayırıcıları
    if (braceDepth > 0) {
      if (ch === syntax.arrayRow) {
        push("arrayRow", i, i + 1);
        continue;
      }
      if (ch === syntax.arrayCol || (syntax.decimal === "," && (ch === "." || ch === "\\")) || (syntax.decimal === "." && ch === ",")) {
        push("arrayCol", i, i + 1);
        continue;
      }
    }
    // Hata değerleri
    if (ch === "#") {
      const rest = src.slice(i).toLocaleUpperCase("tr-TR");
      const restEn = src.slice(i).toUpperCase();
      const e = ALL_ERRORS.find((x) => rest.startsWith(x) || restEn.startsWith(x));
      if (e) {
        const canonical = ERROR_FROM_TR[e] ?? e;
        if (canonical === "#REF!") push("ref", i, i + e.length, { ref: { kind: "cell", a: { row: 0, col: 0, rowAbs: false, colAbs: false }, refError: true } });
        else push("error", i, i + e.length, { value: canonical });
        continue;
      }
      push("unknown", i, i + 1);
      continue;
    }
    // Sayfa önekli başvuru
    if (ch === "'") {
      const m = stickyMatch(QUOTED_SHEET, src, i);
      if (m) {
        const sheet = m[1].replace(/''/g, "'");
        const bodyPos = i + m[0].length;
        const body = matchRefBody(src, bodyPos);
        if (body) {
          push("ref", i, body.end, { ref: { ...body.ref, sheet } });
          continue;
        }
        const errM = /#REF!|#BAŞV!/iy;
        errM.lastIndex = bodyPos;
        const em = errM.exec(src);
        if (em) {
          push("ref", i, bodyPos + em[0].length, { ref: { kind: "cell", sheet, a: { row: 0, col: 0, rowAbs: false, colAbs: false }, refError: true } });
          continue;
        }
      }
      push("unknown", i, i + 1);
      continue;
    }
    {
      const sm = stickyMatch(PLAIN_SHEET, src, i);
      if (sm) {
        const bodyPos = i + sm[0].length;
        const body = matchRefBody(src, bodyPos);
        if (body) {
          push("ref", i, body.end, { ref: { ...body.ref, sheet: sm[1] } });
          continue;
        }
        const errM = /#REF!|#BAŞV!/iy;
        errM.lastIndex = bodyPos;
        const em = errM.exec(src);
        if (em) {
          push("ref", i, bodyPos + em[0].length, { ref: { kind: "cell", sheet: sm[1], a: { row: 0, col: 0, rowAbs: false, colAbs: false }, refError: true } });
          continue;
        }
      }
    }
    // Başvuru (sayfasız)
    if (ch === "$" || /[A-Za-z0-9]/.test(ch)) {
      const body = matchRefBody(src, i);
      if (body) {
        push("ref", i, body.end, { ref: body.ref });
        continue;
      }
    }
    // Sayı
    if (/[0-9]/.test(ch) || (ch === syntax.decimal && /[0-9]/.test(src[i + 1] ?? ""))) {
      const m = stickyMatch(numRe, src, i);
      if (m) {
        const v = parseFloat(m[0].replace(",", "."));
        push("number", i, i + m[0].length, { value: v });
        continue;
      }
    }
    // Tanımlayıcı: fonksiyon, ad, mantıksal
    {
      const m = stickyMatch(IDENT, src, i);
      if (m) {
        const end = i + m[0].length;
        if (src[end] === "(") {
          push("func", i, end, { value: m[0] });
          continue;
        }
        const [u1, u2] = upperBoth(m[0]);
        if (u1 in BOOL_NAMES || u2 in BOOL_NAMES) {
          push("bool", i, end, { value: BOOL_NAMES[u1] ?? BOOL_NAMES[u2] });
          continue;
        }
        push("name", i, end, { value: m[0] });
        continue;
      }
    }
    // İşleçler
    const two = src.slice(i, i + 2);
    if (two === "<>" || two === "<=" || two === ">=") {
      push("op", i, i + 2);
      continue;
    }
    if ("+-*/^&=<>%:@".includes(ch)) {
      push("op", i, i + 1);
      continue;
    }
    if (ch === "(") {
      push("lparen", i, i + 1);
      continue;
    }
    if (ch === ")") {
      push("rparen", i, i + 1);
      continue;
    }
    if (ch === "{") {
      braceDepth++;
      push("lbrace", i, i + 1);
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      push("rbrace", i, i + 1);
      continue;
    }
    if (braceDepth > 0) {
      if (ch === syntax.arrayRow) {
        push("arrayRow", i, i + 1);
        continue;
      }
      if (ch === syntax.arrayCol || (syntax.decimal === "," && (ch === "." || ch === "\\")) || (syntax.decimal === "." && ch === ",")) {
        push("arrayCol", i, i + 1);
        continue;
      }
    }
    if (ch === syntax.argSep || ch === "," || ch === ";") {
      push("sep", i, i + 1);
      continue;
    }
    push("unknown", i, i + 1);
  }
  return tokens;
}

/** Belirteçleri yeniden metne çevirir (başvuru belirteçleri yapıdan üretilir). */
export function tokensToString(tokens: Token[]): string {
  return tokens.map((t) => t.text).join("");
}
