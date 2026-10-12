/** Arama ve başvuru fonksiyonları. */
import { compareScalars, hasWildcards, textEquals, wildcardToRegExp } from "../coerce";
import type { FnContext } from "../context";
import { tokenize, EN_SYNTAX } from "../lexer";
import { MAX_COLS, MAX_ROWS, addr, colToName, quoteSheetName } from "../references";
import { Errors, isArray, isError, isMultiRef, isRef, makeArray, type RefValue, type Scalar, type Value } from "../values";
import { clipRef, dims, isMulti, matrix, num, optBool, optNum, optText, scalar, text } from "./helpers";
import { fn, lazyFn, type FnDef } from "./registry";

/** Arama vektörü: başvuru ise hücreleri tembel okur. */
interface Vec {
  length: number;
  get(i: number): Scalar;
}

function vectorOf(v: Value, c: FnContext, clip = true): { vec: Vec; horizontal: boolean } {
  if (isRef(v)) {
    const r = clip ? clipRef(v, c) ?? { ...v, r2: v.r1 - 1, c2: v.c1 - 1 } : v;
    const rows = Math.max(0, r.r2 - r.r1 + 1);
    const cols = Math.max(0, r.c2 - r.c1 + 1);
    const fullRows = v.r2 - v.r1 + 1;
    const fullCols = v.c2 - v.c1 + 1;
    if (fullRows === 1 && fullCols >= 1) return { vec: { length: cols, get: (i) => c.host.getCellValue(v.sheet, v.r1, v.c1 + i) }, horizontal: true };
    if (fullCols === 1) return { vec: { length: rows, get: (i) => c.host.getCellValue(v.sheet, v.r1 + i, v.c1) }, horizontal: false };
    throw Errors.NA;
  }
  const m = matrix(v, c);
  if (m.length === 1) return { vec: { length: m[0].length, get: (i) => m[0][i] }, horizontal: true };
  if ((m[0]?.length ?? 0) === 1) return { vec: { length: m.length, get: (i) => m[i][0] }, horizontal: false };
  throw Errors.NA;
}

function sameKind(a: Scalar, b: Scalar): boolean {
  if (a === null || b === null) return false;
  return typeof a === typeof b;
}

function rank(v: Scalar): number {
  if (typeof v === "number") return 0;
  if (typeof v === "string") return 1;
  if (typeof v === "boolean") return 2;
  return 3;
}

/** Tam eşleşme (joker isteğe bağlı). */
function exactMatcher(x: Scalar, wildcard: boolean): (v: Scalar) => boolean {
  if (typeof x === "string") {
    if (wildcard && (hasWildcards(x) || x.includes("~"))) {
      const re = wildcardToRegExp(x);
      return (v) => typeof v === "string" && re.test(v);
    }
    return (v) => typeof v === "string" && textEquals(v, x);
  }
  if (typeof x === "number") return (v) => typeof v === "number" && compareScalars(v, x) === 0;
  if (typeof x === "boolean") return (v) => v === x;
  if (x === null) return (v) => v === null || v === 0;
  return () => false;
}

function lookupKey(v: Scalar): string | null {
  if (typeof v === "number") return "n" + v;
  if (typeof v === "string") return "s" + v.toLocaleUpperCase("tr-TR");
  if (typeof v === "boolean") return "b" + v;
  return null;
}

/** Dizin önbelleği: aynı yeniden hesaplamada tekrarlanan tam aramaları hızlandırır. */
function exactIndex(vecId: string | null, vec: Vec, x: Scalar, c: FnContext): number {
  const cache = (c.host as { lookupCache?: Map<string, Map<string, number>> }).lookupCache;
  const key = lookupKey(x);
  if (cache && vecId && key && !(typeof x === "string" && (hasWildcards(x) || x.includes("~")))) {
    let idx = cache.get(vecId);
    if (!idx) {
      idx = new Map();
      for (let i = vec.length - 1; i >= 0; i--) {
        const k = lookupKey(vec.get(i));
        if (k) idx.set(k, i);
      }
      cache.set(vecId, idx);
    }
    return idx.get(key) ?? -1;
  }
  const m = exactMatcher(x, true);
  for (let i = 0; i < vec.length; i++) if (m(vec.get(i))) return i;
  return -1;
}

function refId(v: Value, horizontal: boolean): string | null {
  if (!isRef(v)) return null;
  return `${v.sheet}:${v.r1}:${v.c1}:${horizontal ? "h" : "v"}:${horizontal ? v.c2 : v.r2}`;
}

/** Sıralı vektörde ikili arama: x'ten küçük-eşit son öğe (asc) ya da büyük-eşit son öğe (desc). */
function binarySearch(vec: Vec, x: Scalar, desc = false): number {
  let lo = 0;
  let hi = vec.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = vec.get(mid);
    if (sameKind(v, x)) {
      const cmp = compareScalars(v, x);
      if (desc ? cmp >= 0 : cmp <= 0) {
        best = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    } else if (v === null || isError(v)) {
      // Boşlukları atla: sola doğru en yakın dolu öğeyi bul
      let k = mid - 1;
      while (k >= lo && (vec.get(k) === null || isError(vec.get(k)))) k--;
      if (k < lo) lo = mid + 1;
      else {
        const vk = vec.get(k);
        const cmp = sameKind(vk, x) ? compareScalars(vk, x) : rank(vk) - rank(x);
        if (desc ? cmp >= 0 : cmp <= 0) {
          if (sameKind(vk, x)) best = k;
          lo = mid + 1;
        } else hi = k - 1;
      }
    } else {
      if (desc ? rank(v) > rank(x) : rank(v) < rank(x)) lo = mid + 1;
      else hi = mid - 1;
    }
  }
  return best;
}

function vlookup(a: Value[], c: FnContext, horizontal: boolean): Value {
  const x = scalar(a[0], c);
  if (isError(x)) return x;
  const table = a[1];
  const idx = Math.trunc(num(a[2], c));
  const approx = optBool(a[3], true, c);
  if (idx < 1) throw Errors.VALUE;
  if (isRef(table)) {
    const w = horizontal ? table.r2 - table.r1 + 1 : table.c2 - table.c1 + 1;
    if (idx > w) throw Errors.REF;
    const first: RefValue = horizontal ? { ...table, r2: table.r1 } : { ...table, c2: table.c1 };
    const { vec } = vectorOf(first, c);
    const i = approx ? binarySearch(vec, x) : exactIndex(refId(first, horizontal), vec, x, c);
    if (i < 0) return Errors.NA;
    return horizontal ? c.host.getCellValue(table.sheet, table.r1 + idx - 1, table.c1 + i) : c.host.getCellValue(table.sheet, table.r1 + i, table.c1 + idx - 1);
  }
  const m = matrix(table, c);
  const rows = horizontal ? m[0]?.length ?? 0 : m.length;
  const w = horizontal ? m.length : m[0]?.length ?? 0;
  if (idx > w) throw Errors.REF;
  const vec: Vec = { length: rows, get: (i) => (horizontal ? m[0][i] : m[i][0]) };
  const i = approx ? binarySearch(vec, x) : exactIndex(null, vec, x, c);
  if (i < 0) return Errors.NA;
  return horizontal ? m[idx - 1][i] : m[i][idx - 1];
}

function matchImpl(x: Scalar, arr: Value, type: number, c: FnContext): number {
  const { vec, horizontal } = vectorOf(arr, c);
  let i: number;
  if (type === 0) i = exactIndex(refId(arr, horizontal), vec, x, c);
  else if (type > 0) i = binarySearch(vec, x);
  else i = binarySearch(vec, x, true);
  if (i < 0) throw Errors.NA;
  return i + 1;
}

/** ÇAPRAZEŞLEŞTİR/ÇAPRAZARA ortak arama. */
function xsearch(x: Scalar, vec: Vec, matchMode: number, searchMode: number, vecId: string | null, c: FnContext): number {
  if (searchMode === 2 || searchMode === -2) {
    const desc = searchMode === -2;
    // İkili arama
    let lo = 0;
    let hi = vec.length - 1;
    let found = -1;
    let lower = -1;
    let upper = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = vec.get(mid);
      let cmp = sameKind(v, x) ? compareScalars(v, x) : rank(v) - rank(x);
      if (desc) cmp = -cmp;
      if (cmp === 0) {
        found = mid;
        break;
      }
      if (cmp < 0) {
        desc ? (upper = mid) : (lower = mid);
        lo = mid + 1;
      } else {
        desc ? (lower = mid) : (upper = mid);
        hi = mid - 1;
      }
    }
    if (found >= 0) return found;
    if (matchMode === -1) return lower;
    if (matchMode === 1) return upper;
    return -1;
  }
  const order = searchMode === -1 ? Array.from({ length: vec.length }, (_, i) => vec.length - 1 - i) : null;
  const at = (k: number) => (order ? order[k] : k);
  if (matchMode === 0 && searchMode === 1) return exactIndexNoWild(vec, x, c, vecId);
  const exact = exactMatcher(x, matchMode === 2);
  let best = -1;
  let bestV: Scalar = null;
  for (let k = 0; k < vec.length; k++) {
    const i = at(k);
    const v = vec.get(i);
    if (exact(v)) return i;
    if ((matchMode === -1 || matchMode === 1) && sameKind(v, x)) {
      const cmp = compareScalars(v, x);
      if (matchMode === -1 && cmp < 0 && (best < 0 || compareScalars(v, bestV) > 0)) {
        best = i;
        bestV = v;
      }
      if (matchMode === 1 && cmp > 0 && (best < 0 || compareScalars(v, bestV) < 0)) {
        best = i;
        bestV = v;
      }
    }
  }
  return best;
}

function exactIndexNoWild(vec: Vec, x: Scalar, c: FnContext, vecId: string | null): number {
  if (typeof x === "string" && (hasWildcards(x) || x.includes("~"))) {
    for (let i = 0; i < vec.length; i++) {
      const v = vec.get(i);
      if (typeof v === "string" && textEquals(v, x)) return i;
    }
    return -1;
  }
  return exactIndex(vecId, vec, x, c);
}

function xlookupOne(x: Scalar, a: Value[], c: FnContext): Value {
  const look = a[1];
  const ret = a[2];
  const matchMode = Math.trunc(optNum(a[4], 0, c));
  const searchMode = Math.trunc(optNum(a[5], 1, c));
  const { vec, horizontal } = vectorOf(look, c, false);
  const i = xsearch(x, vec, matchMode, searchMode, refId(look, horizontal), c);
  if (i < 0) {
    if (a.length > 3 && a[3] !== null) return a[3];
    return Errors.NA;
  }
  const rd = dims(ret);
  if (horizontal) {
    if (rd.cols !== vec.length) throw Errors.VALUE;
    if (isRef(ret)) return { ...ret, c1: ret.c1 + i, c2: ret.c1 + i };
    const m = matrix(ret, c);
    return makeArray(m.map((r) => [r[i]]));
  }
  if (rd.rows !== vec.length) throw Errors.VALUE;
  if (isRef(ret)) return { ...ret, r1: ret.r1 + i, r2: ret.r1 + i };
  const m = matrix(ret, c);
  return makeArray([m[i]]);
}

function indexImpl(a: Value[], c: FnContext): Value {
  let src = a[0];
  if (isMultiRef(src)) {
    const area = Math.trunc(optNum(a[3], 1, c));
    if (area < 1 || area > src.areas.length) throw Errors.REF;
    src = src.areas[area - 1];
  }
  const d = dims(src);
  let r = a.length > 1 && a[1] !== null ? Math.trunc(num(a[1], c)) : 0;
  let col = a.length > 2 && a[2] !== null ? Math.trunc(num(a[2], c)) : 0;
  if (a.length === 2 && d.rows === 1 && d.cols > 1) {
    col = r;
    r = 1;
  } else if (a.length === 2 && d.cols === 1) col = 1;
  if (r < 0 || col < 0 || r > d.rows || col > d.cols) throw Errors.REF;
  if (isRef(src)) {
    const r1 = r === 0 ? src.r1 : src.r1 + r - 1;
    const r2 = r === 0 ? src.r2 : r1;
    const c1 = col === 0 ? src.c1 : src.c1 + col - 1;
    const c2 = col === 0 ? src.c2 : c1;
    return { kind: "ref", sheet: src.sheet, r1, r2, c1, c2 };
  }
  const m = matrix(src, c);
  if (r === 0 && col === 0) return makeArray(m);
  if (r === 0) return makeArray(m.map((row) => [row[col - 1]]));
  if (col === 0) return makeArray([m[r - 1]]);
  return m[r - 1][col - 1];
}

/** DOLAYLI: başvuru metnini çözer. */
function indirect(a: Value[], c: FnContext): Value {
  const s = text(a[0], c).trim();
  const a1 = optBool(a[1], true, c);
  let target: Value = Errors.REF;
  if (!a1) {
    target = parseR1C1(s, c);
  } else {
    const toks = tokenize(s, EN_SYNTAX).filter((t) => t.type !== "ws");
    if (toks.length === 1 && toks[0].type === "ref") {
      target = c.evalNode({ type: "ref", ref: toks[0].ref! });
    } else if (toks.length === 3 && toks[0].type === "ref" && toks[1].text === ":" && toks[2].type === "ref") {
      target = c.evalNode({ type: "binary", op: ":", left: { type: "ref", ref: toks[0].ref! }, right: { type: "ref", ref: toks[2].ref! } });
    } else if (toks.length === 1 && toks[0].type === "name") {
      const def = c.host.getDefinedName(toks[0].value as string, c.sheet);
      if (def) target = c.evalNode(def);
    }
  }
  if (isRef(target)) c.onDynamicRef?.(target);
  return isRef(target) || isMultiRef(target) ? target : Errors.REF;
}

function parseR1C1(s: string, c: FnContext): Value {
  let sheet = c.sheet;
  let body = s;
  const bang = s.lastIndexOf("!");
  if (bang >= 0) {
    let name = s.slice(0, bang);
    if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
    const id = c.host.sheetIdByName(name);
    if (id === undefined) return Errors.REF;
    sheet = id;
    body = s.slice(bang + 1);
  }
  const part = /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i;
  const parsePart = (p: string): [number, number] | null => {
    const m = part.exec(p);
    if (!m) return null;
    const conv = (t: string | undefined, base: number) => (t === undefined ? base : t.startsWith("[") ? base + parseInt(t.slice(1, -1), 10) : parseInt(t, 10) - 1);
    return [conv(m[1], c.row), conv(m[2], c.col)];
  };
  const pieces = body.split(":");
  const p1 = parsePart(pieces[0]);
  const p2 = pieces[1] ? parsePart(pieces[1]) : p1;
  if (!p1 || !p2) return Errors.REF;
  return { kind: "ref", sheet, r1: Math.min(p1[0], p2[0]), c1: Math.min(p1[1], p2[1]), r2: Math.max(p1[0], p2[0]), c2: Math.max(p1[1], p2[1]) };
}

function offset(a: Value[], c: FnContext): Value {
  const base = a[0];
  if (!isRef(base)) throw Errors.VALUE;
  const dr = Math.trunc(num(a[1], c));
  const dc = Math.trunc(num(a[2], c));
  const h = a.length > 3 && a[3] !== null ? Math.trunc(num(a[3], c)) : base.r2 - base.r1 + 1;
  const w = a.length > 4 && a[4] !== null ? Math.trunc(num(a[4], c)) : base.c2 - base.c1 + 1;
  if (h === 0 || w === 0) throw Errors.REF;
  let r1 = base.r1 + dr;
  let c1 = base.c1 + dc;
  let r2 = r1 + h - (h > 0 ? 1 : -1);
  let c2 = c1 + w - (w > 0 ? 1 : -1);
  if (r2 < r1) [r1, r2] = [r2, r1];
  if (c2 < c1) [c1, c2] = [c2, c1];
  if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) throw Errors.REF;
  const ref: RefValue = { kind: "ref", sheet: base.sheet, r1, c1, r2, c2 };
  c.onDynamicRef?.(ref);
  return ref;
}

export const lookupFunctions: FnDef[] = [
  fn("VLOOKUP", "DÜŞEYARA", "lookup", "aranan_değer; tablo_dizisi; sütun_indis_sayısı; [aralık_bak]", "Tablonun ilk sütununda değeri arar ve aynı satırdaki belirtilen sütundan değer verir.", 3, 4, (a, c) => vlookup(a, c, false), { lift: [0] }),
  fn("HLOOKUP", "YATAYARA", "lookup", "aranan_değer; tablo_dizisi; satır_indis_sayısı; [aralık_bak]", "Tablonun ilk satırında değeri arar ve aynı sütundaki belirtilen satırdan değer verir.", 3, 4, (a, c) => vlookup(a, c, true), { lift: [0] }),
  fn("XLOOKUP", "ÇAPRAZARA", "lookup", "aranan_değer; arama_dizisi; döndürülen_dizi; [bulunamıyorsa]; [eşleştirme_modu]; [arama_modu]", "Bir aralıkta ya da dizide eşleşme arar ve ikinci aralıktan karşılık gelen öğeyi verir.", 3, 6, (a, c) => {
    const x = a[0];
    if (isMulti(x)) {
      const m = matrix(x, c);
      return makeArray(m.map((r) => r.map((v) => scalar(xlookupOne(v, a, c), c))));
    }
    const xs = scalar(x, c);
    if (isError(xs)) return xs;
    return xlookupOne(xs, a, c);
  }, { xlfn: true }),
  fn("MATCH", "KAÇINCI", "lookup", "aranan_değer; aranan_dizi; [eşleştir_tür]", "Bir öğenin dizideki göreli konumunu verir.", 2, 3, (a, c) => {
    const x = a[0] as Scalar;
    if (isError(x)) return x;
    return matchImpl(x, a[1], Math.trunc(optNum(a[2], 1, c)), c);
  }, { lift: [0] }),
  fn("XMATCH", "ÇAPRAZEŞLEŞTİR", "lookup", "aranan_değer; arama_dizisi; [eşleştirme_modu]; [arama_modu]", "Bir öğenin dizideki göreli konumunu verir.", 2, 4, (a, c) => {
    const x = a[0] as Scalar;
    if (isError(x)) return x;
    const { vec, horizontal } = vectorOf(a[1], c, false);
    const i = xsearch(x, vec, Math.trunc(optNum(a[2], 0, c)), Math.trunc(optNum(a[3], 1, c)), refId(a[1], horizontal), c);
    if (i < 0) return Errors.NA;
    return i + 1;
  }, { lift: [0], xlfn: true }),
  fn("INDEX", "İNDİS", "lookup", "dizi; satır_sayısı; [sütun_sayısı]; [alan_sayısı]", "Tablodaki ya da aralıktaki bir değeri ya da başvuruyu verir.", 1, 4, indexImpl),
  lazyFn("CHOOSE", "ELEMAN", "lookup", "indis_sayısı; değer1; [değer2]; ...", "Değer listesinden indis numarasına göre bir değer seçer.", 2, 255, (t, c) => {
    const iv = t[0]();
    if (isMulti(iv)) {
      const m = matrix(iv, c);
      const vals = t.slice(1).map((th) => th());
      return makeArray(m.map((r) => r.map((k) => {
        const n = typeof k === "number" ? Math.trunc(k) : NaN;
        if (!(n >= 1 && n <= vals.length)) return Errors.VALUE;
        return scalar(vals[n - 1], c);
      })));
    }
    const n = Math.trunc(num(iv, c));
    if (n < 1 || n >= t.length) return Errors.VALUE;
    return t[n]();
  }),
  fn("LOOKUP", "ARA", "lookup", "aranan_değer; aranan_vektör; [sonuç_vektör]", "Tek satırlı/sütunlu aralıkta ya da dizide değer arar.", 2, 3, (a, c) => {
    const x = a[0] as Scalar;
    if (isError(x)) return x;
    let look = a[1];
    let result = a[2];
    if (result === undefined) {
      const d = dims(look);
      if (d.rows > 1 && d.cols > 1) {
        const m = matrix(look, c);
        if (d.cols > d.rows) {
          look = makeArray([m[0]]);
          result = makeArray([m[m.length - 1]]);
        } else {
          look = makeArray(m.map((r) => [r[0]]));
          result = makeArray(m.map((r) => [r[r.length - 1]]));
        }
      } else result = look;
    }
    const { vec } = vectorOf(look, c);
    const i = binarySearch(vec, x);
    if (i < 0) return Errors.NA;
    const rv = vectorOf(result, c, false).vec;
    return rv.get(i);
  }, { lift: [0] }),
  fn("OFFSET", "KAYDIR", "lookup", "başv; satırlar; sütunlar; [yükseklik]; [genişlik]", "Bir başvurudan belirtilen satır ve sütun uzaklığındaki başvuruyu verir.", 3, 5, offset, { volatile: true }),
  fn("INDIRECT", "DOLAYLI", "lookup", "başv_metni; [a1]", "Metinle belirtilen başvuruyu verir.", 1, 2, indirect, { volatile: true }),
  fn("ROW", "SATIR", "lookup", "[başv]", "Başvurunun satır numarasını verir.", 0, 1, (a, c) => {
    if (a.length === 0 || a[0] === null) return c.row + 1;
    const r = a[0];
    if (!isRef(r)) throw Errors.VALUE;
    if (r.r1 === r.r2) return r.r1 + 1;
    const n = Math.min(r.r2, r.r1 + 100000) - r.r1 + 1;
    return makeArray(Array.from({ length: n }, (_, i) => [r.r1 + i + 1]));
  }),
  fn("ROWS", "SATIRSAY", "lookup", "dizi", "Başvurudaki ya da dizideki satır sayısını verir.", 1, 1, (a) => {
    if (isError(a[0])) return a[0];
    return dims(a[0]).rows;
  }),
  fn("COLUMN", "SÜTUN", "lookup", "[başv]", "Başvurunun sütun numarasını verir.", 0, 1, (a, c) => {
    if (a.length === 0 || a[0] === null) return c.col + 1;
    const r = a[0];
    if (!isRef(r)) throw Errors.VALUE;
    if (r.c1 === r.c2) return r.c1 + 1;
    return makeArray([Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => r.c1 + i + 1)]);
  }),
  fn("COLUMNS", "SÜTUNSAY", "lookup", "dizi", "Başvurudaki ya da dizideki sütun sayısını verir.", 1, 1, (a) => {
    if (isError(a[0])) return a[0];
    return dims(a[0]).cols;
  }),
  fn("ADDRESS", "ADRES", "lookup", "satır_num; sütun_num; [mutlak_num]; [a1]; [sayfa_metni]", "Hücre başvurusunu metin olarak verir.", 2, 5, (a, c) => {
    const r = Math.trunc(num(a[0], c));
    const col = Math.trunc(num(a[1], c));
    const abs = Math.trunc(optNum(a[2], 1, c));
    const a1 = optBool(a[3], true, c);
    if (r < 1 || col < 1 || r > MAX_ROWS || col > MAX_COLS || abs < 1 || abs > 4) throw Errors.VALUE;
    let s: string;
    if (a1) {
      const ca = abs === 1 || abs === 3 ? "$" : "";
      const ra = abs === 1 || abs === 2 ? "$" : "";
      s = `${ca}${colToName(col - 1)}${ra}${r}`;
    } else {
      const rs = abs === 1 || abs === 2 ? `R${r}` : `R[${r}]`;
      const cs = abs === 1 || abs === 3 ? `C${col}` : `C[${col}]`;
      s = rs + cs;
    }
    const sheet = optText(a[4], "", c);
    return sheet ? `${quoteSheetName(sheet)}!${s}` : s;
  }, { lift: [0, 1, 2, 3, 4] }),
  fn("TRANSPOSE", "DEVRİK_DÖNÜŞÜM", "lookup", "dizi", "Dizinin satır ve sütunlarının yerini değiştirir.", 1, 1, (a, c) => {
    const m = matrix(a[0], c);
    if (m.length === 0) return Errors.VALUE;
    return makeArray(m[0].map((_, j) => m.map((r) => (r[j] === null ? 0 : r[j]))));
  }),
  fn("HYPERLINK", "KÖPRÜ", "lookup", "bağlantı_konumu; [dost_ad]", "Bir belgeyi ya da web sayfasını açan kısayol oluşturur.", 1, 2, (a, c) => (a.length > 1 && a[1] !== null ? scalar(a[1], c) : text(a[0], c)), { lift: true }),
  fn("AREAS", "ALANSAY", "lookup", "başv", "Başvurudaki alan sayısını verir.", 1, 1, (a) => (isMultiRef(a[0]) ? a[0].areas.length : isRef(a[0]) ? 1 : Errors.VALUE)),
];

void isArray;
void addr;
