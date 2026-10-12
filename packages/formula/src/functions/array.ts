/** Dinamik dizi fonksiyonları (taşan sonuçlar). */
import { compareScalars, toBool } from "../coerce";
import type { FnContext } from "../context";
import { invokeLambda } from "../evaluator";
import { Errors, isError, isLambda, makeArray, type Scalar, type Value } from "../values";
import { matrix, num, optBool, optNum, scalar } from "./helpers";
import { fn, type FnDef } from "./registry";

function transpose(m: Scalar[][]): Scalar[][] {
  if (m.length === 0) return [];
  return m[0].map((_, j) => m.map((r) => r[j]));
}

function cmpForSort(a: Scalar, b: Scalar, c: FnContext): number {
  // Boşlar her zaman sona
  if (a === null && b === null) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  if (isError(a) && isError(b)) return 0;
  if (isError(a)) return 1;
  if (isError(b)) return -1;
  return compareScalars(a, b, c.locale);
}

function keyOf(row: Scalar[]): string {
  return row.map((v) => (v === null ? "∅" : typeof v === "string" ? "s" + v.toLocaleUpperCase("tr-TR") : typeof v + String(v))).join("\u0001");
}

function asResult(m: Scalar[][]): Value {
  if (m.length === 0 || m[0].length === 0) return Errors.CALC;
  if (m.length === 1 && m[0].length === 1) return m[0][0];
  return makeArray(m);
}

function lambdaArg(v: Value): { params: string[]; body: unknown; scope: Map<string, Value> | undefined } {
  if (!isLambda(v)) throw Errors.VALUE;
  return v;
}

export const arrayFunctions: FnDef[] = [
  fn("SORT", "SIRALA", "lookup", "dizi; [sıralama_indisi]; [sıralama_düzeni]; [sütuna_göre]", "Bir aralığın ya da dizinin içeriğini sıralar.", 1, 4, (a, c) => {
    let m = matrix(a[0], c).map((r) => r.slice());
    const byCol = optBool(a[3], false, c);
    if (byCol) m = transpose(m);
    const idxs = matrix(a[1] ?? 1, c).flat().map((x) => Math.trunc(typeof x === "number" ? x : 1));
    const ords = matrix(a[2] ?? 1, c).flat().map((x) => (typeof x === "number" ? x : 1));
    for (const i of idxs) if (i < 1 || i > (m[0]?.length ?? 0)) throw Errors.VALUE;
    for (const o of ords) if (o !== 1 && o !== -1) throw Errors.VALUE;
    const idxList = idxs.length ? idxs : [1];
    const sorted = m
      .map((r, k) => ({ r, k }))
      .sort((p, q) => {
        for (let t = 0; t < idxList.length; t++) {
          const col = idxList[t] - 1;
          const ord = ords[t] ?? ords[0] ?? 1;
          const r = cmpForSort(p.r[col], q.r[col], c);
          if (r !== 0) return p.r[col] === null || q.r[col] === null ? r : r * ord;
        }
        return p.k - q.k;
      })
      .map((x) => x.r);
    return asResult(byCol ? transpose(sorted) : sorted);
  }, { xlfn: true }),
  fn("SORTBY", "SIRALAÖLÇÜT", "lookup", "dizi; ölçüt_dizisi1; [sıralama_düzeni1]; ...", "Bir aralığı ya da diziyi karşılık gelen aralıktaki değerlere göre sıralar.", 2, 255, (a, c) => {
    const m = matrix(a[0], c);
    const keys: { v: Scalar[]; ord: number }[] = [];
    for (let i = 1; i < a.length; i += 2) {
      const k = matrix(a[i], c);
      const flat = k.length === m.length ? k.map((r) => r[0]) : k[0].length === m.length ? k[0] : null;
      if (!flat) throw Errors.VALUE;
      const ord = a[i + 1] !== undefined ? num(a[i + 1], c) : 1;
      keys.push({ v: flat, ord });
    }
    const order = m.map((_, i) => i).sort((p, q) => {
      for (const k of keys) {
        const r = cmpForSort(k.v[p], k.v[q], c);
        if (r !== 0) return r * k.ord;
      }
      return p - q;
    });
    return asResult(order.map((i) => m[i]));
  }, { xlfn: true }),
  fn("FILTER", "FİLTRE", "lookup", "dizi; ekle; [boşsa]", "Bir aralığı ya da diziyi ölçütlere göre filtreler.", 2, 3, (a, c) => {
    const m = matrix(a[0], c);
    const inc = matrix(a[1], c);
    const rows = m.length;
    const cols = m[0]?.length ?? 0;
    let out: Scalar[][];
    const truthy = (v: Scalar) => {
      if (isError(v)) throw v;
      const b = toBool(v);
      if (isError(b)) throw b;
      return b;
    };
    if (inc.length === rows && inc[0].length === 1) out = m.filter((_, i) => truthy(inc[i][0]));
    else if (inc.length === 1 && inc[0].length === cols) {
      const keep = inc[0].map(truthy);
      out = m.map((r) => r.filter((_, j) => keep[j]));
      if (out[0]?.length === 0) out = [];
    } else throw Errors.VALUE;
    if (out.length === 0) return a.length > 2 ? a[2] : Errors.CALC;
    return asResult(out);
  }, { xlfn: true }),
  fn("UNIQUE", "BENZERSİZ", "lookup", "dizi; [sütuna_göre]; [yalnızca_bir_kez]", "Bir aralıktaki ya da dizideki benzersiz değerleri verir.", 1, 3, (a, c) => {
    let m = matrix(a[0], c);
    const byCol = optBool(a[1], false, c);
    const once = optBool(a[2], false, c);
    if (byCol) m = transpose(m);
    const counts = new Map<string, number>();
    for (const r of m) counts.set(keyOf(r), (counts.get(keyOf(r)) ?? 0) + 1);
    const seen = new Set<string>();
    const out: Scalar[][] = [];
    for (const r of m) {
      const k = keyOf(r);
      if (seen.has(k)) continue;
      seen.add(k);
      if (once && counts.get(k)! > 1) continue;
      out.push(r.map((v) => (v === null ? 0 : v)));
    }
    return asResult(byCol ? transpose(out) : out);
  }, { xlfn: true }),
  fn("SEQUENCE", "SIRALI", "math", "satırlar; [sütunlar]; [başlangıç]; [adım]", "Sıralı sayılardan oluşan bir dizi oluşturur.", 1, 4, (a, c) => {
    const rows = Math.trunc(num(a[0], c));
    const cols = Math.trunc(optNum(a[1], 1, c));
    const start = optNum(a[2], 1, c);
    const step = optNum(a[3], 1, c);
    if (rows < 1 || cols < 1) throw Errors.CALC;
    if (rows * cols > 1048576) throw Errors.NUM;
    const out: Scalar[][] = [];
    for (let i = 0; i < rows; i++) {
      const r: Scalar[] = [];
      for (let j = 0; j < cols; j++) r.push(start + (i * cols + j) * step);
      out.push(r);
    }
    return asResult(out);
  }, { xlfn: true }),
  fn("VSTACK", "DÜŞEYYIĞ", "lookup", "dizi1; [dizi2]; ...", "Dizileri dikey olarak sırayla ekler.", 1, 254, (a, c) => {
    const ms = a.map((x) => matrix(x, c));
    const w = Math.max(...ms.map((m) => m[0]?.length ?? 0));
    return asResult(ms.flatMap((m) => m.map((r) => [...r, ...new Array(w - r.length).fill(Errors.NA)])));
  }, { xlfn: true }),
  fn("HSTACK", "YATAYYIĞ", "lookup", "dizi1; [dizi2]; ...", "Dizileri yatay olarak sırayla ekler.", 1, 254, (a, c) => {
    const ms = a.map((x) => matrix(x, c));
    const h = Math.max(...ms.map((m) => m.length));
    const out: Scalar[][] = Array.from({ length: h }, () => []);
    for (const m of ms) {
      const w = m[0]?.length ?? 0;
      for (let i = 0; i < h; i++) out[i].push(...(m[i] ?? new Array(w).fill(Errors.NA)));
    }
    return asResult(out);
  }, { xlfn: true }),
  fn("TAKE", "AL", "lookup", "dizi; satırlar; [sütunlar]", "Dizinin başından ya da sonundan satır/sütun alır.", 2, 3, (a, c) => {
    let m = matrix(a[0], c);
    if (a[1] !== null && a[1] !== undefined) {
      const r = Math.trunc(num(a[1], c));
      m = r >= 0 ? m.slice(0, r) : m.slice(r);
    }
    if (a.length > 2 && a[2] !== null) {
      const k = Math.trunc(num(a[2], c));
      m = m.map((row) => (k >= 0 ? row.slice(0, k) : row.slice(k)));
    }
    return asResult(m);
  }, { xlfn: true }),
  fn("DROP", "BIRAK", "lookup", "dizi; satırlar; [sütunlar]", "Dizinin başından ya da sonundan satır/sütun çıkarır.", 2, 3, (a, c) => {
    let m = matrix(a[0], c);
    if (a[1] !== null && a[1] !== undefined) {
      const r = Math.trunc(num(a[1], c));
      m = r >= 0 ? m.slice(r) : m.slice(0, r);
    }
    if (a.length > 2 && a[2] !== null) {
      const k = Math.trunc(num(a[2], c));
      m = m.map((row) => (k >= 0 ? row.slice(k) : row.slice(0, k)));
    }
    return asResult(m);
  }, { xlfn: true }),
  fn("CHOOSECOLS", "SÜTUNSEÇ", "lookup", "dizi; sütun_num1; [sütun_num2]; ...", "Diziden belirtilen sütunları verir.", 2, 254, (a, c) => {
    const m = matrix(a[0], c);
    const w = m[0]?.length ?? 0;
    const idx = a.slice(1).flatMap((x) => matrix(x, c).flat()).map((x) => {
      const n = Math.trunc(Number(x));
      const k = n < 0 ? w + n : n - 1;
      if (!(k >= 0 && k < w)) throw Errors.VALUE;
      return k;
    });
    return asResult(m.map((r) => idx.map((k) => r[k])));
  }, { xlfn: true }),
  fn("CHOOSEROWS", "SATIRSEÇ", "lookup", "dizi; satır_num1; [satır_num2]; ...", "Diziden belirtilen satırları verir.", 2, 254, (a, c) => {
    const m = matrix(a[0], c);
    const idx = a.slice(1).flatMap((x) => matrix(x, c).flat()).map((x) => {
      const n = Math.trunc(Number(x));
      const k = n < 0 ? m.length + n : n - 1;
      if (!(k >= 0 && k < m.length)) throw Errors.VALUE;
      return k;
    });
    return asResult(idx.map((k) => m[k]));
  }, { xlfn: true }),
  fn("TOCOL", "SÜTUNA", "lookup", "dizi; [yoksay]; [sütuna_göre_tara]", "Diziyi tek sütun olarak verir.", 1, 3, (a, c) => {
    let m = matrix(a[0], c);
    const ign = Math.trunc(optNum(a[1], 0, c));
    if (optBool(a[2], false, c)) m = transpose(m);
    let flat = m.flat();
    if (ign === 1 || ign === 3) flat = flat.filter((v) => v !== null);
    if (ign === 2 || ign === 3) flat = flat.filter((v) => !isError(v));
    return asResult(flat.map((v) => [v]));
  }, { xlfn: true }),
  fn("TOROW", "SATIRA", "lookup", "dizi; [yoksay]; [sütuna_göre_tara]", "Diziyi tek satır olarak verir.", 1, 3, (a, c) => {
    let m = matrix(a[0], c);
    const ign = Math.trunc(optNum(a[1], 0, c));
    if (optBool(a[2], false, c)) m = transpose(m);
    let flat = m.flat();
    if (ign === 1 || ign === 3) flat = flat.filter((v) => v !== null);
    if (ign === 2 || ign === 3) flat = flat.filter((v) => !isError(v));
    return asResult([flat]);
  }, { xlfn: true }),
  fn("WRAPROWS", "SATIRSAR", "lookup", "vektör; sarma_sayısı; [doldurma_değeri]", "Vektörü satırlara sarar.", 2, 3, (a, c) => {
    const flat = matrix(a[0], c).flat();
    const n = Math.trunc(num(a[1], c));
    if (n < 1) throw Errors.NUM;
    const pad = a.length > 2 ? scalar(a[2], c) : Errors.NA;
    const out: Scalar[][] = [];
    for (let i = 0; i < flat.length; i += n) {
      const r = flat.slice(i, i + n);
      while (r.length < n) r.push(pad);
      out.push(r);
    }
    return asResult(out);
  }, { xlfn: true }),
  fn("WRAPCOLS", "SÜTUNSAR", "lookup", "vektör; sarma_sayısı; [doldurma_değeri]", "Vektörü sütunlara sarar.", 2, 3, (a, c) => {
    const flat = matrix(a[0], c).flat();
    const n = Math.trunc(num(a[1], c));
    if (n < 1) throw Errors.NUM;
    const pad = a.length > 2 ? scalar(a[2], c) : Errors.NA;
    const cols: Scalar[][] = [];
    for (let i = 0; i < flat.length; i += n) {
      const r = flat.slice(i, i + n);
      while (r.length < n) r.push(pad);
      cols.push(r);
    }
    return asResult(transpose(cols));
  }, { xlfn: true }),
  fn("EXPAND", "GENİŞLET", "lookup", "dizi; satırlar; [sütunlar]; [doldurma_değeri]", "Diziyi belirtilen boyuta genişletir.", 2, 4, (a, c) => {
    const m = matrix(a[0], c);
    const rows = a[1] === null ? m.length : Math.trunc(num(a[1], c));
    const cols = a.length > 2 && a[2] !== null ? Math.trunc(num(a[2], c)) : m[0]?.length ?? 0;
    if (rows < m.length || cols < (m[0]?.length ?? 0)) throw Errors.VALUE;
    const pad = a.length > 3 ? scalar(a[3], c) : Errors.NA;
    return asResult(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => m[i]?.[j] ?? pad)));
  }, { xlfn: true }),
  fn("MAP", "MAP", "logical", "dizi1; lambda_veya_dizi2; ...", "Dizideki her değere LAMBDA uygular.", 2, 254, (a, c) => {
    const lam = lambdaArg(a[a.length - 1]);
    const ms = a.slice(0, -1).map((x) => matrix(x, c));
    return asResult(ms[0].map((r, i) => r.map((_, j) => scalar(invokeLambda(lam, ms.map((m) => m[i]?.[j] ?? null), c), c))));
  }, { xlfn: true }),
  fn("BYROW", "BYROW", "logical", "dizi; lambda", "Her satıra LAMBDA uygular.", 2, 2, (a, c) => {
    const lam = lambdaArg(a[1]);
    const m = matrix(a[0], c);
    return asResult(m.map((r) => [scalar(invokeLambda(lam, [makeArray([r])], c), c)]));
  }, { xlfn: true }),
  fn("BYCOL", "BYCOL", "logical", "dizi; lambda", "Her sütuna LAMBDA uygular.", 2, 2, (a, c) => {
    const lam = lambdaArg(a[1]);
    const m = matrix(a[0], c);
    return asResult([transpose(m).map((col) => scalar(invokeLambda(lam, [makeArray(col.map((v) => [v]))], c), c))]);
  }, { xlfn: true }),
  fn("REDUCE", "REDUCE", "logical", "başlangıç_değeri; dizi; lambda", "Diziyi bir birikimli değere indirger.", 3, 3, (a, c) => {
    const lam = lambdaArg(a[2]);
    let acc: Value = a[0];
    for (const v of matrix(a[1], c).flat()) acc = invokeLambda(lam, [acc, v], c);
    return acc;
  }, { xlfn: true }),
  fn("SCAN", "SCAN", "logical", "başlangıç_değeri; dizi; lambda", "Diziyi tarar ve ara değerleri verir.", 3, 3, (a, c) => {
    const lam = lambdaArg(a[2]);
    let acc: Value = a[0];
    const m = matrix(a[1], c);
    return asResult(m.map((r) => r.map((v) => {
      acc = invokeLambda(lam, [acc, v], c);
      return scalar(acc, c);
    })));
  }, { xlfn: true }),
  fn("MAKEARRAY", "DİZİOLUŞTUR", "logical", "satırlar; sütunlar; lambda", "LAMBDA ile hesaplanan dizi oluşturur.", 3, 3, (a, c) => {
    const rows = Math.trunc(num(a[0], c));
    const cols = Math.trunc(num(a[1], c));
    const lam = lambdaArg(a[2]);
    if (rows < 1 || cols < 1) throw Errors.VALUE;
    return asResult(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => scalar(invokeLambda(lam, [i + 1, j + 1], c), c))));
  }, { xlfn: true }),
];
