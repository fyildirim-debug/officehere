/**
 * Veritabanı işlevleri (VSEÇTOPLA, VSEÇSAY…): ilk satırı alan adları olan bir liste,
 * alan (başlık metni ya da 1 tabanlı sütun numarası) ve ölçüt aralığı alır.
 * Ölçüt aralığının ilk satırı alan adlarıdır; aynı satırdaki koşullar VE, farklı
 * satırlar VEYA ile birleşir. İşlecsiz metin "ile başlar" demektir (Gelişmiş
 * Filtre ile aynı), "=metin" tam eşleşmedir; boş ölçüt satırı her kaydı seçer.
 */
import { hasWildcards, makeCriteria, parseNumberText, textEquals, wildcardToRegExp, type Criteria } from "../coerce";
import type { FnContext } from "../context";
import { Errors, isError, type Scalar, type Value } from "../values";
import { average, product, stdev, sum, variance } from "./aggregate";
import { matrix } from "./helpers";
import { fn, type FnDef } from "./registry";

const up = (s: string) => s.trim().toLocaleUpperCase("tr-TR");

function headerText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "number") return String(v);
  if (typeof v === "boolean") return v ? "DOĞRU" : "YANLIŞ";
  if (isError(v)) return v.code;
  return v;
}

/** Tek ölçüt hücresinden eşleme işlevi; boş hücre koşul değildir (null). */
function criterionOf(v: Scalar, ctx: FnContext): Criteria | null {
  if (v === null || v === "") return null;
  if (typeof v !== "string") return makeCriteria(v, ctx.locale);
  if (/^(<=|>=|<>|=|<|>)/.test(v)) return makeCriteria(v, ctx.locale);
  // İşlecsiz metin: sayıysa eşitlik, değilse "ile başlar" (joker karakterlerle)
  if (parseNumberText(v, ctx.locale) !== null) return makeCriteria(v, ctx.locale);
  if (hasWildcards(v)) {
    const re = wildcardToRegExp(v + "*", true);
    return (x) => typeof x === "string" && re.test(x);
  }
  const prefix = up(v);
  return (x) => typeof x === "string" && (x.toLocaleUpperCase("tr-TR").startsWith(prefix) || textEquals(x, v));
}

/** Ölçüte uyan kayıtların (veri satırlarının) seçilen alandaki değerleri. */
function selectField(args: Value[], ctx: FnContext, fieldRequired: boolean): { values: Scalar[]; records: number; hasField: boolean } {
  const db = matrix(args[0], ctx);
  if (db.length < 1 || !db[0]?.length) throw Errors.VALUE;
  const headers = db[0].map((h) => up(headerText(h)));
  let field = -1;
  const f = args[1];
  if (f === null || f === undefined) {
    if (fieldRequired) throw Errors.VALUE;
  } else {
    const fv = matrix(f, ctx)[0]?.[0] ?? null;
    if (isError(fv)) throw fv;
    if (typeof fv === "number") {
      field = Math.trunc(fv) - 1;
      if (field < 0 || field >= headers.length) throw Errors.VALUE;
    } else {
      field = headers.indexOf(up(headerText(fv)));
      if (field < 0) throw Errors.VALUE;
    }
  }
  const crit = matrix(args[2], ctx);
  if (crit.length < 1) throw Errors.VALUE;
  // Ölçüt satırları: [sütun, eşleme] listeleri (eşleşmeyen başlıklı ölçüt sütunları yok sayılır)
  const critCols = crit[0].map((h) => headers.indexOf(up(headerText(h))));
  const rows: [number, Criteria][][] = [];
  for (let i = 1; i < crit.length; i++) {
    const conds: [number, Criteria][] = [];
    crit[i].forEach((cv, j) => {
      const col = critCols[j];
      if (col < 0) return;
      const m = criterionOf(cv, ctx);
      if (m) conds.push([col, m]);
    });
    rows.push(conds);
  }
  const values: Scalar[] = [];
  let records = 0;
  for (let r = 1; r < db.length; r++) {
    const rec = db[r];
    // Ölçüt aralığında yalnız başlık varsa hiçbir kayıt seçilmez
    const hit = rows.some((conds) => conds.every(([col, m]) => m(rec[col] ?? null)));
    if (!hit) continue;
    records++;
    if (field >= 0) values.push(rec[field] ?? null);
  }
  return { values, records, hasField: field >= 0 };
}

function numbersOf(xs: Scalar[]): number[] {
  return xs.filter((x): x is number => typeof x === "number");
}

function dbFn(name: string, tr: string, desc: string, calc: (nums: number[], all: Scalar[]) => Scalar, fieldOptional = false): FnDef {
  return fn(name, tr, "database", fieldOptional ? "veritabanı; [alan]; ölçüt" : "veritabanı; alan; ölçüt", desc, 3, 3, (a, c) => {
    const { values, records, hasField } = selectField(a, c, !fieldOptional);
    // Alan boş bırakılırsa (VSEÇSAY(A1:C9;;E1:E2)) ölçüte uyan kayıtlar sayılır
    if (!hasField) return records;
    return calc(numbersOf(values), values);
  });
}

export const databaseFunctions: FnDef[] = [
  dbFn("DSUM", "VSEÇTOPLA", "Ölçüte uyan kayıtların alanındaki sayıları toplar.", (n) => sum(n)),
  dbFn("DAVERAGE", "VSEÇORT", "Ölçüte uyan kayıtların alanındaki sayıların ortalamasını verir.", (n) => (n.length ? average(n) : Errors.DIV0)),
  dbFn("DCOUNT", "VSEÇSAY", "Ölçüte uyan kayıtların alanındaki sayı içeren hücreleri sayar.", (n) => n.length, true),
  dbFn("DCOUNTA", "VSEÇSAYDOLU", "Ölçüte uyan kayıtların alanındaki boş olmayan hücreleri sayar.", (_, all) => all.filter((x) => x !== null).length, true),
  dbFn("DMAX", "VSEÇMAK", "Ölçüte uyan kayıtların alanındaki en büyük sayıyı verir.", (n) => (n.length ? Math.max(...n) : 0)),
  dbFn("DMIN", "VSEÇMİN", "Ölçüte uyan kayıtların alanındaki en küçük sayıyı verir.", (n) => (n.length ? Math.min(...n) : 0)),
  dbFn("DPRODUCT", "VSEÇÇARP", "Ölçüte uyan kayıtların alanındaki sayıları çarpar.", (n) => product(n)),
  dbFn("DSTDEV", "VSEÇSTDSAPMA", "Ölçüte uyan kayıtlardan örneğe dayalı standart sapmayı tahmin eder.", (n) => stdev(n, true)),
  dbFn("DSTDEVP", "VSEÇSTDSAPMAS", "Ölçüte uyan kayıtların tüm popülasyona dayalı standart sapmasını verir.", (n) => stdev(n, false)),
  dbFn("DVAR", "VSEÇVAR", "Ölçüte uyan kayıtlardan örneğe dayalı varyansı tahmin eder.", (n) => variance(n, true)),
  dbFn("DVARP", "VSEÇVARS", "Ölçüte uyan kayıtların tüm popülasyona dayalı varyansını verir.", (n) => variance(n, false)),
  fn("DGET", "VAL", "database", "veritabanı; alan; ölçüt", "Ölçüte uyan tek kaydın alan değerini verir.", 3, 3, (a, c) => {
    const { values } = selectField(a, c, true);
    if (values.length === 0) throw Errors.VALUE;
    if (values.length > 1) throw Errors.NUM;
    return values[0];
  }),
];

