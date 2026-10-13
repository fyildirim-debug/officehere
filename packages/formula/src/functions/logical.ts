/** Mantıksal ve bilgi fonksiyonları. */
import { compareScalars, toBool } from "../coerce";
import type { FnContext } from "../context";
import { OMITTED_KEY, broadcast, broadcastMany, invokeLambda } from "../evaluator";
import { addr } from "../references";
import { Errors, FormulaError, isArray, isError, isLambda, isRef, makeArray, type Scalar, type Value } from "../values";
import { check, forEachValue, isMulti, matrix, num, optText, scalar } from "./helpers";
import { fn, lazyFn, type FnDef, type Thunk } from "./registry";

function toOperand(v: Value, c: FnContext): Scalar | Scalar[][] {
  if (isMulti(v)) return matrix(v, c);
  if (isLambda(v)) return Errors.VALUE;
  return scalar(v, c);
}

/** Mantıksal değerleri toplar (VE/YADA/ÖZELVEYA). */
function logicalValues(args: Value[], c: FnContext): boolean[] {
  const out: boolean[] = [];
  for (const a of args) {
    forEachValue(a, c, (s, fromRef) => {
      if (isError(s)) throw s;
      if (s === null) {
        if (!fromRef) out.push(false);
        return;
      }
      if (typeof s === "boolean") out.push(s);
      else if (typeof s === "number") out.push(s !== 0);
      else if (!fromRef) {
        const b = toBool(s);
        if (isError(b)) throw b;
        out.push(b);
      }
    });
  }
  if (out.length === 0) throw Errors.VALUE;
  return out;
}

function ifImpl(t: Thunk[], c: FnContext): Value {
  const cond = t[0]();
  if (isMulti(cond)) {
    const cm = matrix(cond, c);
    const yes = t.length > 1 ? toOperand(t[1](), c) : true;
    const no = t.length > 2 ? toOperand(t[2](), c) : false;
    return broadcastMany([cm, yes, no], ([cv, y, n]) => {
      if (isError(cv)) return cv;
      const b = toBool(cv);
      if (isError(b)) return b;
      return b ? (y ?? 0) : (n ?? 0);
    });
  }
  const cs = scalar(cond, c);
  if (isError(cs)) return cs;
  const b = toBool(cs);
  if (isError(b)) return b;
  if (b) return t.length > 1 ? t[1]() : true;
  return t.length > 2 ? t[2]() : false;
}

function errType(v: Scalar): number | null {
  if (!isError(v)) return null;
  const map: Record<string, number> = { "#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7, "#GETTING_DATA": 8, "#SPILL!": 9, "#CALC!": 14 };
  return map[v.code] ?? null;
}

/** EĞERHATA/EĞERYOKSA: dizilerde öğe öğe uygulanır. */
function ifErrorLike(t: Thunk[], c: FnContext, pred: (e: FormulaError) => boolean): Value {
  const v = t[0]();
  if (isMulti(v)) {
    const m = matrix(v, c);
    if (!m.flat().some((s) => isError(s) && pred(s))) return isRef(v) ? makeArray(m) : v;
    const alt = toOperand(t[1](), c);
    return broadcast(m, alt, (s, a) => (isError(s) && pred(s) ? (a ?? 0) : s));
  }
  const s = scalar(v, c);
  if (isError(s) && pred(s)) return t[1]();
  return v;
}

export const logicalFunctions: FnDef[] = [
  lazyFn("IF", "EĞER", "logical", "mantıksal_sınama; [eğer_doğruysa_değer]; [eğer_yanlışsa_değer]", "Koşul sağlanıyorsa bir değer, sağlanmıyorsa başka bir değer verir.", 1, 3, ifImpl),
  lazyFn("IFS", "ÇOKEĞER", "logical", "mantıksal_sınama1; eğer_doğruysa_değer1; ...", "İlk DOĞRU koşula karşılık gelen değeri verir.", 2, 254, (t, c) => {
    if (t.length % 2) return Errors.VALUE;
    for (let i = 0; i < t.length; i += 2) {
      const s = scalar(t[i](), c);
      if (isError(s)) return s;
      const b = toBool(s);
      if (isError(b)) return b;
      if (b) return t[i + 1]();
    }
    return Errors.NA;
  }, { xlfn: true }),
  fn("AND", "VE", "logical", "mantıksal1; [mantıksal2]; ...", "Tüm bağımsız değişkenler DOĞRU ise DOĞRU verir.", 1, 255, (a, c) => logicalValues(a, c).every(Boolean)),
  fn("OR", "YADA", "logical", "mantıksal1; [mantıksal2]; ...", "Bağımsız değişkenlerden biri DOĞRU ise DOĞRU verir.", 1, 255, (a, c) => logicalValues(a, c).some(Boolean)),
  fn("XOR", "ÖZELVEYA", "logical", "mantıksal1; [mantıksal2]; ...", "Tüm bağımsız değişkenlerin mantıksal özel veya değerini verir.", 1, 255, (a, c) => logicalValues(a, c).filter(Boolean).length % 2 === 1, { xlfn: true }),
  fn("NOT", "DEĞİL", "logical", "mantıksal", "Bağımsız değişkenin mantıksal tersini verir.", 1, 1, (a) => !check(toBool(a[0] as Scalar)), { lift: true }),
  fn("TRUE", "DOĞRU", "logical", "", "DOĞRU mantıksal değerini verir.", 0, 0, () => true),
  fn("FALSE", "YANLIŞ", "logical", "", "YANLIŞ mantıksal değerini verir.", 0, 0, () => false),
  lazyFn("IFERROR", "EĞERHATA", "logical", "değer; eğer_hataysa_değer", "İfade hata veriyorsa belirtilen değeri, vermiyorsa ifadenin değerini verir.", 2, 2, (t, c) => ifErrorLike(t, c, () => true)),
  lazyFn("IFNA", "EĞERYOKSA", "logical", "değer; eğer_yoksa_değer", "İfade #YOK veriyorsa belirtilen değeri verir.", 2, 2, (t, c) => ifErrorLike(t, c, (e) => e.code === "#N/A"), { xlfn: true }),
  lazyFn("SWITCH", "İLKEŞLEŞEN", "logical", "ifade; değer1; sonuç1; [varsayılan_veya_değer2]; ...", "İfadeyi değer listesiyle karşılaştırır ve ilk eşleşen sonucu verir.", 3, 254, (t, c) => {
    const x = scalar(t[0](), c);
    if (isError(x)) return x;
    let i = 1;
    for (; i + 1 < t.length; i += 2) {
      const v = scalar(t[i](), c);
      if (isError(v)) return v;
      if (compareScalars(x, v, c.locale) === 0 && typeof x === typeof v) return t[i + 1]();
    }
    return i < t.length ? t[i]() : Errors.NA;
  }, { xlfn: true }),
  lazyFn("LET", "LET", "logical", "ad1; ad_değeri1; hesaplama_veya_ad2; ...", "Hesaplama sonuçlarını adlara atar.", 3, 253, (t, c, nodes) => {
    if (t.length % 2 === 0) return Errors.VALUE;
    const scope = new Map(c.scope ?? []);
    for (let i = 0; i + 1 < nodes.length; i += 2) {
      const n = nodes[i];
      if (n.type !== "name") return Errors.VALUE;
      const v = c.evalNode(nodes[i + 1], scope);
      scope.set(n.name.toLocaleUpperCase("tr-TR"), v);
      scope.set(n.name.toUpperCase(), v);
    }
    return c.evalNode(nodes[nodes.length - 1], scope);
  }, { xlfn: true }),
  lazyFn("LAMBDA", "LAMBDA", "logical", "[parametre1]; ...; hesaplama", "Özel, yeniden kullanılabilir bir fonksiyon oluşturur.", 1, 254, (_t, c, nodes) => {
    const params: string[] = [];
    for (const n of nodes.slice(0, -1)) {
      if (n.type !== "name") return Errors.VALUE;
      params.push(n.name.toLocaleUpperCase("tr-TR"));
    }
    return { kind: "lambda", params, body: nodes[nodes.length - 1], scope: c.scope };
  }, { xlfn: true }),
  lazyFn("ISOMITTED", "EKSİKMİ", "logical", "bağımsız_değişken", "LAMBDA parametresinin çağrıda boş bırakılıp bırakılmadığını sınar.", 1, 1, (_t, c, nodes) => {
    const n = nodes[0];
    if (n.type !== "name") return false;
    return c.scope?.get(OMITTED_KEY + n.name.toLocaleUpperCase("tr-TR")) === true;
  }, { xlfn: true }),
];

/* ------------------------------------------------------------------ */

function isAnyRefArg(v: Value): boolean {
  return isRef(v);
}

export const infoFunctions: FnDef[] = [
  fn("ISBLANK", "EBOŞSA", "info", "değer", "Değer boşsa DOĞRU verir.", 1, 1, (a) => a[0] === null && true, { lift: true }),
  fn("ISNUMBER", "ESAYIYSA", "info", "değer", "Değer sayıysa DOĞRU verir.", 1, 1, (a) => typeof a[0] === "number", { lift: true }),
  fn("ISTEXT", "EMETİNSE", "info", "değer", "Değer metinse DOĞRU verir.", 1, 1, (a) => typeof a[0] === "string", { lift: true }),
  fn("ISNONTEXT", "EMETİNDEĞİLSE", "info", "değer", "Değer metin değilse DOĞRU verir.", 1, 1, (a) => typeof a[0] !== "string", { lift: true }),
  fn("ISERROR", "EHATALIYSA", "info", "değer", "Değer herhangi bir hata değeriyse DOĞRU verir.", 1, 1, (a) => isError(a[0]), { lift: true }),
  fn("ISERR", "EHATAYSA", "info", "değer", "Değer #YOK dışında bir hata değeriyse DOĞRU verir.", 1, 1, (a) => isError(a[0]) && a[0].code !== "#N/A", { lift: true }),
  fn("ISNA", "EYOKSA", "info", "değer", "Değer #YOK hata değeriyse DOĞRU verir.", 1, 1, (a) => isError(a[0]) && a[0].code === "#N/A", { lift: true }),
  fn("ISLOGICAL", "EMANTIKSALSA", "info", "değer", "Değer mantıksal değerse DOĞRU verir.", 1, 1, (a) => typeof a[0] === "boolean", { lift: true }),
  fn("ISEVEN", "ÇİFTMİ", "info", "sayı", "Sayı çiftse DOĞRU verir.", 1, 1, (a, c) => Math.trunc(num(a[0], c)) % 2 === 0, { lift: true }),
  fn("ISODD", "TEKMİ", "info", "sayı", "Sayı tekse DOĞRU verir.", 1, 1, (a, c) => Math.abs(Math.trunc(num(a[0], c))) % 2 === 1, { lift: true }),
  fn("ISREF", "EREFSE", "info", "değer", "Değer bir başvuruysa DOĞRU verir.", 1, 1, (a) => isAnyRefArg(a[0]) || (typeof a[0] === "object" && a[0] !== null && (a[0] as { kind?: string }).kind === "refs")),
  fn("ISFORMULA", "EFORMÜLSE", "info", "başvuru", "Başvurulan hücre formül içeriyorsa DOĞRU verir.", 1, 1, (a, c) => {
    const r = a[0];
    if (!isRef(r)) throw Errors.VALUE;
    return c.host.getFormulaText(r.sheet, r.r1, r.c1) !== undefined;
  }, { xlfn: true }),
  fn("FORMULATEXT", "FORMÜLMETNİ", "lookup", "başvuru", "Formülü metin olarak verir.", 1, 1, (a, c) => {
    const r = a[0];
    if (!isRef(r)) throw Errors.VALUE;
    const f = c.host.getFormulaText(r.sheet, r.r1, r.c1);
    if (f === undefined) throw Errors.NA;
    return f;
  }, { xlfn: true }),
  fn("NA", "YOKSAY", "info", "", "#YOK hata değerini verir.", 0, 0, () => Errors.NA),
  fn("TYPE", "TÜR", "info", "değer", "Değerin veri türünü gösteren sayıyı verir.", 1, 1, (a, c) => {
    const v = a[0];
    // Dizi sabiti tek öğeli olsa da dizidir ({1} → 64); çok hücreli başvuru da öyle
    if (isMulti(v) || isArray(v)) return 64;
    const s = scalar(v, c);
    if (typeof s === "number" || s === null) return 1;
    if (typeof s === "string") return 2;
    if (typeof s === "boolean") return 4;
    return 16;
  }),
  fn("ERROR.TYPE", "HATA.TİPİ", "info", "hata_değer", "Hata değerine karşılık gelen sayıyı verir.", 1, 1, (a) => errType(a[0] as Scalar) ?? Errors.NA, { lift: true }),
  fn("N", "S", "info", "değer", "Değeri sayıya çevirir.", 1, 1, (a) => {
    const v = a[0] as Scalar;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (isError(v)) return v;
    return 0;
  }, { lift: true }),
  fn("SHEET", "SAYFA", "info", "[değer]", "Başvurulan sayfanın numarasını verir.", 0, 1, (a, c) => {
    if (a.length === 0 || a[0] === null) return c.host.sheetIndex(c.sheet) + 1;
    const v = a[0];
    if (isRef(v)) return c.host.sheetIndex(v.sheet) + 1;
    if (typeof v === "string") {
      const id = c.host.sheetIdByName(v);
      if (id === undefined) throw Errors.NA;
      return c.host.sheetIndex(id) + 1;
    }
    throw Errors.NA;
  }, { xlfn: true }),
  fn("SHEETS", "SAYFALAR", "info", "[başvuru]", "Başvurudaki sayfa sayısını verir.", 0, 1, (a, c) => (a.length && a[0] !== null ? 1 : c.host.sheetCount()), { xlfn: true }),
  fn("CELL", "HÜCRE", "info", "bilgi_türü; [başvuru]", "Hücrenin biçimi, konumu ya da içeriği hakkında bilgi verir.", 1, 2, (a, c) => {
    const t = optText(a[0], "", c).toLocaleLowerCase("tr-TR");
    const r = a[1];
    const ref = isRef(r) ? r : { kind: "ref" as const, sheet: c.sheet, r1: c.row, c1: c.col, r2: c.row, c2: c.col };
    const v = c.host.getCellValue(ref.sheet, ref.r1, ref.c1);
    switch (t) {
      case "address":
      case "adres":
        return addr(ref.r1, ref.c1).replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");
      case "row":
      case "satır":
        return ref.r1 + 1;
      case "col":
      case "sütun":
        return ref.c1 + 1;
      case "contents":
      case "içerik":
        return v ?? 0;
      case "type":
      case "tür":
        return v === null ? "b" : typeof v === "string" ? "l" : "v";
      case "filename":
      case "dosyaadı":
        return c.host.sheetName(ref.sheet) ?? "";
      default:
        throw Errors.VALUE;
    }
  }, { volatile: true }),
];

export function callLambdaValue(lam: Value, args: Value[], c: FnContext): Value {
  if (!isLambda(lam)) throw Errors.VALUE;
  return invokeLambda(lam, args, c);
}
