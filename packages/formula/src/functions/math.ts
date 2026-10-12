/** Matematik ve trigonometri fonksiyonları. */
import { toNumber } from "../coerce";
import type { FnContext } from "../context";
import { Errors, isError, makeArray, type Scalar, type Value } from "../values";
import * as agg from "./aggregate";
import { check, collectNumbers, finite, int, matchCriteriaRanges, num, optNum, text, valueAt } from "./helpers";
import { fn, type FnDef } from "./registry";
import { roundHalfAway, roundDir, quot, ceilingExcel, floorExcel, fixZero, factorial, gcd2, combin, toRoman, fromRoman, subtotalValues, sumIfImpl, sumProduct, numMatrix, determinant, inverse, pairSum } from "./math-helpers";

export { roundHalfAway, ceilingExcel, floorExcel } from "./math-helpers";

const n1 = (f: (x: number) => number) => (a: Value[], c: FnContext) => finite(f(num(a[0], c)));

export const mathFunctions: FnDef[] = [
  fn("SUM", "TOPLA", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerdeki tüm sayıları toplar.", 1, 255, (a, c) => agg.sum(collectNumbers(a, c))),
  fn("SUMIF", "ETOPLA", "math", "aralık; ölçüt; [toplam_aralığı]", "Verilen koşula uyan hücreleri toplar.", 2, 3, (a, c) => sumIfImpl(a, c)),
  fn(
    "SUMIFS",
    "ÇOKETOPLA",
    "math",
    "toplam_aralığı; ölçüt_aralığı1; ölçüt1; ...",
    "Birden çok koşula uyan hücreleri toplar.",
    3,
    255,
    (a, c) => {
      if ((a.length - 1) % 2) throw Errors.VALUE;
      const pairs: [Value, Value][] = [];
      for (let i = 1; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
      const { mask, rows, cols } = matchCriteriaRanges(pairs, c);
      let total = 0;
      for (let i = 0; i < rows; i++)
        for (let j = 0; j < cols; j++) {
          if (!mask[i][j]) continue;
          const v = valueAt(a[0], i, j, c);
          if (typeof v === "number") total += v;
          else if (isError(v)) throw v;
        }
      return total;
    },
    { xlfn: false },
  ),
  fn("SUMPRODUCT", "TOPLA.ÇARPIM", "math", "dizi1; [dizi2]; ...", "Karşılık gelen dizi öğelerinin çarpımlarının toplamını verir.", 1, 255, sumProduct),
  fn("SUMSQ", "TOPKARE", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerin karelerinin toplamını verir.", 1, 255, (a, c) => agg.sum(collectNumbers(a, c).map((x) => x * x))),
  fn("SUMX2MY2", "TOPX2EY2", "math", "dizi_x; dizi_y", "İki dizideki karşılık gelen değerlerin kareleri farkının toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => x * x - y * y)),
  fn("SUMX2PY2", "TOPX2AY2", "math", "dizi_x; dizi_y", "İki dizideki karşılık gelen değerlerin kareleri toplamının toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => x * x + y * y)),
  fn("SUMXMY2", "TOPXEY2", "math", "dizi_x; dizi_y", "İki dizideki değerlerin farklarının karelerinin toplamı.", 2, 2, (a, c) => pairSum(a, c, (x, y) => (x - y) ** 2)),
  fn("PRODUCT", "ÇARPIM", "math", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerin çarpımını verir.", 1, 255, (a, c) => agg.product(collectNumbers(a, c))),
  fn("ROUND", "YUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı belirtilen basamak sayısına yuvarlar.", 2, 2, (a, c) => roundHalfAway(num(a[0], c), num(a[1], c)), { lift: true }),
  fn("ROUNDUP", "YUKARIYUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı sıfırdan uzağa yuvarlar.", 2, 2, (a, c) => roundDir(num(a[0], c), num(a[1], c), true), { lift: true }),
  fn("ROUNDDOWN", "AŞAĞIYUVARLA", "math", "sayı; sayı_rakamlar", "Sayıyı sıfıra doğru yuvarlar.", 2, 2, (a, c) => roundDir(num(a[0], c), num(a[1], c), false), { lift: true }),
  fn("INT", "TAMSAYI", "math", "sayı", "Sayıyı en yakın alt tamsayıya yuvarlar.", 1, 1, n1(Math.floor), { lift: true }),
  fn("TRUNC", "NSAT", "math", "sayı; [sayı_rakamlar]", "Sayının ondalık kısmını atar.", 1, 2, (a, c) => roundDir(num(a[0], c), optNum(a[1], 0, c), false), { lift: true }),
  fn(
    "MOD",
    "MOD",
    "math",
    "sayı; bölen",
    "Bölme işleminden kalanı verir.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const d = num(a[1], c);
      if (d === 0) throw Errors.DIV0;
      return fixZero(x - d * Math.floor(x / d));
    },
    { lift: true },
  ),
  fn("ABS", "MUTLAK", "math", "sayı", "Sayının mutlak değerini verir.", 1, 1, n1(Math.abs), { lift: true }),
  fn(
    "POWER",
    "KUVVET",
    "math",
    "sayı; üs",
    "Sayının üssünü verir.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const y = num(a[1], c);
      if (x === 0 && y === 0) throw Errors.NUM;
      if (x === 0 && y < 0) throw Errors.DIV0;
      return finite(Math.pow(x, y));
    },
    { lift: true },
  ),
  fn(
    "SQRT",
    "KAREKÖK",
    "math",
    "sayı",
    "Sayının karekökünü verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x < 0) throw Errors.NUM;
      return Math.sqrt(x);
    },
    { lift: true },
  ),
  fn("SQRTPI", "KAREKÖKPİ", "math", "sayı", "(sayı * Pi) değerinin karekökünü verir.", 1, 1, (a, c) => {
    const x = num(a[0], c);
    if (x < 0) throw Errors.NUM;
    return Math.sqrt(x * Math.PI);
  }, { lift: true }),
  fn("EXP", "ÜS", "math", "sayı", "e sayısının verilen kuvvetini verir.", 1, 1, n1(Math.exp), { lift: true }),
  fn(
    "LN",
    "LN",
    "math",
    "sayı",
    "Sayının doğal logaritmasını verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x <= 0) throw Errors.NUM;
      return Math.log(x);
    },
    { lift: true },
  ),
  fn(
    "LOG",
    "LOG",
    "math",
    "sayı; [taban]",
    "Sayının belirtilen tabandaki logaritmasını verir.",
    1,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const b = optNum(a[1], 10, c);
      if (x <= 0 || b <= 0) throw Errors.NUM;
      if (b === 1) throw Errors.DIV0;
      return Math.log(x) / Math.log(b);
    },
    { lift: true },
  ),
  fn(
    "LOG10",
    "LOG10",
    "math",
    "sayı",
    "Sayının 10 tabanında logaritmasını verir.",
    1,
    1,
    (a, c) => {
      const x = num(a[0], c);
      if (x <= 0) throw Errors.NUM;
      return Math.log10(x);
    },
    { lift: true },
  ),
  fn("PI", "Pİ", "math", "", "Pi sayısını (3,14159265358979) verir.", 0, 0, () => Math.PI),
  fn("RAND", "S_SAYI_ÜRET", "math", "", "0 ile 1 arasında rastgele bir sayı verir.", 0, 0, (_a, c) => c.host.random(), { volatile: true }),
  fn(
    "RANDBETWEEN",
    "RASTGELEARADA",
    "math",
    "alt; üst",
    "Belirtilen sayılar arasında rastgele bir tamsayı verir.",
    2,
    2,
    (a, c) => {
      const lo = Math.ceil(num(a[0], c));
      const hi = Math.floor(num(a[1], c));
      if (lo > hi) throw Errors.NUM;
      return lo + Math.floor(c.host.random() * (hi - lo + 1));
    },
    { volatile: true },
  ),
  fn(
    "RANDARRAY",
    "RASTGDİZİ",
    "math",
    "[satırlar]; [sütunlar]; [min]; [maks]; [tamsayı]",
    "Rastgele sayılardan oluşan bir dizi verir.",
    0,
    5,
    (a, c) => {
      const rows = Math.trunc(optNum(a[0], 1, c));
      const cols = Math.trunc(optNum(a[1], 1, c));
      const lo = optNum(a[2], 0, c);
      const hi = optNum(a[3], 1, c);
      const whole = a[4] !== undefined && a[4] !== null ? !!num(a[4], c) : false;
      if (rows < 1 || cols < 1 || lo > hi) throw Errors.VALUE;
      const data: Scalar[][] = [];
      for (let i = 0; i < rows; i++) {
        const r: Scalar[] = [];
        for (let j = 0; j < cols; j++) {
          const x = c.host.random();
          r.push(whole ? Math.ceil(lo) + Math.floor(x * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + x * (hi - lo));
        }
        data.push(r);
      }
      return makeArray(data);
    },
    { volatile: true, xlfn: true },
  ),
  fn("CEILING", "TAVANAYUVARLA", "math", "sayı; anlam", "Sayıyı anlamın en yakın katına yukarı yuvarlar.", 2, 2, (a, c) => fixZero(ceilingExcel(num(a[0], c), num(a[1], c))), { lift: true }),
  fn(
    "CEILING.MATH",
    "TAVANAYUVARLA.MATEMATİK",
    "math",
    "sayı; [anlam]; [mod]",
    "Sayıyı en yakın tamsayıya ya da anlamın en yakın katına yukarı yuvarlar.",
    1,
    3,
    (a, c) => {
      const x = num(a[0], c);
      const s = Math.abs(optNum(a[1], 1, c));
      const mode = optNum(a[2], 0, c);
      if (s === 0) return 0;
      if (x >= 0 || mode === 0) return fixZero(Math.ceil(quot(x, s)) * s);
      return fixZero(-Math.ceil(quot(-x, s)) * s);
    },
    { lift: true, xlfn: true },
  ),
  fn("CEILING.PRECISE", "TAVANAYUVARLA.DUYARLI", "math", "sayı; [anlam]", "Sayıyı anlamın en yakın katına yukarı yuvarlar.", 1, 2, (a, c) => {
    const x = num(a[0], c);
    const s = Math.abs(optNum(a[1], 1, c));
    return s === 0 ? 0 : fixZero(Math.ceil(quot(x, s)) * s);
  }, { lift: true, xlfn: true }),
  fn("FLOOR", "TABANAYUVARLA", "math", "sayı; anlam", "Sayıyı anlamın en yakın katına aşağı yuvarlar.", 2, 2, (a, c) => fixZero(floorExcel(num(a[0], c), num(a[1], c))), { lift: true }),
  fn(
    "FLOOR.MATH",
    "TABANAYUVARLA.MATEMATİK",
    "math",
    "sayı; [anlam]; [mod]",
    "Sayıyı en yakın tamsayıya ya da anlamın en yakın katına aşağı yuvarlar.",
    1,
    3,
    (a, c) => {
      const x = num(a[0], c);
      const s = Math.abs(optNum(a[1], 1, c));
      const mode = optNum(a[2], 0, c);
      if (s === 0) return 0;
      if (x >= 0 || mode === 0) return fixZero(Math.floor(quot(x, s)) * s);
      return fixZero(-Math.floor(quot(-x, s)) * s);
    },
    { lift: true, xlfn: true },
  ),
  fn("FLOOR.PRECISE", "TABANAYUVARLA.DUYARLI", "math", "sayı; [anlam]", "Sayıyı anlamın en yakın katına aşağı yuvarlar.", 1, 2, (a, c) => {
    const x = num(a[0], c);
    const s = Math.abs(optNum(a[1], 1, c));
    return s === 0 ? 0 : fixZero(Math.floor(quot(x, s)) * s);
  }, { lift: true, xlfn: true }),
  fn("SIGN", "İŞARET", "math", "sayı", "Sayının işaretini verir.", 1, 1, n1((x) => (x > 0 ? 1 : x < 0 ? -1 : 0)), { lift: true }),
  fn("FACT", "ÇARPINIM", "math", "sayı", "Sayının faktöriyelini verir.", 1, 1, (a, c) => factorial(num(a[0], c)), { lift: true }),
  fn("FACTDOUBLE", "ÇİFTFAKTÖR", "math", "sayı", "Sayının çift faktöriyelini verir.", 1, 1, (a, c) => {
    const n = Math.floor(num(a[0], c));
    if (n < -1) throw Errors.NUM;
    let r = 1;
    for (let i = n; i > 1; i -= 2) r *= i;
    return r;
  }, { lift: true }),
  fn("GCD", "OBEB", "math", "sayı1; [sayı2]; ...", "En büyük ortak böleni verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    return xs.reduce((g, x) => gcd2(g, x), 0);
  }),
  fn("LCM", "OKEK", "math", "sayı1; [sayı2]; ...", "En küçük ortak katı verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    if (xs.some((x) => x === 0)) return 0;
    return xs.reduce((l, x) => (l * x) / gcd2(l, x), 1);
  }),
  fn("EVEN", "ÇİFT", "math", "sayı", "Sayıyı en yakın çift tamsayıya yukarı yuvarlar.", 1, 1, n1((x) => {
    const r = Math.ceil(Math.abs(x) / 2) * 2;
    return x < 0 ? -r : r;
  }), { lift: true }),
  fn("ODD", "TEK", "math", "sayı", "Sayıyı en yakın tek tamsayıya yukarı yuvarlar.", 1, 1, n1((x) => {
    let r = Math.ceil(Math.abs(x));
    if (r % 2 === 0) r += 1;
    return x < 0 ? -r : r;
  }), { lift: true }),
  fn(
    "MROUND",
    "KYUVARLA",
    "math",
    "sayı; katsayı",
    "Sayıyı istenen katsayıya yuvarlar.",
    2,
    2,
    (a, c) => {
      const x = num(a[0], c);
      const m = num(a[1], c);
      if (m === 0) return 0;
      if (x * m < 0) throw Errors.NUM;
      return fixZero(roundHalfAway(x / m, 0) * m);
    },
    { lift: true },
  ),
  fn(
    "QUOTIENT",
    "BÖLÜM",
    "math",
    "pay; payda",
    "Bölmenin tamsayı kısmını verir.",
    2,
    2,
    (a, c) => {
      const d = num(a[1], c);
      if (d === 0) throw Errors.DIV0;
      return Math.trunc(num(a[0], c) / d);
    },
    { lift: true },
  ),
  fn("SUBTOTAL", "ALTTOPLAM", "math", "işlev_sayısı; başv1; ...", "Bir listedeki ya da veritabanındaki alt toplamı verir.", 2, 255, (a, c) => {
    const code = int(a[0], c);
    const skipHidden = code > 100;
    const k = skipHidden ? code - 100 : code;
    if (k < 1 || k > 11) throw Errors.VALUE;
    const { nums, countA } = subtotalValues(a.slice(1), c, skipHidden);
    return agg.subtotalBy(k, nums, countA);
  }),
  fn("AGGREGATE", "TOPLAMA", "math", "işlev_sayısı; seçenekler; başv1; ...", "Bir listedeki toplamı, gizli satırları ve hataları yok sayma seçenekleriyle verir.", 3, 255, (a, c) => {
    const code = int(a[0], c);
    const opt = int(a[1], c);
    const ignoreErrors = opt === 2 || opt === 3 || opt === 6 || opt === 7;
    const skipHidden = opt === 1 || opt === 3 || opt === 5 || opt === 7;
    if (code >= 14 && code <= 19) {
      const { nums } = subtotalValues([a[2]], c, skipHidden, ignoreErrors);
      const k = num(a[3], c);
      switch (code) {
        case 14:
          return agg.kth(nums, k, true);
        case 15:
          return agg.kth(nums, k, false);
        case 16:
          return agg.percentileInc(nums, k);
        case 17:
          return agg.percentileInc(nums, k / 4);
        case 18:
          return agg.percentileExc(nums, k);
        default:
          return agg.percentileExc(nums, k / 4);
      }
    }
    const { nums, countA } = subtotalValues(a.slice(2), c, skipHidden, ignoreErrors);
    return agg.subtotalBy(code, nums, countA);
  }, { xlfn: true }),
  fn("RADIANS", "RADYAN", "math", "açı", "Dereceyi radyana çevirir.", 1, 1, n1((x) => (x * Math.PI) / 180), { lift: true }),
  fn("DEGREES", "DERECE", "math", "açı", "Radyanı dereceye çevirir.", 1, 1, n1((x) => (x * 180) / Math.PI), { lift: true }),
  fn("SIN", "SİN", "math", "sayı", "Açının sinüsünü verir.", 1, 1, n1(Math.sin), { lift: true }),
  fn("COS", "COS", "math", "sayı", "Açının kosinüsünü verir.", 1, 1, n1(Math.cos), { lift: true }),
  fn("TAN", "TAN", "math", "sayı", "Açının tanjantını verir.", 1, 1, n1(Math.tan), { lift: true }),
  fn("ASIN", "ASİN", "math", "sayı", "Sayının arksinüsünü verir.", 1, 1, n1((x) => {
    if (x < -1 || x > 1) throw Errors.NUM;
    return Math.asin(x);
  }), { lift: true }),
  fn("ACOS", "ACOS", "math", "sayı", "Sayının arkkosinüsünü verir.", 1, 1, n1((x) => {
    if (x < -1 || x > 1) throw Errors.NUM;
    return Math.acos(x);
  }), { lift: true }),
  fn("ATAN", "ATAN", "math", "sayı", "Sayının arktanjantını verir.", 1, 1, n1(Math.atan), { lift: true }),
  fn("ATAN2", "ATAN2", "math", "x_sayısı; y_sayısı", "x ve y koordinatlarının arktanjantını verir.", 2, 2, (a, c) => {
    const x = num(a[0], c);
    const y = num(a[1], c);
    if (x === 0 && y === 0) throw Errors.DIV0;
    return Math.atan2(y, x);
  }, { lift: true }),
  fn("SINH", "SİNH", "math", "sayı", "Hiperbolik sinüsü verir.", 1, 1, n1(Math.sinh), { lift: true }),
  fn("COSH", "COSH", "math", "sayı", "Hiperbolik kosinüsü verir.", 1, 1, n1(Math.cosh), { lift: true }),
  fn("TANH", "TANH", "math", "sayı", "Hiperbolik tanjantı verir.", 1, 1, n1(Math.tanh), { lift: true }),
  fn("ASINH", "ASİNH", "math", "sayı", "Ters hiperbolik sinüsü verir.", 1, 1, n1(Math.asinh), { lift: true }),
  fn("ACOSH", "ACOSH", "math", "sayı", "Ters hiperbolik kosinüsü verir.", 1, 1, n1((x) => {
    if (x < 1) throw Errors.NUM;
    return Math.acosh(x);
  }), { lift: true }),
  fn("ATANH", "ATANH", "math", "sayı", "Ters hiperbolik tanjantı verir.", 1, 1, n1((x) => {
    if (x <= -1 || x >= 1) throw Errors.NUM;
    return Math.atanh(x);
  }), { lift: true }),
  fn("COT", "COT", "math", "sayı", "Açının kotanjantını verir.", 1, 1, n1((x) => {
    if (x === 0) throw Errors.DIV0;
    return 1 / Math.tan(x);
  }), { lift: true, xlfn: true }),
  fn("CSC", "CSC", "math", "sayı", "Açının kosekantını verir.", 1, 1, n1((x) => {
    if (x === 0) throw Errors.DIV0;
    return 1 / Math.sin(x);
  }), { lift: true, xlfn: true }),
  fn("SEC", "SEC", "math", "sayı", "Açının sekantını verir.", 1, 1, n1((x) => 1 / Math.cos(x)), { lift: true, xlfn: true }),
  fn("COMBIN", "KOMBİNASYON", "math", "sayı; sayı_seçilen", "Verilen öğe sayısı için kombinasyon sayısını verir.", 2, 2, (a, c) => combin(num(a[0], c), num(a[1], c)), { lift: true }),
  fn("COMBINA", "KOMBİNASYONA", "math", "sayı; sayı_seçilen", "Yinelemeli kombinasyon sayısını verir.", 2, 2, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const k = Math.floor(num(a[1], c));
    if (n < 0 || k < 0) throw Errors.NUM;
    return n === 0 && k === 0 ? 1 : combin(n + k - 1, k);
  }, { lift: true, xlfn: true }),
  fn("PERMUT", "PERMÜTASYON", "math", "sayı; sayı_seçilen", "Permütasyon sayısını verir.", 2, 2, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const k = Math.floor(num(a[1], c));
    if (n < 0 || k < 0 || k > n) throw Errors.NUM;
    let r = 1;
    for (let i = 0; i < k; i++) r *= n - i;
    return r;
  }, { lift: true }),
  fn("MULTINOMIAL", "ÇOKTERİMLİ", "math", "sayı1; [sayı2]; ...", "Sayılar kümesinin çok terimlisini verir.", 1, 255, (a, c) => {
    const xs = collectNumbers(a, c).map(Math.floor);
    if (xs.some((x) => x < 0)) throw Errors.NUM;
    return Math.round(factorial(agg.sum(xs)) / xs.reduce((p, x) => p * factorial(x), 1));
  }),
  fn("ROMAN", "ROMEN", "math", "sayı; [biçim]", "Arap rakamını Romen rakamına çevirir.", 1, 2, (a, c) => toRoman(num(a[0], c)), { lift: true }),
  fn("ARABIC", "ARAP", "math", "metin", "Romen rakamını Arap rakamına çevirir.", 1, 1, (a, c) => fromRoman(text(a[0], c)), { lift: true, xlfn: true }),
  fn("BASE", "TABAN", "math", "sayı; taban; [min_uzunluk]", "Sayıyı verilen tabanda metne çevirir.", 2, 3, (a, c) => {
    const n = Math.floor(num(a[0], c));
    const b = Math.floor(num(a[1], c));
    if (n < 0 || b < 2 || b > 36) throw Errors.NUM;
    return n.toString(b).toUpperCase().padStart(Math.floor(optNum(a[2], 0, c)), "0");
  }, { lift: true, xlfn: true }),
  fn("DECIMAL", "ONDALIK", "math", "sayı; taban", "Verilen tabandaki metni ondalık sayıya çevirir.", 2, 2, (a, c) => {
    const s = text(a[0], c);
    const b = Math.floor(num(a[1], c));
    if (b < 2 || b > 36) throw Errors.NUM;
    const n = parseInt(s, b);
    if (Number.isNaN(n)) throw Errors.NUM;
    return n;
  }, { lift: true, xlfn: true }),
  fn("MMULT", "DÇARP", "math", "dizi1; dizi2", "İki dizinin matris çarpımını verir.", 2, 2, (a, c) => {
    const A = numMatrix(a[0], c);
    const B = numMatrix(a[1], c);
    if (A[0].length !== B.length) throw Errors.VALUE;
    return makeArray(A.map((row) => B[0].map((_, j) => row.reduce((s, x, k) => s + x * B[k][j], 0))));
  }),
  fn("MDETERM", "DETERMİNANT", "math", "dizi", "Dizinin matris determinantını verir.", 1, 1, (a, c) => {
    const m = numMatrix(a[0], c);
    if (m.length !== m[0].length) throw Errors.VALUE;
    return determinant(m);
  }),
  fn("MINVERSE", "DİZEY_TERS", "math", "dizi", "Dizinin ters matrisini verir.", 1, 1, (a, c) => {
    const m = numMatrix(a[0], c);
    if (m.length !== m[0].length) throw Errors.VALUE;
    return makeArray(inverse(m));
  }),
  fn("MUNIT", "BİRİMMATRİS", "math", "boyut", "Belirtilen boyutta birim matris verir.", 1, 1, (a, c) => {
    const n = int(a[0], c);
    if (n < 1) throw Errors.VALUE;
    return makeArray(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
  }, { xlfn: true }),
];


void check;
void toNumber;
