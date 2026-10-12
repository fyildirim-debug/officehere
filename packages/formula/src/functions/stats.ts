/** İstatistik fonksiyonları. */
import { parseNumberText } from "../coerce";
import type { FnContext } from "../context";
import { Errors, isError, isRef, makeArray, type Scalar, type Value } from "../values";
import * as agg from "./aggregate";
import { clipRef, collectNumbers, flatValues, forEachValue, matchCriteriaRanges, matrix, num, optNum, scalar, valueAt } from "./helpers";
import { fn, type FnDef } from "./registry";
import { lnGammaHP } from "./special";
import { makeCriteria } from "../coerce";

/** ORTALAMAA/MAKA tarzı: başvurudaki metin 0, mantıksal 1/0. */
const numsA = (a: Value[], c: FnContext) => collectNumbers(a, c, { refTextAsZero: true });
const nums = (a: Value[], c: FnContext) => collectNumbers(a, c);

/** İki dizi için eşleşen sayı çiftleri (her ikisi de sayı olan konumlar). */
function pairs(a: Value, b: Value, c: FnContext): [number[], number[]] {
  const X = matrix(a, c).flat();
  const Y = matrix(b, c).flat();
  if (X.length !== Y.length) throw Errors.NA;
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i < X.length; i++) {
    const x = X[i];
    const y = Y[i];
    if (isError(x)) throw x;
    if (isError(y)) throw y;
    if (typeof x === "number" && typeof y === "number") {
      xs.push(x);
      ys.push(y);
    }
  }
  return [xs, ys];
}

function countImpl(a: Value[], c: FnContext): number {
  let n = 0;
  for (const v of a) {
    forEachValue(v, c, (s, fromRef) => {
      if (typeof s === "number") n++;
      else if (!fromRef) {
        if (typeof s === "boolean") n++;
        else if (typeof s === "string" && parseNumberText(s, c.locale) !== null) n++;
      }
    });
  }
  return n;
}

function countAImpl(a: Value[], c: FnContext): number {
  let n = 0;
  for (const v of a) forEachValue(v, c, (s, fromRef) => {
    if (s !== null || !fromRef) n++;
  });
  return n;
}

function countBlank(a: Value[], c: FnContext): number {
  const v = a[0];
  if (isRef(v)) {
    const area = (v.r2 - v.r1 + 1) * (v.c2 - v.c1 + 1);
    let filled = 0;
    const cl = clipRef(v, c);
    if (cl) for (let i = cl.r1; i <= cl.r2; i++) for (let j = cl.c1; j <= cl.c2; j++) {
      const s = c.host.getCellValue(cl.sheet, i, j);
      if (s !== null && s !== "") filled++;
    }
    return area - filled;
  }
  return matrix(v, c).flat().filter((s) => s === null || s === "").length;
}

function countIf(a: Value[], c: FnContext): number {
  const crit = makeCriteria(scalar(a[1], c), c.locale);
  const v = a[0];
  let n = 0;
  if (isRef(v)) {
    const cl = clipRef(v, c);
    if (cl) for (let i = cl.r1; i <= cl.r2; i++) for (let j = cl.c1; j <= cl.c2; j++) if (crit(c.host.getCellValue(cl.sheet, i, j))) n++;
    // Kullanılan alanın dışındaki boş hücreler
    if (crit(null)) {
      const area = (v.r2 - v.r1 + 1) * (v.c2 - v.c1 + 1);
      const inside = cl ? (cl.r2 - cl.r1 + 1) * (cl.c2 - cl.c1 + 1) : 0;
      n += area - inside;
    }
    return n;
  }
  for (const s of matrix(v, c).flat()) if (crit(s)) n++;
  return n;
}

function ifsValues(a: Value[], c: FnContext, start: number): number[] {
  if ((a.length - start) % 2) throw Errors.VALUE;
  const ps: [Value, Value][] = [];
  for (let i = start; i < a.length; i += 2) ps.push([a[i], a[i + 1]]);
  const { mask, rows, cols } = matchCriteriaRanges(ps, c);
  const out: number[] = [];
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++) {
      if (!mask[i][j]) continue;
      const v = valueAt(a[0], i, j, c);
      if (typeof v === "number") out.push(v);
      else if (isError(v)) throw v;
    }
  return out;
}

function averageIf(a: Value[], c: FnContext): number {
  const crit = makeCriteria(scalar(a[1], c), c.locale);
  const range = a[0];
  const avgRange = a[2] ?? range;
  const vals: number[] = [];
  const data = isRef(range) ? (() => {
    const cl = clipRef(range, c);
    return cl ? matrix(cl, c) : [];
  })() : matrix(range, c);
  data.forEach((row, i) =>
    row.forEach((s, j) => {
      if (!crit(s)) return;
      const v = valueAt(avgRange, i, j, c);
      if (typeof v === "number") vals.push(v);
    }),
  );
  return agg.average(vals);
}

function rankImpl(a: Value[], c: FnContext, avg: boolean): number {
  const x = num(a[0], c);
  const list = collectNumbers([a[1]], c);
  const asc = optNum(a[2], 0, c) !== 0;
  if (!list.some((v) => v === x)) throw Errors.NA;
  let better = 0;
  let same = 0;
  for (const v of list) {
    if (v === x) same++;
    else if (asc ? v < x : v > x) better++;
  }
  return avg ? better + (same + 1) / 2 : better + 1;
}

function percentRank(a: Value[], c: FnContext, exc: boolean): number {
  const xs = nums([a[0]], c).sort((p, q) => p - q);
  const x = num(a[1], c);
  const sig = Math.trunc(optNum(a[2], 3, c));
  if (sig < 1 || xs.length === 0) throw Errors.NUM;
  const n = xs.length;
  if (x < xs[0] || x > xs[n - 1]) throw Errors.NA;
  let rank: number;
  const idx = xs.findIndex((v) => v >= x);
  const lessCount = (v: number) => xs.filter((y) => y < v).length;
  if (xs[idx] === x) rank = exc ? (lessCount(x) + 1) / (n + 1) : lessCount(x) / (n - 1 || 1);
  else {
    const lo = xs[idx - 1];
    const hi = xs[idx];
    const rlo = exc ? (lessCount(lo) + 1) / (n + 1) : lessCount(lo) / (n - 1);
    const rhi = exc ? (lessCount(hi) + 1) / (n + 1) : lessCount(hi) / (n - 1);
    rank = rlo + ((x - lo) / (hi - lo)) * (rhi - rlo);
  }
  const f = Math.pow(10, sig);
  return Math.floor(rank * f + 1e-9) / f;
}

/* Normal dağılım yardımcıları */
function erf(x: number): number {
  // Abramowitz–Stegun 7.1.26 yerine yüksek doğruluklu seri/sürekli kesir
  const t = 1 / (1 + 0.5 * Math.abs(x));
  const y =
    1 -
    t *
      Math.exp(
        -x * x -
          1.26551223 +
          t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))),
      );
  return x >= 0 ? y : -y;
}

export function normSDist(z: number): number {
  // Daha hassas: Hart algoritması (W. J. Cody)
  const a = Math.abs(z);
  let p: number;
  if (a > 37) p = 0;
  else {
    const e = Math.exp(-a * a / 2);
    if (a < 7.07106781186547) {
      let b = 3.52624965998911e-2 * a + 0.700383064443688;
      b = b * a + 6.37396220353165;
      b = b * a + 33.912866078383;
      b = b * a + 112.079291497871;
      b = b * a + 221.213596169931;
      b = b * a + 220.206867912376;
      let d = 8.83883476483184e-2 * a + 1.75566716318264;
      d = d * a + 16.064177579207;
      d = d * a + 86.7807322029461;
      d = d * a + 296.564248779674;
      d = d * a + 637.333633378831;
      d = d * a + 793.826512519948;
      d = d * a + 440.413735824752;
      p = (e * b) / d;
    } else {
      let b = a + 0.65;
      b = a + 4 / b;
      b = a + 3 / b;
      b = a + 2 / b;
      b = a + 1 / b;
      p = e / b / 2.506628274631;
    }
  }
  return z > 0 ? 1 - p : p;
}

export function normSInv(p: number): number {
  if (p <= 0 || p >= 1) throw Errors.NUM;
  // Acklam algoritması
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const cc = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  let q: number;
  let x: number;
  if (p < pl) {
    q = Math.sqrt(-2 * Math.log(p));
    x = (((((cc[0] * q + cc[1]) * q + cc[2]) * q + cc[3]) * q + cc[4]) * q + cc[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - pl) {
    q = p - 0.5;
    const r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((cc[0] * q + cc[1]) * q + cc[2]) * q + cc[3]) * q + cc[4]) * q + cc[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  // Newton düzeltmesi
  const e = normSDist(x) - p;
  const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
  return x - u / (1 + (x * u) / 2);
}

function normDist(x: number, m: number, s: number, cum: boolean): number {
  if (s <= 0) throw Errors.NUM;
  const z = (x - m) / s;
  return cum ? normSDist(z) : Math.exp((-z * z) / 2) / (s * Math.sqrt(2 * Math.PI));
}

function lnGamma(x: number): number {
  return lnGammaHP(x);
}

function binomPmf(k: number, n: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === n ? 1 : 0;
  return Math.exp(lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1) + k * Math.log(p) + (n - k) * Math.log(1 - p));
}

function stdevAImpl(a: Value[], c: FnContext, sample: boolean) {
  return agg.stdev(numsA(a, c), sample);
}

export const statsFunctions: FnDef[] = [
  fn("AVERAGE", "ORTALAMA", "stats", "sayı1; [sayı2]; ...", "Bağımsız değişkenlerin ortalamasını (aritmetik ortalama) verir.", 1, 255, (a, c) => agg.average(nums(a, c))),
  fn("AVERAGEA", "ORTALAMAA", "stats", "değer1; [değer2]; ...", "Metin ve mantıksal değerleri de hesaba katarak ortalamayı verir.", 1, 255, (a, c) => agg.average(numsA(a, c))),
  fn("AVERAGEIF", "EĞERORTALAMA", "stats", "aralık; ölçüt; [ortalama_aralığı]", "Koşulu karşılayan hücrelerin ortalamasını verir.", 2, 3, averageIf),
  fn("AVERAGEIFS", "ÇOKEĞERORTALAMA", "stats", "ortalama_aralığı; ölçüt_aralığı1; ölçüt1; ...", "Birden çok koşulu karşılayan hücrelerin ortalamasını verir.", 3, 255, (a, c) => agg.average(ifsValues(a, c, 1))),
  fn("COUNT", "BAĞ_DEĞ_SAY", "stats", "değer1; [değer2]; ...", "Aralıktaki sayı içeren hücreleri sayar.", 1, 255, countImpl),
  fn("COUNTA", "BAĞ_DEĞ_DOLU_SAY", "stats", "değer1; [değer2]; ...", "Aralıktaki boş olmayan hücreleri sayar.", 1, 255, countAImpl),
  fn("COUNTBLANK", "BOŞLUKSAY", "stats", "aralık", "Aralıktaki boş hücreleri sayar.", 1, 1, countBlank),
  fn("COUNTIF", "EĞERSAY", "stats", "aralık; ölçüt", "Verilen koşula uyan hücreleri sayar.", 2, 2, countIf),
  fn("COUNTIFS", "ÇOKEĞERSAY", "stats", "ölçüt_aralığı1; ölçüt1; ...", "Birden çok koşula uyan hücreleri sayar.", 2, 255, (a, c) => {
    if (a.length % 2) throw Errors.VALUE;
    const ps: [Value, Value][] = [];
    for (let i = 0; i < a.length; i += 2) ps.push([a[i], a[i + 1]]);
    const { mask } = matchCriteriaRanges(ps, c);
    let n = 0;
    for (const r of mask) for (const m of r) if (m) n++;
    return n;
  }),
  fn("MIN", "MİN", "stats", "sayı1; [sayı2]; ...", "Değerler kümesindeki en küçük sayıyı verir.", 1, 255, (a, c) => agg.min(nums(a, c))),
  fn("MAX", "MAK", "stats", "sayı1; [sayı2]; ...", "Değerler kümesindeki en büyük sayıyı verir.", 1, 255, (a, c) => agg.max(nums(a, c))),
  fn("MINA", "MİNA", "stats", "değer1; [değer2]; ...", "Mantıksal değer ve metinleri de içeren en küçük değeri verir.", 1, 255, (a, c) => agg.min(numsA(a, c))),
  fn("MAXA", "MAKA", "stats", "değer1; [değer2]; ...", "Mantıksal değer ve metinleri de içeren en büyük değeri verir.", 1, 255, (a, c) => agg.max(numsA(a, c))),
  fn("MINIFS", "ÇOKEĞERMİN", "stats", "min_aralığı; ölçüt_aralığı1; ölçüt1; ...", "Koşullara uyan hücrelerin en küçüğünü verir.", 3, 255, (a, c) => agg.min(ifsValues(a, c, 1)), { xlfn: true }),
  fn("MAXIFS", "ÇOKEĞERMAK", "stats", "maks_aralığı; ölçüt_aralığı1; ölçüt1; ...", "Koşullara uyan hücrelerin en büyüğünü verir.", 3, 255, (a, c) => agg.max(ifsValues(a, c, 1)), { xlfn: true }),
  fn("MEDIAN", "ORTANCA", "stats", "sayı1; [sayı2]; ...", "Verilen sayıların ortancasını verir.", 1, 255, (a, c) => agg.median(nums(a, c))),
  fn("MODE", "ENÇOK_OLAN", "stats", "sayı1; [sayı2]; ...", "En sık görülen değeri verir.", 1, 255, (a, c) => agg.modes(nums(a, c))[0]),
  fn("MODE.SNGL", "ENÇOK_OLAN.TEK", "stats", "sayı1; [sayı2]; ...", "En sık görülen değeri verir.", 1, 255, (a, c) => agg.modes(nums(a, c))[0], { xlfn: true }),
  fn("MODE.MULT", "ENÇOK_OLAN.ÇOK", "stats", "sayı1; [sayı2]; ...", "En sık görülen değerlerin dikey dizisini verir.", 1, 255, (a, c) => makeArray(agg.modes(nums(a, c)).map((x) => [x])), { xlfn: true }),
  fn("STDEV", "STDSAPMA", "stats", "sayı1; [sayı2]; ...", "Örneğe dayalı standart sapmayı tahmin eder.", 1, 255, (a, c) => agg.stdev(nums(a, c), true)),
  fn("STDEV.S", "STDSAPMA.S", "stats", "sayı1; [sayı2]; ...", "Örneğe dayalı standart sapmayı tahmin eder.", 1, 255, (a, c) => agg.stdev(nums(a, c), true), { xlfn: true }),
  fn("STDEV.P", "STDSAPMA.P", "stats", "sayı1; [sayı2]; ...", "Tüm popülasyona dayalı standart sapmayı hesaplar.", 1, 255, (a, c) => agg.stdev(nums(a, c), false), { xlfn: true }),
  fn("STDEVP", "STDSAPMAS", "stats", "sayı1; [sayı2]; ...", "Tüm popülasyona dayalı standart sapmayı hesaplar.", 1, 255, (a, c) => agg.stdev(nums(a, c), false)),
  fn("STDEVA", "STDSAPMAA", "stats", "değer1; [değer2]; ...", "Metin ve mantıksal değerleri içeren örneğe dayalı standart sapma.", 1, 255, (a, c) => stdevAImpl(a, c, true)),
  fn("STDEVPA", "STDSAPMASA", "stats", "değer1; [değer2]; ...", "Metin ve mantıksal değerleri içeren popülasyon standart sapması.", 1, 255, (a, c) => stdevAImpl(a, c, false)),
  fn("VAR", "VAR", "stats", "sayı1; [sayı2]; ...", "Örneğe dayalı varyansı tahmin eder.", 1, 255, (a, c) => agg.variance(nums(a, c), true)),
  fn("VAR.S", "VAR.S", "stats", "sayı1; [sayı2]; ...", "Örneğe dayalı varyansı tahmin eder.", 1, 255, (a, c) => agg.variance(nums(a, c), true), { xlfn: true }),
  fn("VAR.P", "VAR.P", "stats", "sayı1; [sayı2]; ...", "Tüm popülasyona dayalı varyansı hesaplar.", 1, 255, (a, c) => agg.variance(nums(a, c), false), { xlfn: true }),
  fn("VARP", "VARS", "stats", "sayı1; [sayı2]; ...", "Tüm popülasyona dayalı varyansı hesaplar.", 1, 255, (a, c) => agg.variance(nums(a, c), false)),
  fn("VARA", "VARA", "stats", "değer1; [değer2]; ...", "Metin ve mantıksal değerleri içeren örnek varyansı.", 1, 255, (a, c) => agg.variance(numsA(a, c), true)),
  fn("VARPA", "VARSA", "stats", "değer1; [değer2]; ...", "Metin ve mantıksal değerleri içeren popülasyon varyansı.", 1, 255, (a, c) => agg.variance(numsA(a, c), false)),
  fn("LARGE", "BÜYÜK", "stats", "dizi; k", "Veri kümesindeki k. en büyük değeri verir.", 2, 2, (a, c) => kthLifted(a, c, true)),
  fn("SMALL", "KÜÇÜK", "stats", "dizi; k", "Veri kümesindeki k. en küçük değeri verir.", 2, 2, (a, c) => kthLifted(a, c, false)),
  fn("RANK", "RANK", "stats", "sayı; başv; [sıra]", "Sayının listedeki sırasını verir.", 2, 3, (a, c) => rankImpl(a, c, false)),
  fn("RANK.EQ", "RANK.EŞİT", "stats", "sayı; başv; [sıra]", "Sayının listedeki sırasını verir.", 2, 3, (a, c) => rankImpl(a, c, false), { xlfn: true }),
  fn("RANK.AVG", "RANK.ORT", "stats", "sayı; başv; [sıra]", "Eşit değerler için ortalama sırayı verir.", 2, 3, (a, c) => rankImpl(a, c, true), { xlfn: true }),
  fn("PERCENTILE", "YÜZDEBİRLİK", "stats", "dizi; k", "Aralıktaki değerlerin k. yüzdebirliğini verir.", 2, 2, (a, c) => agg.percentileInc(nums([a[0]], c), num(a[1], c))),
  fn("PERCENTILE.INC", "YÜZDEBİRLİK.DHL", "stats", "dizi; k", "k. yüzdebirliği verir (0 ve 1 dahil).", 2, 2, (a, c) => agg.percentileInc(nums([a[0]], c), num(a[1], c)), { xlfn: true }),
  fn("PERCENTILE.EXC", "YÜZDEBİRLİK.HRC", "stats", "dizi; k", "k. yüzdebirliği verir (0 ve 1 hariç).", 2, 2, (a, c) => agg.percentileExc(nums([a[0]], c), num(a[1], c)), { xlfn: true }),
  fn("QUARTILE", "DÖRTTEBİRLİK", "stats", "dizi; dörttebir", "Veri kümesinin dörttebirliğini verir.", 2, 2, (a, c) => quartile(a, c, false)),
  fn("QUARTILE.INC", "DÖRTTEBİRLİK.DHL", "stats", "dizi; dörttebir", "Veri kümesinin dörttebirliğini verir.", 2, 2, (a, c) => quartile(a, c, false), { xlfn: true }),
  fn("QUARTILE.EXC", "DÖRTTEBİRLİK.HRC", "stats", "dizi; dörttebir", "Veri kümesinin dörttebirliğini verir (hariç).", 2, 2, (a, c) => quartile(a, c, true), { xlfn: true }),
  fn("PERCENTRANK", "YÜZDERANK", "stats", "dizi; x; [anlam]", "Değerin veri kümesindeki yüzde sırasını verir.", 2, 3, (a, c) => percentRank(a, c, false)),
  fn("PERCENTRANK.INC", "YÜZDERANK.DHL", "stats", "dizi; x; [anlam]", "Değerin yüzde sırasını verir (dahil).", 2, 3, (a, c) => percentRank(a, c, false), { xlfn: true }),
  fn("PERCENTRANK.EXC", "YÜZDERANK.HRC", "stats", "dizi; x; [anlam]", "Değerin yüzde sırasını verir (hariç).", 2, 3, (a, c) => percentRank(a, c, true), { xlfn: true }),
  fn("CORREL", "KORELASYON", "stats", "dizi1; dizi2", "İki veri kümesi arasındaki korelasyon katsayısını verir.", 2, 2, (a, c) => agg.correl(...pairs(a[0], a[1], c))),
  fn("PEARSON", "PEARSON", "stats", "dizi1; dizi2", "Pearson çarpım moment korelasyon katsayısını verir.", 2, 2, (a, c) => agg.correl(...pairs(a[0], a[1], c))),
  fn("RSQ", "RKARE", "stats", "bilinen_y'ler; bilinen_x'ler", "Pearson katsayısının karesini verir.", 2, 2, (a, c) => agg.correl(...pairs(a[1], a[0], c)) ** 2),
  fn("COVAR", "KOVARYANS", "stats", "dizi1; dizi2", "Kovaryansı verir.", 2, 2, (a, c) => agg.covariance(...pairs(a[0], a[1], c), false)),
  fn("COVARIANCE.P", "KOVARYANS.P", "stats", "dizi1; dizi2", "Popülasyon kovaryansını verir.", 2, 2, (a, c) => agg.covariance(...pairs(a[0], a[1], c), false), { xlfn: true }),
  fn("COVARIANCE.S", "KOVARYANS.S", "stats", "dizi1; dizi2", "Örnek kovaryansını verir.", 2, 2, (a, c) => agg.covariance(...pairs(a[0], a[1], c), true), { xlfn: true }),
  fn("SLOPE", "EĞİM", "stats", "bilinen_y'ler; bilinen_x'ler", "Doğrusal regresyon çizgisinin eğimini verir.", 2, 2, (a, c) => regression(a[0], a[1], c).slope),
  fn("INTERCEPT", "KESMENOKTASI", "stats", "bilinen_y'ler; bilinen_x'ler", "Doğrusal regresyon çizgisinin y eksenini kestiği noktayı verir.", 2, 2, (a, c) => regression(a[0], a[1], c).intercept),
  fn("FORECAST", "TAHMİN", "stats", "x; bilinen_y'ler; bilinen_x'ler", "Doğrusal eğilime göre değer tahmin eder.", 3, 3, (a, c) => {
    const r = regression(a[1], a[2], c);
    return r.intercept + r.slope * num(a[0], c);
  }),
  fn("FORECAST.LINEAR", "TAHMİN.DOĞRUSAL", "stats", "x; bilinen_y'ler; bilinen_x'ler", "Doğrusal eğilime göre değer tahmin eder.", 3, 3, (a, c) => {
    const r = regression(a[1], a[2], c);
    return r.intercept + r.slope * num(a[0], c);
  }, { xlfn: true }),
  fn("STEYX", "STHYX", "stats", "bilinen_y'ler; bilinen_x'ler", "Regresyondaki tahmin edilen y değerinin standart hatasını verir.", 2, 2, (a, c) => {
    const [ys, xs] = pairs(a[0], a[1], c);
    const n = xs.length;
    if (n < 3) throw Errors.DIV0;
    const r = regression(a[0], a[1], c);
    let s = 0;
    for (let i = 0; i < n; i++) s += (ys[i] - (r.intercept + r.slope * xs[i])) ** 2;
    return Math.sqrt(s / (n - 2));
  }),
  fn("GEOMEAN", "GEOORT", "stats", "sayı1; [sayı2]; ...", "Geometrik ortalamayı verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    if (xs.length === 0 || xs.some((x) => x <= 0)) throw Errors.NUM;
    return Math.exp(agg.sum(xs.map(Math.log)) / xs.length);
  }),
  fn("HARMEAN", "HARORT", "stats", "sayı1; [sayı2]; ...", "Harmonik ortalamayı verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    if (xs.length === 0 || xs.some((x) => x <= 0)) throw Errors.NUM;
    return xs.length / agg.sum(xs.map((x) => 1 / x));
  }),
  fn("AVEDEV", "ORTSAP", "stats", "sayı1; [sayı2]; ...", "Veri noktalarının ortalamalarından mutlak sapmalarının ortalamasını verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    const m = agg.average(xs);
    return agg.sum(xs.map((x) => Math.abs(x - m))) / xs.length;
  }),
  fn("DEVSQ", "SAPKARE", "stats", "sayı1; [sayı2]; ...", "Sapmaların karelerinin toplamını verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    if (!xs.length) throw Errors.NUM;
    const m = agg.average(xs);
    return agg.sum(xs.map((x) => (x - m) ** 2));
  }),
  fn("KURT", "BASIKLIK", "stats", "sayı1; [sayı2]; ...", "Veri kümesinin basıklığını verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    const n = xs.length;
    if (n < 4) throw Errors.DIV0;
    const m = agg.average(xs);
    const s = agg.stdev(xs, true);
    if (s === 0) throw Errors.DIV0;
    const k = agg.sum(xs.map((x) => ((x - m) / s) ** 4));
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  }),
  fn("SKEW", "ÇARPIKLIK", "stats", "sayı1; [sayı2]; ...", "Dağılımın çarpıklığını verir.", 1, 255, (a, c) => {
    const xs = nums(a, c);
    const n = xs.length;
    if (n < 3) throw Errors.DIV0;
    const m = agg.average(xs);
    const s = agg.stdev(xs, true);
    if (s === 0) throw Errors.DIV0;
    return (n / ((n - 1) * (n - 2))) * agg.sum(xs.map((x) => ((x - m) / s) ** 3));
  }),
  fn("STANDARDIZE", "STANDARTLAŞTIRMA", "stats", "x; ortalama; standart_sapma", "Normalleştirilmiş değeri verir.", 3, 3, (a, c) => {
    const s = num(a[2], c);
    if (s <= 0) throw Errors.NUM;
    return (num(a[0], c) - num(a[1], c)) / s;
  }, { lift: true }),
  fn("TRIMMEAN", "KIRPORTALAMA", "stats", "dizi; yüzde", "Veri kümesinin iç kısmının ortalamasını verir.", 2, 2, (a, c) => {
    const xs = nums([a[0]], c).sort((p, q) => p - q);
    const pct = num(a[1], c);
    if (pct < 0 || pct >= 1) throw Errors.NUM;
    const k = Math.floor((xs.length * pct) / 2);
    return agg.average(xs.slice(k, xs.length - k));
  }),
  fn("FREQUENCY", "SIKLIK", "stats", "veri_dizisi; bölme_dizisi", "Değerlerin aralıklarda kaç kez geçtiğini dikey dizi olarak verir.", 2, 2, (a, c) => {
    const data = nums([a[0]], c);
    const bins = nums([a[1]], c);
    const order = bins.map((b, i) => [b, i] as [number, number]).sort((p, q) => p[0] - q[0]);
    const counts = new Array(bins.length + 1).fill(0);
    for (const x of data) {
      const k = order.findIndex(([b]) => x <= b);
      counts[k < 0 ? bins.length : order[k][1]]++;
    }
    return makeArray(counts.map((n) => [n]));
  }),
  fn("NORM.DIST", "NORM.DAĞ", "stats", "x; ortalama; standart_sapma; kümülatif", "Normal dağılımı verir.", 4, 4, (a, c) => normDist(num(a[0], c), num(a[1], c), num(a[2], c), !!num(a[3], c)), { lift: true, xlfn: true }),
  fn("NORMDIST", "NORMDAĞ", "stats", "x; ortalama; standart_sapma; kümülatif", "Normal dağılımı verir.", 4, 4, (a, c) => normDist(num(a[0], c), num(a[1], c), num(a[2], c), !!num(a[3], c)), { lift: true }),
  fn("NORM.INV", "NORM.TERS", "stats", "olasılık; ortalama; standart_sapma", "Normal dağılımın tersini verir.", 3, 3, (a, c) => {
    const s = num(a[2], c);
    if (s <= 0) throw Errors.NUM;
    return num(a[1], c) + s * normSInv(num(a[0], c));
  }, { lift: true, xlfn: true }),
  fn("NORMINV", "NORMTERS", "stats", "olasılık; ortalama; standart_sapma", "Normal dağılımın tersini verir.", 3, 3, (a, c) => {
    const s = num(a[2], c);
    if (s <= 0) throw Errors.NUM;
    return num(a[1], c) + s * normSInv(num(a[0], c));
  }, { lift: true }),
  fn("NORM.S.DIST", "NORM.S.DAĞ", "stats", "z; kümülatif", "Standart normal dağılımı verir.", 2, 2, (a, c) => normDist(num(a[0], c), 0, 1, !!num(a[1], c)), { lift: true, xlfn: true }),
  fn("NORMSDIST", "NORMSDAĞ", "stats", "z", "Standart normal kümülatif dağılımı verir.", 1, 1, (a, c) => normSDist(num(a[0], c)), { lift: true }),
  fn("NORM.S.INV", "NORM.S.TERS", "stats", "olasılık", "Standart normal dağılımın tersini verir.", 1, 1, (a, c) => normSInv(num(a[0], c)), { lift: true, xlfn: true }),
  fn("NORMSINV", "NORMSTERS", "stats", "olasılık", "Standart normal dağılımın tersini verir.", 1, 1, (a, c) => normSInv(num(a[0], c)), { lift: true }),
  fn("BINOM.DIST", "BİNOM.DAĞ", "stats", "başarı_sayısı; denemeler; başarı_olasılığı; kümülatif", "Binom dağılımı olasılığını verir.", 4, 4, (a, c) => {
    const k = Math.floor(num(a[0], c));
    const n = Math.floor(num(a[1], c));
    const p = num(a[2], c);
    if (k < 0 || k > n || p < 0 || p > 1) throw Errors.NUM;
    if (!num(a[3], c)) return binomPmf(k, n, p);
    let s = 0;
    for (let i = 0; i <= k; i++) s += binomPmf(i, n, p);
    return Math.min(1, s);
  }, { lift: true, xlfn: true }),
  fn("POISSON.DIST", "POISSON.DAĞ", "stats", "x; ortalama; kümülatif", "Poisson dağılımını verir.", 3, 3, (a, c) => {
    const x = Math.floor(num(a[0], c));
    const m = num(a[1], c);
    if (x < 0 || m < 0) throw Errors.NUM;
    const pmf = (k: number) => Math.exp(-m + k * Math.log(m || 1e-300) - lnGamma(k + 1));
    if (!num(a[2], c)) return m === 0 ? (x === 0 ? 1 : 0) : pmf(x);
    let s = 0;
    for (let k = 0; k <= x; k++) s += m === 0 ? (k === 0 ? 1 : 0) : pmf(k);
    return Math.min(1, s);
  }, { lift: true, xlfn: true }),
  fn("EXPON.DIST", "ÜSTEL.DAĞ", "stats", "x; lambda; kümülatif", "Üstel dağılımı verir.", 3, 3, (a, c) => {
    const x = num(a[0], c);
    const l = num(a[1], c);
    if (x < 0 || l <= 0) throw Errors.NUM;
    return num(a[2], c) ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x);
  }, { lift: true, xlfn: true }),
  fn("CONFIDENCE.NORM", "GÜVENİLİRLİK.NORM", "stats", "alfa; standart_sapma; boyut", "Normal dağılımla popülasyon ortalaması için güven aralığını verir.", 3, 3, (a, c) => {
    const alpha = num(a[0], c);
    const s = num(a[1], c);
    const n = Math.floor(num(a[2], c));
    if (alpha <= 0 || alpha >= 1 || s <= 0 || n < 1) throw Errors.NUM;
    return (normSInv(1 - alpha / 2) * s) / Math.sqrt(n);
  }, { xlfn: true }),
  fn("GAMMALN", "GAMALN", "stats", "x", "Gama fonksiyonunun doğal logaritmasını verir.", 1, 1, (a, c) => {
    const x = num(a[0], c);
    if (x <= 0) throw Errors.NUM;
    return lnGamma(x);
  }, { lift: true }),
];

function kthLifted(a: Value[], c: FnContext, largest: boolean): Value {
  const xs = nums([a[0]], c);
  const kv = a[1];
  const ks = matrix(kv, c);
  if (ks.length === 1 && ks[0].length === 1) return agg.kth(xs, num(ks[0][0] as Value, c), largest);
  return makeArray(ks.map((r) => r.map((k): Scalar => {
    try {
      return agg.kth(xs, num(k as Value, c), largest);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  })));
}

function quartile(a: Value[], c: FnContext, exc: boolean): number {
  const q = Math.floor(num(a[1], c));
  const xs = nums([a[0]], c);
  if (exc) {
    if (q <= 0 || q >= 4) throw Errors.NUM;
    return agg.percentileExc(xs, q / 4);
  }
  if (q < 0 || q > 4) throw Errors.NUM;
  return agg.percentileInc(xs, q / 4);
}

function regression(yv: Value, xv: Value, c: FnContext): { slope: number; intercept: number } {
  const [ys, xs] = pairs(yv, xv, c);
  const n = xs.length;
  if (n === 0) throw Errors.DIV0;
  const mx = agg.sum(xs) / n;
  const my = agg.sum(ys) / n;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  if (sxx === 0) throw Errors.DIV0;
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx };
}

void flatValues;
void erf;
