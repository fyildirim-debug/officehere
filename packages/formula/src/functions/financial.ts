/** Finansal fonksiyonlar. */
import type { FnContext } from "../context";
import { Errors, isError, type Value } from "../values";
import { collectNumbers, matrix, num, optNum } from "./helpers";
import { fn, type FnDef } from "./registry";

export function pmt(r: number, n: number, pv: number, fv = 0, type = 0): number {
  if (n === 0) throw Errors.NUM;
  if (r === 0) return -(pv + fv) / n;
  const f = Math.pow(1 + r, n);
  return (-(r * (fv + pv * f)) / ((1 + r * type) * (f - 1)));
}

export function fv(r: number, n: number, p: number, pv = 0, type = 0): number {
  if (r === 0) return -(pv + p * n);
  const f = Math.pow(1 + r, n);
  return -(pv * f + (p * (1 + r * type) * (f - 1)) / r);
}

export function pv(r: number, n: number, p: number, fvv = 0, type = 0): number {
  if (r === 0) return -(fvv + p * n);
  const f = Math.pow(1 + r, n);
  return -(fvv + (p * (1 + r * type) * (f - 1)) / r) / f;
}

function nper(r: number, p: number, pvv: number, fvv = 0, type = 0): number {
  if (r === 0) {
    if (p === 0) throw Errors.NUM;
    return -(pvv + fvv) / p;
  }
  const num1 = p * (1 + r * type) - fvv * r;
  const den = pvv * r + p * (1 + r * type);
  if (num1 / den <= 0) throw Errors.NUM;
  return Math.log(num1 / den) / Math.log(1 + r);
}

function ipmt(r: number, per: number, n: number, pvv: number, fvv = 0, type = 0): number {
  if (per < 1 || per > n) throw Errors.NUM;
  const p = pmt(r, n, pvv, fvv, type);
  let ip: number;
  if (per === 1) ip = type === 1 ? 0 : -pvv;
  else ip = type === 1 ? fv(r, per - 2, p, pvv, 1) - p : fv(r, per - 1, p, pvv, 0);
  return ip * r;
}

/** Newton–Raphson ile kök bulma. */
function newton(f: (x: number) => number, guess: number, df?: (x: number) => number): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) break;
    if (Math.abs(y) < 1e-10) return x;
    const d = df ? df(x) : (f(x + 1e-7) - y) / 1e-7;
    if (d === 0 || !Number.isFinite(d)) break;
    const nx = x - y / d;
    if (Math.abs(nx - x) < 1e-12) return nx;
    x = nx;
  }
  if (Math.abs(f(x)) < 1e-6) return x;
  throw Errors.NUM;
}

function irr(values: number[], guess: number): number {
  if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) throw Errors.NUM;
  const f = (r: number) => values.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0);
  const df = (r: number) => values.reduce((s, v, i) => s - (i * v) / Math.pow(1 + r, i + 1), 0);
  return newton(f, guess, df);
}

function cashflows(v: Value, c: FnContext): number[] {
  const out: number[] = [];
  for (const s of matrix(v, c).flat()) {
    if (isError(s)) throw s;
    if (typeof s === "number") out.push(s);
  }
  return out;
}

function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0) throw Errors.NUM;
  let book = cost;
  let dep = 0;
  for (let p = 1; p <= Math.ceil(period); p++) {
    dep = Math.min(book * (factor / life), Math.max(0, book - salvage));
    book -= dep;
  }
  return dep;
}

function db(cost: number, salvage: number, life: number, period: number, month: number): number {
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + 1) throw Errors.NUM;
  if (cost === 0) return 0;
  const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
  let total = 0;
  let dep = (cost * rate * month) / 12;
  if (period === 1) return dep;
  total = dep;
  for (let p = 2; p <= period; p++) {
    if (p === Math.floor(life) + 1) dep = ((cost - total) * rate * (12 - month)) / 12;
    else dep = (cost - total) * rate;
    total += dep;
  }
  return dep;
}

export const financialFunctions: FnDef[] = [
  fn("PMT", "DEVRESEL_ÖDEME", "financial", "oran; dönem_sayısı; bd; [gd]; [tür]", "Sabit ödemeli ve sabit faizli bir borcun dönemsel ödemesini hesaplar.", 3, 5, (a, c) => pmt(num(a[0], c), num(a[1], c), num(a[2], c), optNum(a[3], 0, c), optNum(a[4], 0, c) ? 1 : 0), { lift: true }),
  fn("FV", "GD", "financial", "oran; dönem_sayısı; devresel_ödeme; [bd]; [tür]", "Yatırımın gelecek değerini verir.", 3, 5, (a, c) => fv(num(a[0], c), num(a[1], c), num(a[2], c), optNum(a[3], 0, c), optNum(a[4], 0, c) ? 1 : 0), { lift: true }),
  fn("PV", "BD", "financial", "oran; dönem_sayısı; devresel_ödeme; [gd]; [tür]", "Yatırımın bugünkü değerini verir.", 3, 5, (a, c) => pv(num(a[0], c), num(a[1], c), num(a[2], c), optNum(a[3], 0, c), optNum(a[4], 0, c) ? 1 : 0), { lift: true }),
  fn("NPER", "TAKSİT_SAYISI", "financial", "oran; devresel_ödeme; bd; [gd]; [tür]", "Yatırımın dönem sayısını verir.", 3, 5, (a, c) => nper(num(a[0], c), num(a[1], c), num(a[2], c), optNum(a[3], 0, c), optNum(a[4], 0, c) ? 1 : 0), { lift: true }),
  fn("NPV", "NBD", "financial", "oran; değer1; [değer2]; ...", "Dönemsel nakit akışlarının net bugünkü değerini verir.", 2, 255, (a, c) => {
    const r = num(a[0], c);
    if (r === -1) throw Errors.DIV0;
    return collectNumbers(a.slice(1), c).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0);
  }),
  fn("IRR", "İÇ_VERİM_ORANI", "financial", "değerler; [tahmin]", "Nakit akışlarının iç verim oranını verir.", 1, 2, (a, c) => irr(cashflows(a[0], c), optNum(a[1], 0.1, c))),
  fn("RATE", "FAİZ_ORANI", "financial", "dönem_sayısı; devresel_ödeme; bd; [gd]; [tür]; [tahmin]", "Bir yatırımın dönem başına faiz oranını verir.", 3, 6, (a, c) => {
    const n = num(a[0], c);
    const p = num(a[1], c);
    const pvv = num(a[2], c);
    const fvv = optNum(a[3], 0, c);
    const type = optNum(a[4], 0, c) ? 1 : 0;
    const guess = optNum(a[5], 0.1, c);
    if (n <= 0) throw Errors.NUM;
    const f = (r: number) => {
      if (Math.abs(r) < 1e-12) return pvv + p * n + fvv;
      const g = Math.pow(1 + r, n);
      return pvv * g + (p * (1 + r * type) * (g - 1)) / r + fvv;
    };
    return newton(f, guess);
  }, { lift: true }),
  fn("IPMT", "FAİZTUTARI", "financial", "oran; dönem; dönem_sayısı; bd; [gd]; [tür]", "Belirli bir dönemdeki faiz ödemesini verir.", 4, 6, (a, c) => ipmt(num(a[0], c), num(a[1], c), num(a[2], c), num(a[3], c), optNum(a[4], 0, c), optNum(a[5], 0, c) ? 1 : 0), { lift: true }),
  fn("PPMT", "ANA_PARA_ÖDEMESİ", "financial", "oran; dönem; dönem_sayısı; bd; [gd]; [tür]", "Belirli bir dönemdeki anapara ödemesini verir.", 4, 6, (a, c) => {
    const r = num(a[0], c);
    const per = num(a[1], c);
    const n = num(a[2], c);
    const pvv = num(a[3], c);
    const fvv = optNum(a[4], 0, c);
    const type = optNum(a[5], 0, c) ? 1 : 0;
    return pmt(r, n, pvv, fvv, type) - ipmt(r, per, n, pvv, fvv, type);
  }, { lift: true }),
  fn("CUMIPMT", "TOPÖDENENFAİZ", "financial", "oran; dönem_sayısı; bd; başlangıç_dönemi; bitiş_dönemi; tür", "İki dönem arasında ödenen toplam faizi verir.", 6, 6, (a, c) => {
    const [r, n, pvv, s, e, t] = a.map((x) => num(x, c));
    if (r <= 0 || n <= 0 || pvv <= 0 || s < 1 || e < s || e > n) throw Errors.NUM;
    let sum = 0;
    for (let p = Math.ceil(s); p <= Math.floor(e); p++) sum += ipmt(r, p, n, pvv, 0, t ? 1 : 0);
    return sum;
  }),
  fn("CUMPRINC", "TOPANAPARA", "financial", "oran; dönem_sayısı; bd; başlangıç_dönemi; bitiş_dönemi; tür", "İki dönem arasında ödenen toplam anaparayı verir.", 6, 6, (a, c) => {
    const [r, n, pvv, s, e, t] = a.map((x) => num(x, c));
    if (r <= 0 || n <= 0 || pvv <= 0 || s < 1 || e < s || e > n) throw Errors.NUM;
    const p0 = pmt(r, n, pvv, 0, t ? 1 : 0);
    let sum = 0;
    for (let p = Math.ceil(s); p <= Math.floor(e); p++) sum += p0 - ipmt(r, p, n, pvv, 0, t ? 1 : 0);
    return sum;
  }),
  fn("XNPV", "ANBD", "financial", "oran; değerler; tarihler", "Dönemsel olmayan nakit akışlarının net bugünkü değerini verir.", 3, 3, (a, c) => {
    const r = num(a[0], c);
    const vs = cashflows(a[1], c);
    const ds = cashflows(a[2], c);
    if (vs.length !== ds.length) throw Errors.NUM;
    return vs.reduce((s, v, i) => s + v / Math.pow(1 + r, (ds[i] - ds[0]) / 365), 0);
  }),
  fn("XIRR", "AİÇVERİMORANI", "financial", "değerler; tarihler; [tahmin]", "Dönemsel olmayan nakit akışlarının iç verim oranını verir.", 2, 3, (a, c) => {
    const vs = cashflows(a[0], c);
    const ds = cashflows(a[1], c);
    if (vs.length !== ds.length) throw Errors.NUM;
    const f = (r: number) => vs.reduce((s, v, i) => s + v / Math.pow(1 + r, (ds[i] - ds[0]) / 365), 0);
    return newton(f, optNum(a[2], 0.1, c));
  }),
  fn("MIRR", "D_İÇ_VERİM_ORANI", "financial", "değerler; finansman_faiz_oranı; yeniden_yatırım_oranı", "Değiştirilmiş iç verim oranını verir.", 3, 3, (a, c) => {
    const vs = cashflows(a[0], c);
    const fr = num(a[1], c);
    const rr = num(a[2], c);
    const n = vs.length;
    let npvPos = 0;
    let npvNeg = 0;
    vs.forEach((v, i) => {
      if (v > 0) npvPos += v / Math.pow(1 + rr, i);
      else npvNeg += v / Math.pow(1 + fr, i);
    });
    if (npvNeg === 0 || npvPos === 0) throw Errors.DIV0;
    return Math.pow((-npvPos * Math.pow(1 + rr, n)) / (npvNeg * (1 + rr)), 1 / (n - 1)) - 1;
  }),
  fn("SLN", "DA", "financial", "maliyet; hurda; ömür", "Bir malın bir dönem için doğrusal amortismanını verir.", 3, 3, (a, c) => {
    const life = num(a[2], c);
    if (life === 0) throw Errors.DIV0;
    return (num(a[0], c) - num(a[1], c)) / life;
  }, { lift: true }),
  fn("SYD", "YAT", "financial", "maliyet; hurda; ömür; dönem", "Yılların toplamı yöntemiyle amortismanı verir.", 4, 4, (a, c) => {
    const [cost, salvage, life, per] = a.map((x) => num(x, c));
    if (life <= 0 || per <= 0 || per > life) throw Errors.NUM;
    return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
  }, { lift: true }),
  fn("DDB", "ÇİFTAZALANBAKİYE", "financial", "maliyet; hurda; ömür; dönem; [faktör]", "Çift azalan bakiye yöntemiyle amortismanı verir.", 4, 5, (a, c) => ddb(num(a[0], c), num(a[1], c), num(a[2], c), num(a[3], c), optNum(a[4], 2, c)), { lift: true }),
  fn("DB", "AZALANBAKİYE", "financial", "maliyet; hurda; ömür; dönem; [ay]", "Sabit azalan bakiye yöntemiyle amortismanı verir.", 4, 5, (a, c) => db(num(a[0], c), num(a[1], c), num(a[2], c), num(a[3], c), optNum(a[4], 12, c)), { lift: true }),
  fn("EFFECT", "ETKİN", "financial", "nominal_oran; dönem_sayısı_yıl", "Yıllık efektif faiz oranını verir.", 2, 2, (a, c) => {
    const r = num(a[0], c);
    const n = Math.trunc(num(a[1], c));
    if (r <= 0 || n < 1) throw Errors.NUM;
    return Math.pow(1 + r / n, n) - 1;
  }, { lift: true }),
  fn("NOMINAL", "NOMİNAL", "financial", "efektif_oran; dönem_sayısı_yıl", "Yıllık nominal faiz oranını verir.", 2, 2, (a, c) => {
    const r = num(a[0], c);
    const n = Math.trunc(num(a[1], c));
    if (r <= 0 || n < 1) throw Errors.NUM;
    return n * (Math.pow(1 + r, 1 / n) - 1);
  }, { lift: true }),
];
