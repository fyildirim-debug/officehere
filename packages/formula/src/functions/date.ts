/** Tarih ve saat fonksiyonları (1900 tarih sistemi). */
import type { FnContext } from "../context";
import { dateToSerial, daysInMonth, isLeapYear, parseDateTime, serialToDate, weekdayOf } from "../dates";
import { Errors, type Value } from "../values";
import { collectNumbers, num, optNum, scalar, text } from "./helpers";
import { fn, type FnDef } from "./registry";

/** Gün içindeki saniye (Excel gibi en yakın saniyeye yuvarlanır). */
function wholeSeconds(serial: number): number {
  const secs = Math.round((serial - Math.floor(serial)) * 86400);
  return secs >= 86400 ? 0 : secs;
}

function serialArg(v: Value, c: FnContext): number {
  const n = num(v, c);
  if (n < 0 || n > 2958465.99999) throw Errors.NUM;
  return n;
}

function dateArg(v: Value, c: FnContext): number {
  return Math.floor(serialArg(v, c));
}

function makeDate(y: number, m: number, d: number): number {
  y = Math.trunc(y);
  m = Math.trunc(m);
  d = Math.trunc(d);
  if (y < 0 || y > 9999) throw Errors.NUM;
  if (y < 1900) y += 1900;
  // 29.02.1900 özel durumu
  if (y === 1900 && m === 2 && d === 29) return 60;
  const s = dateToSerial(y, m, d);
  if (s < 0 || s > 2958465) throw Errors.NUM;
  return s;
}

function addMonths(serial: number, months: number, endOfMonth: boolean): number {
  const p = serialToDate(serial);
  let m = p.m - 1 + Math.trunc(months);
  const y = p.y + Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  const dim = daysInMonth(y, m + 1);
  const d = endOfMonth ? dim : Math.min(p.d, dim);
  const s = dateToSerial(y, m + 1, d);
  if (s < 0) throw Errors.NUM;
  return s;
}

/** Hafta sonu maskesi: dizin 0 = Pazartesi … 6 = Pazar. */
function weekendMask(v: Value | undefined, c: FnContext): boolean[] {
  if (v === undefined || v === null) return [false, false, false, false, false, true, true];
  const s = scalar(v, c);
  if (typeof s === "string") {
    if (!/^[01]{7}$/.test(s) || s === "1111111") throw Errors.VALUE;
    return s.split("").map((ch) => ch === "1");
  }
  const n = Math.trunc(num(v, c));
  const mask = [false, false, false, false, false, false, false];
  if (n >= 1 && n <= 7) {
    // 1: Cmt-Paz, 2: Paz-Pzt, … 7: Cuma-Cmt
    const first = (n + 4) % 7; // 1 → 5 (Cmt)
    mask[first] = true;
    mask[(first + 1) % 7] = true;
  } else if (n >= 11 && n <= 17) mask[(n - 11 + 6) % 7] = true;
  else throw Errors.NUM;
  return mask;
}

/** Pazartesi=0 … Pazar=6. */
function monIdx(serial: number): number {
  return (weekdayOf(serial) + 6) % 7;
}

function holidaysOf(v: Value | undefined, c: FnContext): Set<number> {
  if (v === undefined || v === null) return new Set();
  return new Set(collectNumbers([v], c).map(Math.floor));
}

function networkDays(start: number, end: number, mask: boolean[], hol: Set<number>): number {
  const sign = start <= end ? 1 : -1;
  const [a, b] = start <= end ? [start, end] : [end, start];
  let n = 0;
  for (let d = a; d <= b; d++) if (!mask[monIdx(d)] && !hol.has(d)) n++;
  return n * sign;
}

function workday(start: number, days: number, mask: boolean[], hol: Set<number>): number {
  if (mask.every(Boolean)) throw Errors.VALUE;
  let d = start;
  const step = days >= 0 ? 1 : -1;
  let left = Math.abs(Math.trunc(days));
  while (left > 0) {
    d += step;
    if (!mask[monIdx(d)] && !hol.has(d)) left--;
  }
  return d;
}

function weekNum(serial: number, type: number): number {
  const p = serialToDate(serial);
  if (type === 21) return isoWeek(serial);
  // Haftanın başlangıç günü (0 = Pazar)
  const startMap: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };
  const sd = startMap[type];
  if (sd === undefined) throw Errors.NUM;
  const jan1 = dateToSerial(p.y, 1, 1);
  const offset = (weekdayOf(jan1) - sd + 7) % 7;
  return Math.floor((Math.floor(serial) - jan1 + offset) / 7) + 1;
}

function isoWeek(serial: number): number {
  const s = Math.floor(serial);
  const wd = monIdx(s); // 0 = Pazartesi
  const thursday = s - wd + 3;
  const y = serialToDate(thursday).y;
  const jan1 = dateToSerial(y, 1, 1);
  return Math.floor((thursday - jan1) / 7) + 1;
}

function datedif(start: number, end: number, unit: string): number {
  if (start > end) throw Errors.NUM;
  const a = serialToDate(start);
  const b = serialToDate(end);
  const u = unit.toUpperCase();
  let months = (b.y - a.y) * 12 + (b.m - a.m);
  if (b.d < a.d) months--;
  switch (u) {
    case "Y":
      return Math.floor(months / 12);
    case "M":
      return months;
    case "D":
      return Math.floor(end) - Math.floor(start);
    case "MD": {
      if (b.d >= a.d) return b.d - a.d;
      const pm = b.m === 1 ? 12 : b.m - 1;
      const py = b.m === 1 ? b.y - 1 : b.y;
      return daysInMonth(py, pm) - a.d + b.d;
    }
    case "YM":
      return ((months % 12) + 12) % 12;
    case "YD": {
      let s = dateToSerial(b.y, a.m, Math.min(a.d, daysInMonth(b.y, a.m)));
      if (s > end) s = dateToSerial(b.y - 1, a.m, Math.min(a.d, daysInMonth(b.y - 1, a.m)));
      return Math.floor(end) - s;
    }
    default:
      throw Errors.NUM;
  }
}

function days360(start: number, end: number, european: boolean): number {
  const a = serialToDate(start);
  const b = serialToDate(end);
  let d1 = a.d;
  let d2 = b.d;
  let m2 = b.m;
  if (d1 === 31) d1 = 30;
  else if (!european && a.m === 2 && (d1 === 29 || (d1 === 28 && !isLeapYear(a.y)))) d1 = 30;
  if (d2 === 31) {
    if (!european && d1 !== 30) {
      d2 = 1;
      m2++;
    } else d2 = 30;
  }
  return (b.y - a.y) * 360 + (m2 - a.m) * 30 + (d2 - d1);
}

function yearFrac(s1: number, s2: number, basis: number): number {
  let start = Math.floor(s1);
  let end = Math.floor(s2);
  if (start > end) [start, end] = [end, start];
  const a = serialToDate(start);
  const b = serialToDate(end);
  switch (basis) {
    case 0:
      return days360(start, end, false) / 360;
    case 1: {
      if (a.y === b.y) return (end - start) / (isLeapYear(a.y) ? 366 : 365);
      const oneYear = b.y === a.y + 1 && (b.m < a.m || (b.m === a.m && b.d <= a.d));
      if (oneYear) {
        const leap = (isLeapYear(a.y) && (a.m < 3)) || (isLeapYear(b.y) && (b.m > 2 || (b.m === 2 && b.d === 29)));
        return (end - start) / (leap ? 366 : 365);
      }
      let total = 0;
      for (let y = a.y; y <= b.y; y++) total += isLeapYear(y) ? 366 : 365;
      return (end - start) / (total / (b.y - a.y + 1));
    }
    case 2:
      return (end - start) / 360;
    case 3:
      return (end - start) / 365;
    case 4:
      return days360(start, end, true) / 360;
    default:
      throw Errors.NUM;
  }
}

export const dateFunctions: FnDef[] = [
  fn("TODAY", "BUGÜN", "date", "", "Bugünün tarihini verir.", 0, 0, (_a, c) => Math.floor(c.host.now()), { volatile: true }),
  fn("NOW", "ŞİMDİ", "date", "", "Geçerli tarih ve saati verir.", 0, 0, (_a, c) => c.host.now(), { volatile: true }),
  fn("DATE", "TARİH", "date", "yıl; ay; gün", "Belirli bir tarihin seri numarasını verir.", 3, 3, (a, c) => makeDate(num(a[0], c), num(a[1], c), num(a[2], c)), { lift: true }),
  fn("TIME", "ZAMAN", "date", "saat; dakika; saniye", "Belirli bir saatin seri numarasını verir.", 3, 3, (a, c) => {
    const t = (Math.trunc(num(a[0], c)) * 3600 + Math.trunc(num(a[1], c)) * 60 + Math.trunc(num(a[2], c))) / 86400;
    if (t < 0) throw Errors.NUM;
    return t - Math.floor(t);
  }, { lift: true }),
  fn("YEAR", "YIL", "date", "seri_no", "Tarihin yılını verir.", 1, 1, (a, c) => serialToDate(serialArg(a[0], c)).y, { lift: true }),
  fn("MONTH", "AY", "date", "seri_no", "Tarihin ayını (1–12) verir.", 1, 1, (a, c) => serialToDate(serialArg(a[0], c)).m, { lift: true }),
  fn("DAY", "GÜN", "date", "seri_no", "Tarihin gününü (1–31) verir.", 1, 1, (a, c) => serialToDate(serialArg(a[0], c)).d, { lift: true }),
  fn("HOUR", "SAAT", "date", "seri_no", "Saati (0–23) verir.", 1, 1, (a, c) => Math.floor(wholeSeconds(serialArg(a[0], c)) / 3600), { lift: true }),
  fn("MINUTE", "DAKİKA", "date", "seri_no", "Dakikayı (0–59) verir.", 1, 1, (a, c) => Math.floor(wholeSeconds(serialArg(a[0], c)) / 60) % 60, { lift: true }),
  fn("SECOND", "SANİYE", "date", "seri_no", "Saniyeyi (0–59) verir.", 1, 1, (a, c) => wholeSeconds(serialArg(a[0], c)) % 60, { lift: true }),
  fn("WEEKDAY", "HAFTANINGÜNÜ", "date", "seri_no; [döndür_tür]", "Haftanın gününü gösteren sayıyı verir.", 1, 2, (a, c) => {
    const s = dateArg(a[0], c);
    const t = Math.trunc(optNum(a[1], 1, c));
    const w = weekdayOf(s); // 0 = Pazar
    if (t === 1) return w + 1;
    if (t === 2) return ((w + 6) % 7) + 1;
    if (t === 3) return (w + 6) % 7;
    if (t >= 11 && t <= 17) {
      const start = (t - 10) % 7; // 11 → Pazartesi(1)
      return ((w - start + 7) % 7) + 1;
    }
    throw Errors.NUM;
  }, { lift: true }),
  fn("WEEKNUM", "HAFTASAY", "date", "seri_no; [döndür_tür]", "Yılın kaçıncı haftası olduğunu verir.", 1, 2, (a, c) => weekNum(dateArg(a[0], c), Math.trunc(optNum(a[1], 1, c))), { lift: true }),
  fn("ISOWEEKNUM", "ISOHAFTASAY", "date", "tarih", "ISO hafta numarasını verir.", 1, 1, (a, c) => isoWeek(dateArg(a[0], c)), { lift: true, xlfn: true }),
  fn("EDATE", "SERİTARİH", "date", "başlangıç_tarihi; aylar", "Başlangıç tarihinden belirtilen ay sayısı önceki/sonraki tarihi verir.", 2, 2, (a, c) => addMonths(dateArg(a[0], c), num(a[1], c), false), { lift: true }),
  fn("EOMONTH", "SERİAY", "date", "başlangıç_tarihi; aylar", "Belirtilen ay sayısı önceki/sonraki ayın son gününü verir.", 2, 2, (a, c) => addMonths(dateArg(a[0], c), num(a[1], c), true), { lift: true }),
  fn("DATEDIF", "ETARİHLİ", "date", "başlangıç_tarihi; bitiş_tarihi; birim", "İki tarih arasındaki gün, ay ya da yıl sayısını verir.", 3, 3, (a, c) => datedif(dateArg(a[0], c), dateArg(a[1], c), text(a[2], c)), { lift: true }),
  fn("DAYS", "GÜNSAY", "date", "bitiş_tarihi; başlangıç_tarihi", "İki tarih arasındaki gün sayısını verir.", 2, 2, (a, c) => dateArg(a[0], c) - dateArg(a[1], c), { lift: true, xlfn: true }),
  fn("DAYS360", "GÜN360", "date", "başlangıç_tarihi; bitiş_tarihi; [yöntem]", "360 günlük yıla göre gün sayısını verir.", 2, 3, (a, c) => days360(dateArg(a[0], c), dateArg(a[1], c), !!optNum(a[2], 0, c)), { lift: true }),
  fn("NETWORKDAYS", "TAMİŞGÜNÜ", "date", "başlangıç_tarihi; bitiş_tarihi; [tatiller]", "İki tarih arasındaki tam iş günü sayısını verir.", 2, 3, (a, c) => networkDays(dateArg(a[0], c), dateArg(a[1], c), weekendMask(undefined, c), holidaysOf(a[2], c))),
  fn("NETWORKDAYS.INTL", "TAMİŞGÜNÜ.ULUSL", "date", "başlangıç_tarihi; bitiş_tarihi; [hafta_sonu]; [tatiller]", "Özel hafta sonu parametreleriyle tam iş günü sayısını verir.", 2, 4, (a, c) => networkDays(dateArg(a[0], c), dateArg(a[1], c), weekendMask(a[2], c), holidaysOf(a[3], c))),
  fn("WORKDAY", "İŞGÜNÜ", "date", "başlangıç_tarihi; gün_sayısı; [tatiller]", "Belirtilen iş günü sayısı önceki/sonraki tarihi verir.", 2, 3, (a, c) => workday(dateArg(a[0], c), num(a[1], c), weekendMask(undefined, c), holidaysOf(a[2], c))),
  fn("WORKDAY.INTL", "İŞGÜNÜ.ULUSL", "date", "başlangıç_tarihi; gün_sayısı; [hafta_sonu]; [tatiller]", "Özel hafta sonu parametreleriyle iş günü tarihini verir.", 2, 4, (a, c) => workday(dateArg(a[0], c), num(a[1], c), weekendMask(a[2], c), holidaysOf(a[3], c))),
  fn("DATEVALUE", "TARİHSAYISI", "date", "tarih_metni", "Metin biçimindeki tarihi seri numarasına çevirir.", 1, 1, (a, c) => {
    const p = parseDateTime(text(a[0], c));
    if (!p || !p.hasDate) throw Errors.VALUE;
    return Math.floor(p.serial);
  }, { lift: true }),
  fn("TIMEVALUE", "ZAMANSAYISI", "date", "saat_metni", "Metin biçimindeki saati seri numarasına çevirir.", 1, 1, (a, c) => {
    const p = parseDateTime(text(a[0], c));
    if (!p || !p.hasTime) throw Errors.VALUE;
    return p.serial - Math.floor(p.serial);
  }, { lift: true }),
  fn("YEARFRAC", "YILORAN", "date", "başlangıç_tarihi; bitiş_tarihi; [temel]", "İki tarih arasındaki tam gün sayısının yıl kesrini verir.", 2, 3, (a, c) => yearFrac(dateArg(a[0], c), dateArg(a[1], c), Math.trunc(optNum(a[2], 0, c))), { lift: true }),
];
