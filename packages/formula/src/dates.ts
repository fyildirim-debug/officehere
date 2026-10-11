/**
 * 1900 tarih sistemi seri numaraları. Excel ile uyumluluk için 1900 yılı
 * artık yıl kabul edilir: seri 60 = 29.02.1900 (aslında var olmayan gün).
 */

const MS_PER_DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

export const MONTHS_TR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
export const MONTHS_TR_SHORT = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];
export const DAYS_TR = ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"];
export const DAYS_TR_SHORT = ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"];
export const MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const DAYS_EN = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

export interface DateParts {
  y: number;
  m: number; // 1..12
  d: number;
}

/** Yıl/ay/gün → seri (taşan ay ve günler Excel DATE gibi düzeltilir). */
export function dateToSerial(y: number, m: number, d: number): number {
  const ms = Date.UTC(y, m - 1, 1) + (d - 1) * MS_PER_DAY;
  let days = Math.round((ms - EPOCH) / MS_PER_DAY);
  // 01.03.1900 öncesi tarihler bir gün geride sayılır (hayali 29.02.1900).
  if (days < 61) days -= 1;
  return days;
}

/** Seri → yıl/ay/gün (yalnız tam kısım). */
export function serialToDate(serial: number): DateParts {
  const s = Math.floor(serial);
  if (s === 60) return { y: 1900, m: 2, d: 29 };
  if (s === 0) return { y: 1900, m: 1, d: 0 };
  const days = s < 60 ? s + 1 : s;
  const dt = new Date(EPOCH + days * MS_PER_DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

/** Haftanın günü: 0 = Pazar. */
export function weekdayOf(serial: number): number {
  const s = Math.floor(serial);
  // Seri 1 (01.01.1900) Excel'e göre Pazar; 60'tan sonra gerçek takvimle hizalanır.
  if (s < 61) return (s + 6) % 7;
  return (s + 6) % 7;
}

export interface TimeParts {
  h: number;
  mi: number;
  s: number;
  ms: number;
}

/** Günün kesir kısmı → saat/dakika/saniye (Excel gibi en yakın milisaniyeye yuvarlanır). */
export function serialToTime(serial: number): TimeParts {
  let frac = serial - Math.floor(serial);
  let totalMs = Math.round(frac * MS_PER_DAY);
  if (totalMs >= MS_PER_DAY) totalMs = MS_PER_DAY - 1;
  frac = totalMs;
  const h = Math.floor(frac / 3600000);
  const mi = Math.floor((frac % 3600000) / 60000);
  const s = Math.floor((frac % 60000) / 1000);
  const ms = frac % 1000;
  return { h, mi, s, ms };
}

export function timeToSerial(h: number, mi: number, s: number): number {
  return (h * 3600 + mi * 60 + s) / 86400;
}

/** JS Date (yerel saat) → seri. */
export function jsDateToSerial(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
  let v = (utc - EPOCH) / MS_PER_DAY;
  if (v < 61) v -= 1;
  return v;
}

/** Seri → JS Date (UTC bileşenleriyle). */
export function serialToJsDate(serial: number): Date {
  const { y, m, d } = serialToDate(serial);
  const t = serialToTime(serial);
  return new Date(Date.UTC(y, m - 1, d, t.h, t.mi, t.s, t.ms));
}

export function isLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

export function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

const MONTH_LOOKUP: Record<string, number> = {};
[MONTHS_TR, MONTHS_TR_SHORT, MONTHS_EN, MONTHS_EN.map((m) => m.slice(0, 3))].forEach((list) =>
  list.forEach((name, i) => {
    MONTH_LOOKUP[name.toLocaleLowerCase("tr-TR")] = i + 1;
    MONTH_LOOKUP[name.toLowerCase()] = i + 1;
  }),
);

/**
 * Metinden tarih/saat ayrıştırır: "31.12.2024", "31/12/2024", "2024-12-31",
 * "31 Aralık 2024", "14:30", "31.12.2024 14:30:15". Başarısızsa null.
 */
export function parseDateTime(text: string): { serial: number; hasDate: boolean; hasTime: boolean; format: string } | null {
  const t = text.trim();
  if (!t) return null;
  let datePart = t;
  let timePart = "";
  const tm = /^(.*?)\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:[.,]\d+)?))?\s*(ÖÖ|ÖS|AM|PM)?$/i.exec(t);
  if (tm) {
    datePart = tm[1].trim();
    timePart = tm[0].slice(tm[1].length).trim();
  }
  let serial = 0;
  let hasDate = false;
  let format = "";
  if (datePart) {
    let y: number | undefined;
    let m: number | undefined;
    let d: number | undefined;
    let mm: RegExpExecArray | null;
    if ((mm = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/.exec(datePart))) {
      d = +mm[1];
      m = +mm[2];
      y = +mm[3];
      format = "dd.mm.yyyy";
    } else if ((mm = /^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/.exec(datePart))) {
      y = +mm[1];
      m = +mm[2];
      d = +mm[3];
      format = "yyyy-mm-dd";
    } else if ((mm = /^(\d{1,2})[\s.-]+([\p{L}]+)\.?[\s.-]+(\d{2,4})$/u.exec(datePart))) {
      d = +mm[1];
      m = MONTH_LOOKUP[mm[2].toLocaleLowerCase("tr-TR")];
      y = +mm[3];
      format = "d mmmm yyyy";
    } else if ((mm = /^(\d{1,2})[\s.-]+([\p{L}]+)$/u.exec(datePart))) {
      d = +mm[1];
      m = MONTH_LOOKUP[mm[2].toLocaleLowerCase("tr-TR")];
      y = new Date().getFullYear();
      format = "d mmm";
    } else if ((mm = /^(\d{1,2})[./](\d{1,2})$/.exec(datePart)) && timePart === "") {
      d = +mm[1];
      m = +mm[2];
      y = new Date().getFullYear();
      format = "d mmm";
    } else {
      return null;
    }
    if (m === undefined || !y || m < 1 || m > 12 || d < 1) return null;
    if (y < 100) y += y < 30 ? 2000 : 1900;
    if (d > daysInMonth(y, m) && !(y === 1900 && m === 2 && d === 29)) return null;
    serial = y === 1900 && m === 2 && d === 29 ? 60 : dateToSerial(y, m, d);
    hasDate = true;
  }
  let hasTime = false;
  if (tm) {
    let h = +tm[2];
    const mi = +tm[3];
    const s = tm[4] ? parseFloat(tm[4].replace(",", ".")) : 0;
    const ap = tm[5]?.toUpperCase();
    if (ap === "PM" || ap === "ÖS") {
      if (h < 12) h += 12;
    } else if ((ap === "AM" || ap === "ÖÖ") && h === 12) h = 0;
    if (mi > 59 || s >= 60) return null;
    serial += timeToSerial(h, mi, s);
    hasTime = true;
    const tf = tm[4] ? "h:mm:ss" : "h:mm";
    format = hasDate ? `${format} ${tf}` : tf;
  }
  if (!hasDate && !hasTime) return null;
  return { serial, hasDate, hasTime, format };
}
