/** Metin fonksiyonları (büyük/küçük harf dönüşümleri tr-TR yerel ayarıyla). */
import { parseNumberText, wildcardToRegExp, hasWildcards } from "../coerce";
import type { FnContext } from "../context";
import { formatFromLocal, formatValue, looksLocalFormat } from "../numfmt";
import { Errors, isError, makeArray, type Scalar, type Value } from "../values";
import { roundHalfAway } from "./math";
import { bool, forEachValue, int, matrix, num, optBool, optNum, optText, scalar, text } from "./helpers";
import { fn, type FnDef } from "./registry";

/** Windows-1254 (Türkçe) kod sayfası 128–255 arası karakterleri. */
const CP1254_HIGH =
  "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008d\u008e\u008f\u0090‘’“”•–—˜™š›œ\u009d\u009eŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ";

function charOf(n: number): string {
  if (n < 1 || n > 255) throw Errors.VALUE;
  return n < 128 ? String.fromCharCode(n) : CP1254_HIGH[n - 128];
}

function codeOf(s: string): number {
  if (s === "") throw Errors.VALUE;
  const ch = s[0];
  const c = ch.charCodeAt(0);
  if (c < 128) return c;
  const i = CP1254_HIGH.indexOf(ch);
  return i >= 0 ? 128 + i : 63;
}

export function properTr(s: string): string {
  let out = "";
  let prevLetter = false;
  for (const ch of s) {
    const isLetter = /\p{L}/u.test(ch);
    out += isLetter ? (prevLetter ? ch.toLocaleLowerCase("tr-TR") : ch.toLocaleUpperCase("tr-TR")) : ch;
    prevLetter = isLetter;
  }
  return out;
}

function textFormat(v: Scalar, fmt: string, c: FnContext): string {
  const code = looksLocalFormat(fmt) ? formatFromLocal(fmt) : fmt;
  let x: Scalar = v;
  if (typeof v === "string") {
    const n = parseNumberText(v, c.locale);
    if (n !== null) x = n;
  }
  if (typeof x === "boolean") return x ? "DOĞRU" : "YANLIŞ";
  if (x === null) x = 0;
  return formatValue(x, code, { lang: c.locale.lang, decimal: c.locale.decimal, thousands: c.locale.decimal === "," ? "." : "," }).text;
}

function fixed(n: number, decimals: number, noCommas: boolean, c: FnContext): string {
  const r = roundHalfAway(n, decimals);
  const d = Math.max(0, decimals);
  const code = `${noCommas ? "0" : "#,##0"}${d ? "." + "0".repeat(d) : ""}`;
  return formatValue(r, code, { decimal: c.locale.decimal, thousands: c.locale.decimal === "," ? "." : "," }).text;
}

function findIndex(hay: string, needle: string, start: number, ci: boolean): number {
  if (start < 1 || start > hay.length + 1) throw Errors.VALUE;
  if (!ci) {
    const i = hay.indexOf(needle, start - 1);
    if (i < 0) throw Errors.VALUE;
    return i + 1;
  }
  if (needle === "") return start;
  const sub = hay.slice(start - 1);
  if (hasWildcards(needle) || /~/.test(needle)) {
    const re = wildcardToRegExp(needle, false);
    const m = re.exec(sub);
    if (!m) throw Errors.VALUE;
    return start + m.index;
  }
  const i = sub.toLocaleLowerCase("tr-TR").indexOf(needle.toLocaleLowerCase("tr-TR"));
  if (i < 0) {
    const j = sub.toLowerCase().indexOf(needle.toLowerCase());
    if (j < 0) throw Errors.VALUE;
    return start + j;
  }
  return start + i;
}

function joinTexts(args: Value[], c: FnContext): string[] {
  const out: string[] = [];
  for (const a of args) {
    forEachValue(a, c, (s) => {
      if (isError(s)) throw s;
      out.push(s === null ? "" : typeof s === "string" ? s : text(s, c));
    }, false);
  }
  return out;
}

function splitBy(s: string, delims: string[], ci: boolean): string[] {
  const ds = delims.filter((d) => d !== "");
  if (ds.length === 0) return [s];
  const re = new RegExp(ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"), ci ? "giu" : "gu");
  return s.split(re);
}

function delimList(v: Value, c: FnContext): string[] {
  return matrix(v, c).flat().map((x) => (x === null ? "" : typeof x === "string" ? x : text(x, c)));
}

function textBeforeAfter(a: Value[], c: FnContext, before: boolean): Value {
  const s = text(a[0], c);
  const delims = delimList(a[1], c);
  let n = Math.trunc(optNum(a[2], 1, c));
  const ci = optNum(a[3], 0, c) === 1;
  const matchEnd = optNum(a[4], 0, c) === 1;
  if (n === 0 || Math.abs(n) > s.length + 1) throw Errors.VALUE;
  const hay = ci ? s.toLocaleLowerCase("tr-TR") : s;
  const ds = delims.map((d) => (ci ? d.toLocaleLowerCase("tr-TR") : d));
  // Tüm eşleşme konumları
  const hits: [number, number][] = [];
  for (let i = 0; i <= hay.length; i++) {
    for (const d of ds) {
      if (d === "" ? true : hay.startsWith(d, i)) {
        hits.push([i, d.length]);
        if (d !== "") i += d.length - 1;
        break;
      }
    }
    if (ds.every((d) => d === "")) break;
  }
  if (matchEnd) {
    if (n > 0) hits.push([hay.length, 0]);
    else hits.unshift([0, 0]);
  }
  let hit: [number, number] | undefined;
  if (n > 0) hit = hits[n - 1];
  else {
    n = -n;
    hit = hits[hits.length - n];
  }
  if (!hit) {
    if (a.length > 5) return a[5];
    throw Errors.NA;
  }
  return before ? s.slice(0, hit[0]) : s.slice(hit[0] + hit[1]);
}

export const textFunctions: FnDef[] = [
  fn("CONCAT", "ARALIKBİRLEŞTİR", "text", "metin1; [metin2]; ...", "Metin dizelerini ve aralıkları birleştirir.", 1, 255, (a, c) => {
    const s = joinTexts(a, c).join("");
    if (s.length > 32767) throw Errors.CALC;
    return s;
  }, { xlfn: true }),
  fn("CONCATENATE", "BİRLEŞTİR", "text", "metin1; [metin2]; ...", "Birden çok metin dizesini tek bir metin dizesinde birleştirir.", 1, 255, (a, c) => a.map((x) => text(x, c)).join(""), { lift: true }),
  fn("TEXTJOIN", "METİNBİRLEŞTİR", "text", "sınırlayıcı; boşları_yoksay; metin1; ...", "Metinleri sınırlayıcıyla birleştirir.", 3, 255, (a, c) => {
    const delims = delimList(a[0], c);
    const ignore = bool(a[1], c);
    let parts = joinTexts(a.slice(2), c);
    if (ignore) parts = parts.filter((p) => p !== "");
    let out = "";
    parts.forEach((p, i) => {
      if (i > 0) out += delims[(i - 1) % delims.length] ?? "";
      out += p;
    });
    return out;
  }, { xlfn: true }),
  fn("LEFT", "SOLDAN", "text", "metin; [sayı_karakterler]", "Metnin başındaki karakterleri verir.", 1, 2, (a, c) => {
    const n = Math.trunc(optNum(a[1], 1, c));
    if (n < 0) throw Errors.VALUE;
    return text(a[0], c).slice(0, n);
  }, { lift: true }),
  fn("RIGHT", "SAĞDAN", "text", "metin; [sayı_karakterler]", "Metnin sonundaki karakterleri verir.", 1, 2, (a, c) => {
    const n = Math.trunc(optNum(a[1], 1, c));
    if (n < 0) throw Errors.VALUE;
    const s = text(a[0], c);
    return n === 0 ? "" : s.slice(-n);
  }, { lift: true }),
  fn("MID", "PARÇAAL", "text", "metin; başlangıç_sayısı; sayı_karakterler", "Metnin ortasından belirtilen sayıda karakter verir.", 3, 3, (a, c) => {
    const st = Math.trunc(num(a[1], c));
    const n = Math.trunc(num(a[2], c));
    if (st < 1 || n < 0) throw Errors.VALUE;
    return text(a[0], c).substr(st - 1, n);
  }, { lift: true }),
  fn("LEN", "UZUNLUK", "text", "metin", "Metindeki karakter sayısını verir.", 1, 1, (a, c) => text(a[0], c).length, { lift: true }),
  fn("UPPER", "BÜYÜKHARF", "text", "metin", "Metni büyük harfe çevirir.", 1, 1, (a, c) => text(a[0], c).toLocaleUpperCase(c.locale.lang === "tr" ? "tr-TR" : "en-US"), { lift: true }),
  fn("LOWER", "KÜÇÜKHARF", "text", "metin", "Metni küçük harfe çevirir.", 1, 1, (a, c) => text(a[0], c).toLocaleLowerCase(c.locale.lang === "tr" ? "tr-TR" : "en-US"), { lift: true }),
  fn("PROPER", "YAZIM.DÜZENİ", "text", "metin", "Her sözcüğün ilk harfini büyük harfe çevirir.", 1, 1, (a, c) => properTr(text(a[0], c)), { lift: true }),
  fn("TRIM", "KIRP", "text", "metin", "Sözcükler arasındaki tek boşluklar dışındaki boşlukları kaldırır.", 1, 1, (a, c) => text(a[0], c).replace(/ +/g, " ").replace(/^ | $/g, ""), { lift: true }),
  fn("CLEAN", "TEMİZ", "text", "metin", "Yazdırılamayan karakterleri kaldırır.", 1, 1, (a, c) => text(a[0], c).replace(/[\x00-\x1f]/g, ""), { lift: true }),
  fn("SUBSTITUTE", "YERİNEKOY", "text", "metin; eski_metin; yeni_metin; [yineleme_sayısı]", "Metindeki eski metni yeni metinle değiştirir.", 3, 4, (a, c) => {
    const s = text(a[0], c);
    const o = text(a[1], c);
    const n = text(a[2], c);
    if (o === "") return s;
    if (a.length > 3 && a[3] !== null) {
      const k = Math.trunc(num(a[3], c));
      if (k < 1) throw Errors.VALUE;
      let idx = -1;
      for (let i = 0; i < k; i++) {
        idx = s.indexOf(o, idx + 1);
        if (idx < 0) return s;
      }
      return s.slice(0, idx) + n + s.slice(idx + o.length);
    }
    return s.split(o).join(n);
  }, { lift: true }),
  fn("REPLACE", "DEĞİŞTİR", "text", "eski_metin; başlangıç_sayısı; sayı_karakterler; yeni_metin", "Metnin bir kısmını başka bir metinle değiştirir.", 4, 4, (a, c) => {
    const s = text(a[0], c);
    const st = Math.trunc(num(a[1], c));
    const n = Math.trunc(num(a[2], c));
    if (st < 1 || n < 0) throw Errors.VALUE;
    return s.slice(0, st - 1) + text(a[3], c) + s.slice(st - 1 + n);
  }, { lift: true }),
  fn("FIND", "BUL", "text", "bul_metin; metin_içinde; [başlangıç_sayısı]", "Bir metni başka bir metin içinde bulur (büyük/küçük harf duyarlı).", 2, 3, (a, c) => findIndex(text(a[1], c), text(a[0], c), Math.trunc(optNum(a[2], 1, c)), false), { lift: true }),
  fn("SEARCH", "MBUL", "text", "bul_metin; metin; [başlangıç_sayısı]", "Bir metni başka bir metin içinde bulur (büyük/küçük harf duyarsız, joker destekli).", 2, 3, (a, c) => findIndex(text(a[1], c), text(a[0], c), Math.trunc(optNum(a[2], 1, c)), true), { lift: true }),
  fn("TEXT", "METNEÇEVİR", "text", "değer; biçim_metni", "Değeri belirli bir sayı biçiminde metne çevirir.", 2, 2, (a, c) => {
    const v = scalar(a[0], c);
    if (isError(v)) throw v;
    return textFormat(v, text(a[1], c), c);
  }, { lift: true }),
  fn("VALUE", "SAYIYAÇEVİR", "text", "metin", "Metni sayıya çevirir.", 1, 1, (a, c) => {
    const v = a[0] as Scalar;
    if (typeof v === "number") return v;
    if (v === null) return 0;
    if (typeof v === "boolean") throw Errors.VALUE;
    const n = parseNumberText(text(v, c), c.locale);
    if (n === null) throw Errors.VALUE;
    return n;
  }, { lift: true }),
  fn("NUMBERVALUE", "SAYIDEĞERİ", "text", "metin; [ondalık_ayırıcı]; [grup_ayırıcı]", "Metni yerel ayardan bağımsız olarak sayıya çevirir.", 1, 3, (a, c) => {
    let s = text(a[0], c).replace(/\s/g, "");
    const dec = optText(a[1], c.locale.decimal, c)[0] ?? ".";
    // Grup ayırıcısı verilmezse yerel ayarınki; ondalık ayırıcısıyla çakışırsa diğeri
    // (Türkçe ayarda SAYIDEĞERİ("1.5";".") 1,5 olur, 15 değil)
    const localGrp = c.locale.decimal === "," ? "." : ",";
    const grp = optText(a[2], localGrp === dec ? (dec === "." ? "," : ".") : localGrp, c)[0] ?? ",";
    if (s === "") return 0;
    let pct = 0;
    while (s.endsWith("%")) {
      pct++;
      s = s.slice(0, -1);
    }
    s = s.split(grp).join("").replace(dec, ".");
    if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) throw Errors.VALUE;
    return parseFloat(s) / Math.pow(100, pct);
  }, { lift: true, xlfn: true }),
  fn("REPT", "YİNELE", "text", "metin; sayı_kere", "Metni belirtilen sayıda yineler.", 2, 2, (a, c) => {
    const n = Math.trunc(num(a[1], c));
    if (n < 0) throw Errors.VALUE;
    const s = text(a[0], c);
    if (s.length * n > 32767) throw Errors.VALUE;
    return s.repeat(n);
  }, { lift: true }),
  fn("CHAR", "DAMGA", "text", "sayı", "Kod numarasıyla belirtilen karakteri verir.", 1, 1, (a, c) => charOf(Math.trunc(num(a[0], c))), { lift: true }),
  fn("CODE", "KOD", "text", "metin", "Metindeki ilk karakterin sayısal kodunu verir.", 1, 1, (a, c) => codeOf(text(a[0], c)), { lift: true }),
  fn("UNICHAR", "UNICODEKARAKTERİ", "text", "sayı", "Unicode karakterini verir.", 1, 1, (a, c) => {
    const n = Math.trunc(num(a[0], c));
    if (n < 1 || n > 0x10ffff) throw Errors.VALUE;
    return String.fromCodePoint(n);
  }, { lift: true, xlfn: true }),
  fn("UNICODE", "UNICODE", "text", "metin", "İlk karakterin Unicode kod noktasını verir.", 1, 1, (a, c) => {
    const s = text(a[0], c);
    if (s === "") throw Errors.VALUE;
    return s.codePointAt(0)!;
  }, { lift: true, xlfn: true }),
  fn("EXACT", "ÖZDEŞ", "text", "metin1; metin2", "İki metnin tam olarak aynı olup olmadığını denetler.", 2, 2, (a, c) => text(a[0], c) === text(a[1], c), { lift: true }),
  fn("FIXED", "SAYIDÜZENLE", "text", "sayı; [onluklar]; [virgül_yok]", "Sayıyı sabit ondalıklı metin olarak biçimlendirir.", 1, 3, (a, c) => fixed(num(a[0], c), Math.trunc(optNum(a[1], 2, c)), optBool(a[2], false, c), c), { lift: true }),
  fn("DOLLAR", "LİRA", "text", "sayı; [onluklar]", "Sayıyı para birimi biçiminde metne çevirir.", 1, 2, (a, c) => {
    const d = Math.trunc(optNum(a[1], 2, c));
    const n = roundHalfAway(num(a[0], c), d);
    const dd = Math.max(0, d);
    const code = `"₺"#,##0${dd ? "." + "0".repeat(dd) : ""};-"₺"#,##0${dd ? "." + "0".repeat(dd) : ""}`;
    return formatValue(n, code, { decimal: c.locale.decimal, thousands: c.locale.decimal === "," ? "." : "," }).text;
  }, { lift: true }),
  fn("T", "M", "text", "değer", "Değer metinse metni, değilse boş metin verir.", 1, 1, (a) => (typeof a[0] === "string" ? a[0] : isError(a[0]) ? a[0] : ""), { lift: true }),
  fn("TEXTBEFORE", "ÖNCEKİMETİN", "text", "metin; sınırlayıcı; [örnek_num]; [eşleşme_modu]; [eşleşme_sonu]; [bulunamazsa]", "Sınırlayıcıdan önceki metni verir.", 2, 6, (a, c) => textBeforeAfter(a, c, true), { xlfn: true }),
  fn("TEXTAFTER", "SONRAKİMETİN", "text", "metin; sınırlayıcı; [örnek_num]; [eşleşme_modu]; [eşleşme_sonu]; [bulunamazsa]", "Sınırlayıcıdan sonraki metni verir.", 2, 6, (a, c) => textBeforeAfter(a, c, false), { xlfn: true }),
  fn("TEXTSPLIT", "METİNBÖL", "text", "metin; sütun_sınırlayıcı; [satır_sınırlayıcı]; [boşları_yoksay]; [eşleşme_modu]; [doldurma_değeri]", "Metni sınırlayıcılarla satır ve sütunlara böler.", 2, 6, (a, c) => {
    const s = text(a[0], c);
    const colD = a[1] === null ? [] : delimList(a[1], c);
    const rowD = a.length > 2 && a[2] !== null ? delimList(a[2], c) : [];
    const ignore = optBool(a[3], false, c);
    const ci = optNum(a[4], 0, c) === 1;
    const pad: Scalar = a.length > 5 ? (a[5] as Scalar) : Errors.NA;
    let rows = rowD.length ? splitBy(s, rowD, ci) : [s];
    if (ignore) rows = rows.filter((r) => r !== "");
    const data = rows.map((r) => {
      let cells = colD.length ? splitBy(r, colD, ci) : [r];
      if (ignore) cells = cells.filter((x) => x !== "");
      return cells as Scalar[];
    });
    const w = Math.max(...data.map((r) => r.length));
    return makeArray(data.map((r) => [...r, ...new Array(w - r.length).fill(pad)]));
  }, { xlfn: true }),
];

void int;
