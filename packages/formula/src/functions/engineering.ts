/**
 * Mühendislik işlevleri: taban dönüşümleri (ikili/sekizli/onlu/onaltılı, 10 basamaklı
 * ikiye tümleyen negatifler), bit işlemleri, DELTA/BESINIR ve birim dönüştürme (ÇEVİR).
 */
import type { FnContext } from "../context";
import { Errors, isError, type Value } from "../values";
import { num, scalar } from "./helpers";
import { fn, type FnDef } from "./registry";

/* ------------------------------ Taban dönüşümleri ------------------------------ */

interface Base {
  radix: number;
  /** 10 basamakta işaretli aralık: [-2^(n-1), 2^(n-1)-1] ve tümleyen için 2^n. */
  bits: number;
  re: RegExp;
}

const BIN: Base = { radix: 2, bits: 10, re: /^[01]{1,10}$/ };
const OCT: Base = { radix: 8, bits: 30, re: /^[0-7]{1,10}$/ };
const HEX: Base = { radix: 16, bits: 40, re: /^[0-9A-F]{1,10}$/ };
const DEC: Base = { radix: 10, bits: 0, re: /^$/ };

/** Kaynak tabandaki metni (en çok 10 karakter) işaretli tamsayıya çevirir. */
function parseBase(v: Value, b: Base, ctx: FnContext): number {
  const s = scalar(v, ctx);
  if (isError(s)) throw s;
  if (typeof s === "boolean") throw Errors.VALUE;
  let t: string;
  if (s === null) t = "0";
  else if (typeof s === "number") {
    if (s < 0 || !Number.isInteger(s)) throw Errors.NUM;
    t = String(s);
  } else t = s.trim().toUpperCase();
  if (t === "") t = "0";
  if (!b.re.test(t)) throw Errors.NUM;
  const n = parseInt(t, b.radix);
  const full = 2 ** b.bits;
  // 10 basamak dolu ve en üst bit 1 ise negatif (ikiye tümleyen)
  return t.length === 10 && n >= full / 2 ? n - full : n;
}

/** İşaretli tamsayıyı hedef tabana yazar; negatifler 10 basamaklı tümleyendir. */
function formatBase(n: number, b: Base, places: Value | undefined, ctx: FnContext): string {
  const half = 2 ** (b.bits - 1);
  if (n < -half || n > half - 1) throw Errors.NUM;
  if (n < 0) return (2 ** b.bits + n).toString(b.radix).toUpperCase();
  let s = n.toString(b.radix).toUpperCase();
  if (places !== undefined && places !== null) {
    const p = Math.trunc(num(places, ctx));
    if (p < 0 || p > 10 || p < s.length) throw Errors.NUM;
    s = s.padStart(p, "0");
  }
  return s;
}

function decimalArg(v: Value, ctx: FnContext): number {
  const s = scalar(v, ctx);
  if (isError(s)) throw s;
  if (typeof s === "boolean") throw Errors.VALUE;
  return Math.trunc(num(v, ctx));
}

function conv(name: string, tr: string, from: Base, to: Base, desc: string): FnDef {
  const toDec = to === DEC;
  return fn(name, tr, "engineering", toDec ? "sayı" : "sayı; [basamak]", desc, 1, toDec ? 1 : 2, (a, c) => {
    const n = from === DEC ? decimalArg(a[0], c) : parseBase(a[0], from, c);
    if (toDec) return n;
    return formatBase(n, to, a[1], c);
  }, { lift: true });
}

/* ------------------------------ Bit işlemleri ------------------------------ */

const BIT_LIMIT = 2 ** 48;

function bitArg(v: Value, ctx: FnContext): bigint {
  const s = scalar(v, ctx);
  if (isError(s)) throw s;
  if (typeof s === "boolean") throw Errors.VALUE;
  const n = num(v, ctx);
  if (n < 0 || n >= BIT_LIMIT || !Number.isInteger(n)) throw Errors.NUM;
  return BigInt(n);
}

function bitFn(name: string, tr: string, desc: string, op: (a: bigint, b: bigint) => bigint): FnDef {
  return fn(name, tr, "engineering", "sayı1; sayı2", desc, 2, 2, (a, c) => Number(op(bitArg(a[0], c), bitArg(a[1], c))), { lift: true, xlfn: true });
}

function shift(a: Value[], c: FnContext, left: boolean): number {
  const n = bitArg(a[0], c);
  const amt = Math.trunc(num(a[1], c));
  if (Math.abs(amt) > 53) throw Errors.NUM;
  const k = left ? amt : -amt;
  const r = k >= 0 ? n << BigInt(k) : n >> BigInt(-k);
  if (r >= BigInt(BIT_LIMIT)) throw Errors.NUM;
  return Number(r);
}

/* ------------------------------ ÇEVİR (CONVERT) ------------------------------ */

type Unit = { cat: string; f: number; prefix?: boolean; bin?: boolean };

/**
 * Birimler: kategori ve SI temel birimine çarpan. prefix: SI önekleri alabilir,
 * bin: ikili önekler (ki, Mi…) alabilir. Katsayılar uluslararası tanımlardan.
 */
const U: Record<string, Unit> = {};
function def(names: string[], cat: string, f: number, prefix = false, bin = false): void {
  for (const n of names) U[n] = { cat, f, prefix, bin };
}
// Kütle (kg temel değil: gram temel alınır)
def(["g"], "mass", 1, true);
def(["sg"], "mass", 14593.902937206364);
def(["lbm"], "mass", 453.59237);
def(["u"], "mass", 1.66053886e-24, true);
def(["ozm"], "mass", 28.349523125);
def(["grain"], "mass", 0.06479891);
def(["cwt", "shweight"], "mass", 45359.237);
def(["uk_cwt", "lcwt", "hweight"], "mass", 50802.34544);
def(["stone"], "mass", 6350.29318);
def(["ton"], "mass", 907184.74);
def(["uk_ton", "LTON", "brton"], "mass", 1016046.9088);
// Uzunluk (metre)
def(["m"], "dist", 1, true);
def(["mi"], "dist", 1609.344);
def(["Nmi"], "dist", 1852);
def(["in"], "dist", 0.0254);
def(["ft"], "dist", 0.3048);
def(["yd"], "dist", 0.9144);
def(["ang"], "dist", 1e-10, true);
def(["ell"], "dist", 1.143);
def(["ly"], "dist", 9460730472580800, true);
def(["parsec", "pc"], "dist", 30856775812815500, true);
def(["Picapt", "Pica"], "dist", 0.0254 / 72);
def(["pica"], "dist", 0.0254 / 6);
def(["survey_mi"], "dist", 1609.3472186944373);
// Zaman (saniye)
def(["yr"], "time", 31557600);
def(["day", "d"], "time", 86400);
def(["hr"], "time", 3600);
def(["mn", "min"], "time", 60);
def(["sec", "s"], "time", 1, true);
// Basınç (paskal)
def(["Pa", "p"], "press", 1, true);
def(["atm", "at"], "press", 101325, true);
def(["mmHg"], "press", 133.322, true);
def(["psi"], "press", 6894.757293168361);
def(["Torr"], "press", 101325 / 760);
// Kuvvet (newton)
def(["N"], "force", 1, true);
def(["dyn", "dy"], "force", 1e-5, true);
def(["lbf"], "force", 4.4482216152605);
def(["pond"], "force", 0.00980665, true);
// Enerji (joule)
def(["J"], "energy", 1, true);
def(["e"], "energy", 1e-7, true);
def(["c"], "energy", 4.184, true);
def(["cal"], "energy", 4.1868, true);
def(["eV", "ev"], "energy", 1.602176487e-19, true);
def(["HPh", "hh"], "energy", 2684519.537696173);
def(["Wh", "wh"], "energy", 3600, true);
def(["flb"], "energy", 0.04214011009380476);
def(["BTU", "btu"], "energy", 1055.05585262);
// Güç (watt)
def(["HP", "h"], "power", 745.69987158227);
def(["PS"], "power", 735.49875);
def(["W", "w"], "power", 1, true);
// Manyetizma (tesla)
def(["T"], "mag", 1, true);
def(["ga"], "mag", 1e-4, true);
// Hacim (metreküp)
def(["tsp"], "vol", 4.92892159375e-6);
def(["tspm"], "vol", 5e-6);
def(["tbs"], "vol", 1.478676478125e-5);
def(["oz"], "vol", 2.95735295625e-5);
def(["cup"], "vol", 2.365882365e-4);
def(["pt", "us_pt"], "vol", 4.73176473e-4);
def(["uk_pt"], "vol", 5.6826125e-4);
def(["qt"], "vol", 9.46352946e-4);
def(["uk_qt"], "vol", 1.1365225e-3);
def(["gal"], "vol", 3.785411784e-3);
def(["uk_gal"], "vol", 4.54609e-3);
def(["l", "L", "lt"], "vol", 1e-3, true);
def(["ang3", "ang^3"], "vol", 1e-30, true);
def(["barrel"], "vol", 0.158987294928);
def(["bushel"], "vol", 0.03523907016688);
def(["ft3", "ft^3"], "vol", 0.028316846592);
def(["in3", "in^3"], "vol", 1.6387064e-5);
def(["ly3", "ly^3"], "vol", 8.46786664623715e47);
def(["m3", "m^3"], "vol", 1, true);
def(["mi3", "mi^3"], "vol", 4168181825.440579584);
def(["yd3", "yd^3"], "vol", 0.764554857984);
def(["Nmi3", "Nmi^3"], "vol", 6352182208);
def(["Picapt3", "Picapt^3", "Pica3", "Pica^3"], "vol", (0.0254 / 72) ** 3);
def(["GRT", "regton"], "vol", 2.8316846592);
def(["MTON"], "vol", 1.13267386368);
// Alan (metrekare)
def(["uk_acre"], "area", 4046.8564224);
def(["us_acre"], "area", 4046.872609874252);
def(["ang2", "ang^2"], "area", 1e-20, true);
def(["ar"], "area", 100, true);
def(["ft2", "ft^2"], "area", 0.09290304);
def(["ha"], "area", 10000);
def(["in2", "in^2"], "area", 6.4516e-4);
def(["ly2", "ly^2"], "area", 8.95054210748189e31);
def(["m2", "m^2"], "area", 1, true);
def(["Morgen"], "area", 2500);
def(["mi2", "mi^2"], "area", 2589988.110336);
def(["Nmi2", "Nmi^2"], "area", 3429904);
def(["Picapt2", "Pica2", "Pica^2", "Picapt^2"], "area", (0.0254 / 72) ** 2);
def(["yd2", "yd^2"], "area", 0.83612736);
// Bilgi (bit)
def(["bit"], "info", 1, true, true);
def(["byte"], "info", 8, true, true);
// Hız (m/s)
def(["admkn"], "speed", 0.514773333);
def(["kn"], "speed", 0.5144444444444445);
def(["m/h", "m/hr"], "speed", 1 / 3600, true);
def(["m/s", "m/sec"], "speed", 1, true);
def(["mph"], "speed", 0.44704);
// Sıcaklık: ayrı ele alınır
const TEMP = new Set(["C", "cel", "F", "fah", "K", "kel", "Rank", "Reau"]);

const SI: Record<string, number> = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, e: 10, d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 };
const BIN_PREFIX: Record<string, number> = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };

function lookupUnit(name: string): { cat: string; f: number } | null {
  const direct = U[name];
  if (direct) return direct;
  for (const [p, mul] of Object.entries(BIN_PREFIX)) {
    if (!name.startsWith(p)) continue;
    const base = U[name.slice(p.length)];
    if (base?.bin) return { cat: base.cat, f: base.f * mul };
  }
  for (const [p, mul] of Object.entries(SI)) {
    if (!name.startsWith(p) || name.length === p.length) continue;
    const rest = name.slice(p.length);
    const base = U[rest];
    if (!base?.prefix) continue;
    // Alan/hacim birimlerinde önek uzunlukla birlikte üslenir (km2 = 1e6 m2)
    const pow = /(\^?2)$/.test(rest) && base.cat === "area" ? 2 : /(\^?3)$/.test(rest) && base.cat === "vol" ? 3 : 1;
    return { cat: base.cat, f: base.f * mul ** pow };
  }
  return null;
}

function toKelvin(v: number, u: string): number {
  switch (u) {
    case "C":
    case "cel":
      return v + 273.15;
    case "F":
    case "fah":
      return ((v - 32) * 5) / 9 + 273.15;
    case "K":
    case "kel":
      return v;
    case "Rank":
      return (v * 5) / 9;
    default:
      return v * 1.25 + 273.15; // Reau
  }
}

function fromKelvin(k: number, u: string): number {
  switch (u) {
    case "C":
    case "cel":
      return k - 273.15;
    case "F":
    case "fah":
      return ((k - 273.15) * 9) / 5 + 32;
    case "K":
    case "kel":
      return k;
    case "Rank":
      return (k * 9) / 5;
    default:
      return (k - 273.15) / 1.25;
  }
}

function unitText(v: Value, ctx: FnContext): string {
  const s = scalar(v, ctx);
  if (isError(s)) throw s;
  if (typeof s !== "string") throw Errors.NA;
  return s;
}

function convert(a: Value[], c: FnContext): number {
  const x = num(a[0], c);
  const from = unitText(a[1], c);
  const to = unitText(a[2], c);
  if (TEMP.has(from) || TEMP.has(to)) {
    if (!TEMP.has(from) || !TEMP.has(to)) throw Errors.NA;
    return clean(fromKelvin(toKelvin(x, from), to));
  }
  const uf = lookupUnit(from);
  const ut = lookupUnit(to);
  if (!uf || !ut || uf.cat !== ut.cat) throw Errors.NA;
  return clean((x * uf.f) / ut.f);
}

/** Kayan nokta gürültüsünü 15 anlamlı basamağa yuvarlar. */
function clean(n: number): number {
  return n === 0 ? 0 : Number(n.toPrecision(15));
}

export const engineeringFunctions: FnDef[] = [
  conv("BIN2DEC", "İKİONDA", BIN, DEC, "İkili sayıyı onluya dönüştürür."),
  conv("BIN2HEX", "İKİONALTI", BIN, HEX, "İkili sayıyı onaltılıya dönüştürür."),
  conv("BIN2OCT", "İKİSEKİZ", BIN, OCT, "İkili sayıyı sekizliye dönüştürür."),
  conv("DEC2BIN", "ONDAİKİ", DEC, BIN, "Onlu sayıyı ikiliye dönüştürür."),
  conv("DEC2HEX", "ONDAONALTI", DEC, HEX, "Onlu sayıyı onaltılıya dönüştürür."),
  conv("DEC2OCT", "ONDASEKİZ", DEC, OCT, "Onlu sayıyı sekizliye dönüştürür."),
  conv("HEX2BIN", "ONALTIİKİ", HEX, BIN, "Onaltılı sayıyı ikiliye dönüştürür."),
  conv("HEX2DEC", "ONALTIONDA", HEX, DEC, "Onaltılı sayıyı onluya dönüştürür."),
  conv("HEX2OCT", "ONALTISEKİZ", HEX, OCT, "Onaltılı sayıyı sekizliye dönüştürür."),
  conv("OCT2BIN", "SEKİZİKİ", OCT, BIN, "Sekizli sayıyı ikiliye dönüştürür."),
  conv("OCT2DEC", "SEKİZONDA", OCT, DEC, "Sekizli sayıyı onluya dönüştürür."),
  conv("OCT2HEX", "SEKİZONALTI", OCT, HEX, "Sekizli sayıyı onaltılıya dönüştürür."),
  bitFn("BITAND", "BİTVE", "İki sayının bit düzeyinde VE'sini verir.", (x, y) => x & y),
  bitFn("BITOR", "BİTVEYA", "İki sayının bit düzeyinde VEYA'sını verir.", (x, y) => x | y),
  bitFn("BITXOR", "BİTÖZELVEYA", "İki sayının bit düzeyinde ÖZEL VEYA'sını verir.", (x, y) => x ^ y),
  fn("BITLSHIFT", "BİTSOLAKAYDIR", "engineering", "sayı; kaydırma_miktarı", "Sayıyı belirtilen bit kadar sola kaydırır.", 2, 2, (a, c) => shift(a, c, true), { lift: true, xlfn: true }),
  fn("BITRSHIFT", "BİTSAĞAKAYDIR", "engineering", "sayı; kaydırma_miktarı", "Sayıyı belirtilen bit kadar sağa kaydırır.", 2, 2, (a, c) => shift(a, c, false), { lift: true, xlfn: true }),
  fn("DELTA", "DELTA", "engineering", "sayı1; [sayı2]", "İki değerin eşit olup olmadığını sınar (eşitse 1).", 1, 2, (a, c) => (num(a[0], c) === (a[1] === undefined || a[1] === null ? 0 : num(a[1], c)) ? 1 : 0), { lift: true }),
  fn("GESTEP", "BESINIR", "engineering", "sayı; [adım]", "Sayı eşik değerinden büyük ya da eşitse 1 verir.", 1, 2, (a, c) => (num(a[0], c) >= (a[1] === undefined || a[1] === null ? 0 : num(a[1], c)) ? 1 : 0), { lift: true }),
  fn("CONVERT", "ÇEVİR", "engineering", "sayı; birimden; birime", "Sayıyı bir ölçü biriminden diğerine dönüştürür.", 3, 3, (a, c) => convert(a, c), { lift: true }),
];

