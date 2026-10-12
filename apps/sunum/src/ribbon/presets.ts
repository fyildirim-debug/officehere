/** Hazır biçim önayarları: Hızlı Stiller, gölge/parlama, gradyan, madde işaretleri, çizgi türleri. */
import type { AutoNumScheme, BulletDef, Color, DashType, Effects, Fill, Line, SchemeColorName, Shadow } from "../model/types";

export interface QuickStyle {
  id: string;
  label: string;
  fill: Fill;
  line: Line;
  effects?: Effects;
  text: Color;
}

const STYLE_COLORS: { key: SchemeColorName; label: string }[] = [
  { key: "dk1", label: "Koyu 1" },
  { key: "accent1", label: "Vurgu 1" },
  { key: "accent2", label: "Vurgu 2" },
  { key: "accent3", label: "Vurgu 3" },
  { key: "accent4", label: "Vurgu 4" },
  { key: "accent5", label: "Vurgu 5" },
  { key: "accent6", label: "Vurgu 6" },
];

const sc = (n: SchemeColorName, m: Partial<Color> = {}): Color => ({ scheme: n, ...m });
const ln = (c: Color, w = 1): Line => ({ fill: { type: "solid", color: c }, width: w, dash: "solid" });

export const SOFT_SHADOW: Shadow = { color: { rgb: "#000000", alpha: 0.4 }, blur: 5, dist: 3, dir: 45 };

export const QUICK_STYLES: QuickStyle[] = (() => {
  const out: QuickStyle[] = [];
  const rows: { name: string; make: (k: SchemeColorName) => Omit<QuickStyle, "id" | "label"> }[] = [
    { name: "Renkli Anahat", make: (k) => ({ fill: { type: "solid", color: sc("lt1") }, line: ln(sc(k)), text: sc(k === "dk1" ? "dk1" : k, k === "dk1" ? {} : { lumMod: 0.75 }) }) },
    { name: "Renkli Dolgu", make: (k) => ({ fill: { type: "solid", color: sc(k) }, line: ln(sc(k, { lumMod: 0.5 })), text: sc("lt1") }) },
    { name: "Hafif Efekt", make: (k) => ({ fill: { type: "solid", color: sc(k, { lumMod: 0.2, lumOff: 0.8 }) }, line: ln(sc(k)), text: sc(k, k === "dk1" ? {} : { lumMod: 0.5 }) }) },
    {
      name: "Orta Efekt",
      make: (k) => ({
        fill: { type: "gradient", kind: "linear", angle: 90, stops: [{ pos: 0, color: sc(k, { lumMod: 0.6, lumOff: 0.4 }) }, { pos: 1, color: sc(k) }] },
        line: { fill: { type: "none" }, width: 1, dash: "solid" },
        text: sc("lt1"),
      }),
    },
    { name: "Yoğun Efekt", make: (k) => ({ fill: { type: "solid", color: sc(k) }, line: ln(sc("lt1"), 1.5), effects: { shadow: SOFT_SHADOW }, text: sc("lt1") }) },
    { name: "Yarı Saydam", make: (k) => ({ fill: { type: "solid", color: sc(k, { alpha: 0.5 }) }, line: { fill: { type: "none" }, width: 1, dash: "solid" }, text: sc("dk1") }) },
  ];
  for (const r of rows) for (const c of STYLE_COLORS) out.push({ id: `${r.name}-${c.key}`, label: `${r.name} - ${c.label}`, ...r.make(c.key) });
  return out;
})();

export const SHADOW_PRESETS: { group: string; items: { label: string; shadow: Shadow | null }[] }[] = [
  { group: "Gölge Yok", items: [{ label: "Gölge Yok", shadow: null }] },
  {
    group: "Dış",
    items: [
      { label: "Uzaklık: Aşağı Sağ", shadow: { ...SOFT_SHADOW, dir: 45 } },
      { label: "Uzaklık: Aşağı", shadow: { ...SOFT_SHADOW, dir: 90 } },
      { label: "Uzaklık: Aşağı Sol", shadow: { ...SOFT_SHADOW, dir: 135 } },
      { label: "Uzaklık: Sağ", shadow: { ...SOFT_SHADOW, dir: 0 } },
      { label: "Uzaklık: Orta", shadow: { ...SOFT_SHADOW, dist: 0, blur: 8 } },
      { label: "Uzaklık: Sol", shadow: { ...SOFT_SHADOW, dir: 180 } },
      { label: "Uzaklık: Yukarı Sağ", shadow: { ...SOFT_SHADOW, dir: -45 } },
      { label: "Uzaklık: Yukarı", shadow: { ...SOFT_SHADOW, dir: -90 } },
      { label: "Uzaklık: Yukarı Sol", shadow: { ...SOFT_SHADOW, dir: -135 } },
    ],
  },
  {
    group: "Belirgin",
    items: [
      { label: "Belirgin: Aşağı Sağ", shadow: { color: { rgb: "#000000", alpha: 0.6 }, blur: 10, dist: 8, dir: 45 } },
      { label: "Belirgin: Aşağı", shadow: { color: { rgb: "#000000", alpha: 0.6 }, blur: 10, dist: 8, dir: 90 } },
      { label: "Belirgin: Orta", shadow: { color: { rgb: "#000000", alpha: 0.6 }, blur: 16, dist: 0, dir: 0 } },
    ],
  },
];

export const GLOW_PRESETS: { label: string; radius: number }[] = [
  { label: "5 pt parlama", radius: 6.7 },
  { label: "8 pt parlama", radius: 10.7 },
  { label: "11 pt parlama", radius: 14.7 },
  { label: "18 pt parlama", radius: 24 },
];

export const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

export const DASHES: { id: DashType; label: string }[] = [
  { id: "solid", label: "Düz" },
  { id: "sysDot", label: "Yuvarlak Nokta" },
  { id: "sysDash", label: "Kare Nokta" },
  { id: "dash", label: "Tire" },
  { id: "dashDot", label: "Tire Nokta" },
  { id: "lgDash", label: "Uzun Tire" },
  { id: "lgDashDot", label: "Uzun Tire Nokta" },
  { id: "lgDashDotDot", label: "Uzun Tire Nokta Nokta" },
];

export const ARROW_PRESETS: { label: string; head?: Line["head"]; tail?: Line["tail"] }[] = [
  { label: "Ok Yok" },
  { label: "Ok: Sona", tail: "triangle" },
  { label: "Ok: Başa", head: "triangle" },
  { label: "Çift Ok", head: "triangle", tail: "triangle" },
  { label: "Açık Ok: Sona", tail: "arrow" },
  { label: "Gizli Ok: Sona", tail: "stealth" },
  { label: "Oval Uçlar", head: "oval", tail: "oval" },
  { label: "Elmas Uçlar", head: "diamond", tail: "diamond" },
];

export const BULLET_PRESETS: { label: string; bullet: BulletDef }[] = [
  { label: "Yok", bullet: { type: "none" } },
  { label: "Dolu Yuvarlak Madde İşaretleri", bullet: { type: "char", char: "•", font: "Arial" } },
  { label: "İçi Boş Yuvarlak Madde İşaretleri", bullet: { type: "char", char: "o", font: "Courier New" } },
  { label: "Dolu Kare Madde İşaretleri", bullet: { type: "char", char: "▪", font: "Arial" } },
  { label: "İçi Boş Kare Madde İşaretleri", bullet: { type: "char", char: "□", font: "Arial" } },
  { label: "Yıldız Madde İşaretleri", bullet: { type: "char", char: "★", font: "Arial" } },
  { label: "Ok Madde İşaretleri", bullet: { type: "char", char: "➢", font: "Arial" } },
  { label: "Onay İşareti Madde İşaretleri", bullet: { type: "char", char: "✓", font: "Arial" } },
];

export const NUMBER_PRESETS: { label: string; scheme: AutoNumScheme | null; sample: string }[] = [
  { label: "Yok", scheme: null, sample: "" },
  { label: "1. 2. 3.", scheme: "arabicPeriod", sample: "1." },
  { label: "1) 2) 3)", scheme: "arabicParenR", sample: "1)" },
  { label: "I. II. III.", scheme: "romanUcPeriod", sample: "I." },
  { label: "A. B. C.", scheme: "alphaUcPeriod", sample: "A." },
  { label: "a) b) c)", scheme: "alphaLcParenR", sample: "a)" },
  { label: "a. b. c.", scheme: "alphaLcPeriod", sample: "a." },
  { label: "i. ii. iii.", scheme: "romanLcPeriod", sample: "i." },
];

export const SPACING_PRESETS: { label: string; value: number | undefined }[] = [
  { label: "Çok Sıkı", value: -3 },
  { label: "Sıkı", value: -1.5 },
  { label: "Normal", value: undefined },
  { label: "Gevşek", value: 3 },
  { label: "Çok Gevşek", value: 6 },
];

/** Dolgunun ana rengini alır (gradyan önayarları için). */
export function fillBaseColor(fill: Fill | undefined): Color {
  if (fill?.type === "solid") return fill.color;
  if (fill?.type === "gradient") return fill.stops[fill.stops.length - 1]?.color ?? { scheme: "accent1" };
  return { scheme: "accent1" };
}

export function gradientPresets(base: Color): { label: string; fill: Fill }[] {
  const light = (m: number): Color => ({ ...base, lumMod: (base.lumMod ?? 1) * (1 - m), lumOff: (base.lumOff ?? 0) + m });
  const dark = (m: number): Color => ({ ...base, lumMod: (base.lumMod ?? 1) * (1 - m) });
  const g = (a: Color, b: Color, angle: number, kind: "linear" | "radial" = "linear"): Fill => ({ type: "gradient", kind, angle, stops: [{ pos: 0, color: a }, { pos: 1, color: b }] });
  return [
    { label: "Doğrusal Çapraz - Sol Üstten Sağ Alta", fill: g(light(0.5), base, 45) },
    { label: "Doğrusal Aşağı", fill: g(light(0.5), base, 90) },
    { label: "Doğrusal Sağa", fill: g(light(0.5), base, 0) },
    { label: "Doğrusal Yukarı", fill: g(light(0.5), base, 270) },
    { label: "Ortadan", fill: g(light(0.6), base, 0, "radial") },
    { label: "Koyu Doğrusal Aşağı", fill: g(base, dark(0.5), 90) },
    { label: "Koyu Doğrusal Sağa", fill: g(base, dark(0.5), 0) },
    { label: "Koyu Ortadan", fill: g(base, dark(0.5), 0, "radial") },
  ];
}

export const WORDART_PRESETS: { label: string; color: Color; outline?: Color; shadow?: boolean; bold?: boolean; glow?: boolean }[] = [
  { label: "Dolgu: Siyah, Metin rengi 1; Gölge", color: { scheme: "dk1" }, shadow: true },
  { label: "Dolgu: Vurgu 1", color: { scheme: "accent1" }, bold: true },
  { label: "Dolgu: Vurgu 2; Anahat", color: { scheme: "accent2" }, outline: { scheme: "accent2", lumMod: 0.5 } },
  { label: "Dolgu: Beyaz; Anahat: Vurgu 1; Gölge", color: { scheme: "lt1" }, outline: { scheme: "accent1" }, shadow: true },
  { label: "Dolgu: Vurgu 4; Parlama", color: { scheme: "accent4" }, glow: true },
  { label: "Dolgu: Vurgu 5; Gölge", color: { scheme: "accent5" }, shadow: true, bold: true },
  { label: "Dolgu: Vurgu 6", color: { scheme: "accent6" }, bold: true },
  { label: "Dolgu: Gri; Gölge", color: { rgb: "#7F7F7F" }, shadow: true },
];
