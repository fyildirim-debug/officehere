/**
 * Çalışma kitabı temaları (Sayfa Düzeni > Temalar): renk şeması (ECMA-376
 * a:clrScheme), yazı tipi şeması (a:fontScheme). Tema değişince tema
 * renk paletinden seçilmiş hücre renkleri yeni temanın karşılığına
 * eşlenir, tema yazı tipi kullanan hücreler yeni gövde yazı tipini alır.
 */
import { themeShades, THEMES } from "@officehere/core";
import type { CellStyle } from "./styles";

export interface WorkbookTheme {
  name: string;
  /** lt1, dk1, lt2, dk2, accent1..accent6 */
  colors: [string, string, string, string, string, string, string, string, string, string];
  colorsName: string;
  majorFont: string;
  minorFont: string;
  fontsName: string;
}

export interface ColorSet {
  name: string;
  colors: WorkbookTheme["colors"];
}

export interface FontSet {
  name: string;
  major: string;
  minor: string;
}

const core = (name: string) => THEMES.find((t) => t.name === name)!;

export const COLOR_SETS: ColorSet[] = [
  { name: "OfficeHere", colors: core("OfficeHere").colors },
  { name: "Klasik", colors: core("Klasik").colors },
  { name: "Gri Tonlamalı", colors: ["#FFFFFF", "#000000", "#F2F2F2", "#000000", "#DDDDDD", "#B2B2B2", "#969696", "#808080", "#5F5F5F", "#4D4D4D"] },
  { name: "Mavi", colors: ["#FFFFFF", "#000000", "#DEF5FA", "#17406D", "#0F6FC6", "#009DD9", "#0BD0D9", "#10CF9B", "#7CCA62", "#A5C249"] },
  { name: "Yeşil", colors: ["#FFFFFF", "#000000", "#E2EFDA", "#455F51", "#549E39", "#8AB833", "#C0CF3A", "#029676", "#4AB5C4", "#0989B1"] },
  { name: "Sarı Turuncu", colors: ["#FFFFFF", "#000000", "#FBE6CD", "#4E3B30", "#E48312", "#BD582C", "#865640", "#9B8357", "#C2BC80", "#94A088"] },
  { name: "Kırmızı", colors: ["#FFFFFF", "#000000", "#F4DCCB", "#323232", "#A5300F", "#D55816", "#E19825", "#B19C7D", "#7F5F52", "#B27D49"] },
  { name: "Mor", colors: ["#FFFFFF", "#000000", "#EAE5EB", "#373545", "#AD84C6", "#8784C7", "#5D739A", "#6997AF", "#84ACB6", "#6F8183"] },
  { name: "Gece", colors: core("Gece").colors },
  { name: "Toprak", colors: core("Toprak").colors },
  { name: "Okyanus", colors: core("Okyanus").colors },
];

export const FONT_SETS: FontSet[] = [
  { name: "Calibri", major: "Calibri Light", minor: "Calibri" },
  { name: "Aptos", major: "Aptos Display", minor: "Aptos" },
  { name: "Cambria-Calibri", major: "Cambria", minor: "Calibri" },
  { name: "Arial", major: "Arial", minor: "Arial" },
  { name: "Times New Roman", major: "Times New Roman", minor: "Times New Roman" },
  { name: "Georgia", major: "Georgia", minor: "Georgia" },
  { name: "Trebuchet MS", major: "Trebuchet MS", minor: "Trebuchet MS" },
  { name: "Verdana", major: "Verdana", minor: "Verdana" },
  { name: "Segoe UI", major: "Segoe UI Light", minor: "Segoe UI" },
];

function theme(name: string, colorsName: string, fontsName: string): WorkbookTheme {
  const c = COLOR_SETS.find((x) => x.name === colorsName)!;
  const f = FONT_SETS.find((x) => x.name === fontsName)!;
  return { name, colors: [...c.colors] as WorkbookTheme["colors"], colorsName, majorFont: f.major, minorFont: f.minor, fontsName };
}

/** Tema galerisi. Varsayılan "OfficeHere": güncel renkler, Calibri yazı tipleri. */
export const THEME_LIST: WorkbookTheme[] = [
  theme("OfficeHere", "OfficeHere", "Calibri"),
  theme("Klasik", "Klasik", "Calibri"),
  theme("Modern", "OfficeHere", "Aptos"),
  theme("Gece", "Gece", "Georgia"),
  theme("Toprak", "Toprak", "Cambria-Calibri"),
  theme("Okyanus", "Okyanus", "Trebuchet MS"),
  theme("Doğa", "Yeşil", "Verdana"),
  theme("Gün Batımı", "Sarı Turuncu", "Arial"),
  theme("Lavanta", "Mor", "Segoe UI"),
  theme("Kurumsal", "Mavi", "Arial"),
];

export function defaultTheme(): WorkbookTheme {
  return structuredClone(THEME_LIST[0]);
}

/** Renk seçicideki tema paleti (ui ColorPalette biçimi). */
export function paletteOf(t: WorkbookTheme): { name: string; colors: WorkbookTheme["colors"]; majorFont: string; minorFont: string } {
  return { name: t.name, colors: t.colors, majorFont: t.majorFont, minorFont: t.minorFont };
}

/** Eski temanın palet renklerini (ana renk + 5 ton) yenisine eşleyen tablo. */
export function themeColorMap(oldColors: readonly string[], newColors: readonly string[]): Map<string, string> {
  const m = new Map<string, string>();
  oldColors.forEach((c, i) => {
    const a = themeShades(c);
    const b = themeShades(newColors[i]);
    a.forEach((x, k) => {
      const key = x.toUpperCase();
      if (!m.has(key)) m.set(key, b[k].toUpperCase());
    });
  });
  // Değişmeyen renkleri tablodan çıkar
  for (const [k, v] of m) if (k === v) m.delete(k);
  return m;
}

/** Stilin renklerini eşleme tablosuna göre değiştirir (değişiklik yoksa null). */
export function remapStyle(s: CellStyle, map: Map<string, string>): CellStyle | null {
  let changed = false;
  const out: CellStyle = { ...s };
  const fix = (c: string | undefined) => {
    if (!c) return c;
    const n = map.get(c.toUpperCase());
    if (n) {
      changed = true;
      return n;
    }
    return c;
  };
  out.color = fix(s.color);
  out.fill = fix(s.fill);
  if (s.pattern) out.pattern = { ...s.pattern, color: fix(s.pattern.color) ?? s.pattern.color };
  for (const k of ["bTop", "bBottom", "bLeft", "bRight"] as const) {
    const b = s[k];
    if (b?.color) out[k] = { ...b, color: fix(b.color) };
  }
  if (!changed) return null;
  for (const k of Object.keys(out) as (keyof CellStyle)[]) if (out[k] === undefined) delete out[k];
  return out;
}

/** theme1.xml renk/yazı tipi şemalarını okur. */
export function parseThemeXml(xml: string | undefined): Partial<WorkbookTheme> | null {
  if (!xml) return null;
  const names = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const colors: string[] = [];
  for (const n of names) {
    const m = new RegExp(`<a:${n}>\\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")`).exec(xml);
    if (!m) return null;
    colors.push("#" + (m[1] ?? m[2]).toUpperCase());
  }
  const major = /<a:majorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
  const minor = /<a:minorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
  const name = /<a:theme\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1];
  const colorsName = /<a:clrScheme\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1];
  const fontsName = /<a:fontScheme\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1];
  return { name, colors: colors as WorkbookTheme["colors"], majorFont: major || undefined, minorFont: minor || undefined, colorsName, fontsName };
}

function escAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

/** theme1.xml içindeki renk ve yazı tipi şemalarını temayla değiştirir (biçim şeması korunur). */
export function patchThemeXml(xml: string, t: WorkbookTheme): string {
  const c = t.colors.map((x) => x.replace("#", "").toUpperCase());
  const clr = `<a:clrScheme name="${escAttr(t.colorsName)}"><a:dk1><a:sysClr val="windowText" lastClr="${c[1]}"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="${c[0]}"/></a:lt1><a:dk2><a:srgbClr val="${c[3]}"/></a:dk2><a:lt2><a:srgbClr val="${c[2]}"/></a:lt2>${c
    .slice(4)
    .map((x, i) => `<a:accent${i + 1}><a:srgbClr val="${x}"/></a:accent${i + 1}>`)
    .join("")}<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>`;
  const fonts = `<a:fontScheme name="${escAttr(t.fontsName)}"><a:majorFont><a:latin typeface="${escAttr(t.majorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${escAttr(t.minorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`;
  let out = xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, clr).replace(/<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/, fonts);
  out = out.replace(/(<a:theme\b[^>]*\bname=")[^"]*(")/, `$1${escAttr(t.name)}$2`);
  return out;
}
