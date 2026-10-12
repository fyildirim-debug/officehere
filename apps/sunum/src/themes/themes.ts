/**
 * OfficeHere Sunum'un kendi tasarım temaları. Renkler, yazı tipleri ve
 * süsleme şekilleri bu projeye özgüdür; tüm süslemeler tema renk yuvalarına
 * (accent1…) bağlıdır, böylece çeşitleme/tema değişince güncellenir.
 */
import { makeSp, NO_LINE, solid } from "../model/factory";
import type { Color, Fill, SchemeColorName, SpShape } from "../model/types";
import type { ThemeStyle } from "./build";

export type ColorScheme = Record<SchemeColorName, string>;

export interface ThemeDef {
  id: string;
  name: string;
  fonts: { major: string; minor: string };
  variants: { name: string; colors: ColorScheme }[];
  style: ThemeStyle;
}

const cs = (lt1: string, dk1: string, lt2: string, dk2: string, a: string[], hlink = "#0563C1", fol = "#954F72"): ColorScheme => ({
  lt1,
  dk1,
  lt2,
  dk2,
  accent1: a[0],
  accent2: a[1],
  accent3: a[2],
  accent4: a[3],
  accent5: a[4],
  accent6: a[5],
  hlink,
  folHlink: fol,
});

const S = (name: SchemeColorName, mods: Partial<Color> = {}): Color => ({ scheme: name, ...mods });

const deco = (geom: string, x: number, y: number, w: number, h: number, fill: Fill, extra: Partial<SpShape> = {}): SpShape =>
  makeSp(geom, x, y, w, h, { name: "Süsleme", fill, line: NO_LINE, ...extra });

/** Özel yol: alt dalga. */
function wave(x: number, y: number, w: number, h: number, fill: Fill, phase = 0): SpShape {
  const cmds =
    phase === 0
      ? [
          { c: "M" as const, pts: [0, 40] },
          { c: "C" as const, pts: [250, 0, 500, 90, 750, 45] },
          { c: "C" as const, pts: [880, 22, 950, 30, 1000, 20] },
          { c: "L" as const, pts: [1000, 100] },
          { c: "L" as const, pts: [0, 100] },
          { c: "Z" as const, pts: [] },
        ]
      : [
          { c: "M" as const, pts: [0, 20] },
          { c: "C" as const, pts: [200, 70, 420, 10, 640, 40] },
          { c: "C" as const, pts: [800, 62, 900, 50, 1000, 30] },
          { c: "L" as const, pts: [1000, 100] },
          { c: "L" as const, pts: [0, 100] },
          { c: "Z" as const, pts: [] },
        ];
  return makeSp("custom", x, y, w, h, { name: "Dalga", fill, line: NO_LINE, paths: [{ w: 1000, h: 100, cmds }] });
}

const grad = (a: Color, b: Color, angle = 90): Fill => ({ type: "gradient", kind: "linear", angle, stops: [{ pos: 0, color: a }, { pos: 1, color: b }] });

export const THEMES: ThemeDef[] = [
  {
    id: "varsayilan",
    name: "OfficeHere",
    fonts: { major: "Aptos Display", minor: "Aptos" },
    variants: [
      { name: "OfficeHere", colors: cs("#FFFFFF", "#000000", "#E8E8E8", "#0E2841", ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"], "#467886", "#96607D") },
      { name: "Klasik", colors: cs("#FFFFFF", "#000000", "#E7E6E6", "#44546A", ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]) },
      { name: "Mavi Tonlar", colors: cs("#FFFFFF", "#000000", "#DEEAF6", "#17406D", ["#0F6FC6", "#009DD9", "#0BD0D9", "#10CF9B", "#7CCA62", "#A5C249"]) },
      { name: "Kırmızı Turuncu", colors: cs("#FFFFFF", "#000000", "#F4E7E2", "#505046", ["#E84C22", "#FFBD47", "#B64926", "#FF8427", "#CC9900", "#B22600"]) },
    ],
    style: { bg: solid(S("lt1")) },
  },
  {
    id: "kiyi",
    name: "Kıyı",
    fonts: { major: "Trebuchet MS", minor: "Trebuchet MS" },
    variants: [
      { name: "Kıyı", colors: cs("#FFFFFF", "#1B2A3A", "#E3F1F6", "#0B3954", ["#087E8B", "#3C91E6", "#FF5A5F", "#9FD356", "#FA824C", "#342E37"]) },
      { name: "Lagün", colors: cs("#FFFFFF", "#10303A", "#E0F4F1", "#0D4D4D", ["#14A098", "#0F6D8C", "#F2A541", "#6C5B7B", "#C06C84", "#355C7D"]) },
      { name: "Fırtına", colors: cs("#FFFFFF", "#202833", "#E6E9EE", "#2C3E50", ["#34495E", "#5D8AA8", "#E67E22", "#95A5A6", "#16A085", "#8E44AD"]) },
      { name: "Mercan", colors: cs("#FFFFFF", "#2B2B2B", "#FDEFEA", "#5A2A27", ["#E2725B", "#2A9D8F", "#E9C46A", "#264653", "#F4A261", "#8AB17D"]) },
    ],
    style: {
      bg: solid(S("lt1")),
      titleColor: S("dk2"),
      deco: (w, h) => [wave(0, h * 0.86, w, h * 0.14, solid(S("accent1", { alpha: 0.35 })), 1), wave(0, h * 0.89, w, h * 0.11, solid(S("accent1")))],
      titleDeco: (w, h) => [
        deco("rect", 0, 0, w, h, grad(S("accent1"), S("dk2"), 45), { name: "Arka Plan" }),
        wave(0, h * 0.7, w, h * 0.3, solid(S("lt1", { alpha: 0.18 })), 1),
        wave(0, h * 0.78, w, h * 0.22, solid(S("lt1", { alpha: 0.3 }))),
      ],
      titleSlideColor: S("lt1"),
      titleSlideSubColor: S("lt1", { lumMod: 0.95 }),
      area: { t: 0 },
    },
  },
  {
    id: "gunbatimi",
    name: "Gün Batımı",
    fonts: { major: "Georgia", minor: "Verdana" },
    variants: [
      { name: "Gün Batımı", colors: cs("#FFFFFF", "#3A1F1D", "#FFF1E6", "#6B2D2A", ["#F26B38", "#F7B32B", "#EC4E7C", "#8E3A59", "#4E8098", "#2D8C6F"]) },
      { name: "Şafak", colors: cs("#FFFFFF", "#2E2440", "#F6EEFA", "#4B3869", ["#F08A5D", "#B83B5E", "#6A2C70", "#F9ED69", "#3FC1C9", "#364F6B"]) },
      { name: "Kum", colors: cs("#FFFFFF", "#3D3027", "#F7F0E6", "#6E5643", ["#D97941", "#C2A35D", "#8C5B3E", "#5E8C61", "#4F6D7A", "#B5523B"]) },
      { name: "Gül", colors: cs("#FFFFFF", "#3B1F2B", "#FCECF1", "#6D2E46", ["#D9455F", "#F79D84", "#A26769", "#582C4D", "#ECE2D0", "#3E7CB1"]) },
    ],
    style: {
      bg: grad(S("accent2", { lumMod: 0.2, lumOff: 0.8 }), S("lt1"), 90),
      titleColor: S("dk2"),
      deco: (w, h) => [deco("rect", 0, 0, w, h * 0.022, grad(S("accent1"), S("accent3"), 0), { name: "Üst Şerit" })],
      titleDeco: (w, h) => [
        deco("rect", 0, 0, w, h, grad(S("accent2"), S("accent3"), 100), { name: "Arka Plan" }),
        deco("ellipse", w * 0.62, -h * 0.35, w * 0.6, w * 0.6, solid(S("lt1", { alpha: 0.12 })), { name: "Güneş" }),
        deco("ellipse", w * 0.72, -h * 0.18, w * 0.38, w * 0.38, solid(S("lt1", { alpha: 0.16 })), { name: "Güneş 2" }),
      ],
      titleSlideColor: S("lt1"),
      titleSlideSubColor: S("lt1"),
    },
  },
  {
    id: "orman",
    name: "Orman",
    fonts: { major: "Cambria", minor: "Calibri" },
    variants: [
      { name: "Orman", colors: cs("#FFFFFF", "#1E2B22", "#E6EFE4", "#1F3B2C", ["#7CB518", "#F3A712", "#5C8001", "#FBB02D", "#2E86AB", "#E4572E"], "#9BD1E5", "#C9A0DC") },
      { name: "Yosun", colors: cs("#FFFFFF", "#26302B", "#EEF2EC", "#2F4538", ["#A3B18A", "#DDA15E", "#588157", "#BC6C25", "#7F9C96", "#E9EDC9"], "#9BD1E5", "#C9A0DC") },
      { name: "Çam", colors: cs("#FFFFFF", "#132A20", "#E3EEE8", "#0F3D2E", ["#2BB673", "#F9C80E", "#43BCCD", "#EA3546", "#F86624", "#9AE19D"], "#9BD1E5", "#C9A0DC") },
      { name: "Sonbahar", colors: cs("#FFFFFF", "#2E2117", "#F4EDE4", "#3D2B1F", ["#D98E04", "#B3541E", "#8AA624", "#F2C14E", "#5F7367", "#C1666B"], "#9BD1E5", "#C9A0DC") },
    ],
    style: {
      bg: solid(S("dk2")),
      titleColor: S("lt1"),
      bodyColor: S("lt1", { lumMod: 0.95 }),
      deco: (w, h) => [
        deco("ellipse", -w * 0.08, h * 0.78, w * 0.22, w * 0.22, solid(S("accent1", { alpha: 0.25 })), { name: "Yaprak 1" }),
        deco("ellipse", w * 0.06, h * 0.9, w * 0.12, w * 0.12, solid(S("accent3", { alpha: 0.35 })), { name: "Yaprak 2" }),
        deco("rect", w * 0.069, h * 0.25, w * 0.08, 4, solid(S("accent1")), { name: "Çizgi" }),
      ],
      titleDeco: (w, h) => [
        deco("ellipse", w * 0.7, -h * 0.25, w * 0.45, w * 0.45, solid(S("accent1", { alpha: 0.22 })), { name: "Yaprak 1" }),
        deco("ellipse", w * 0.82, h * 0.45, w * 0.28, w * 0.28, solid(S("accent3", { alpha: 0.3 })), { name: "Yaprak 2" }),
        deco("ellipse", -w * 0.1, h * 0.7, w * 0.3, w * 0.3, solid(S("accent2", { alpha: 0.18 })), { name: "Yaprak 3" }),
      ],
      titleSlideColor: S("lt1"),
      titleSlideSubColor: S("accent1", { lumMod: 0.4, lumOff: 0.6 }),
    },
  },
  {
    id: "murekkep",
    name: "Mürekkep",
    fonts: { major: "Segoe UI", minor: "Segoe UI" },
    variants: [
      { name: "Mürekkep", colors: cs("#FFFFFF", "#111111", "#E9E9EB", "#22252A", ["#4CC9F0", "#F72585", "#7209B7", "#FFD166", "#06D6A0", "#EF476F"], "#8ECAE6", "#CDB4DB") },
      { name: "Grafit", colors: cs("#FFFFFF", "#121212", "#ECECEC", "#2B2B2B", ["#F5A623", "#50E3C2", "#B8E986", "#4A90E2", "#D0021B", "#9013FE"], "#8ECAE6", "#CDB4DB") },
      { name: "Neon", colors: cs("#FFFFFF", "#0B0B12", "#E8E8F0", "#1A1A2E", ["#00F5D4", "#F15BB5", "#FEE440", "#00BBF9", "#9B5DE5", "#FF006E"], "#8ECAE6", "#CDB4DB") },
      { name: "Sepya", colors: cs("#FFFFFF", "#1E1A16", "#EFE9E1", "#3A322B", ["#D4A373", "#CCD5AE", "#E9EDC9", "#FAEDCD", "#B5838D", "#6D6875"], "#8ECAE6", "#CDB4DB") },
    ],
    style: {
      bg: solid(S("dk2")),
      titleColor: S("lt1"),
      bodyColor: S("lt1", { lumMod: 0.85 }),
      deco: (w, h) => [deco("rect", w * 0.069, h * 0.245, w * 0.862, 3, solid(S("accent1")), { name: "Başlık Çizgisi" })],
      titleDeco: (w, h) => [
        deco("rect", 0, 0, w * 0.035, h, solid(S("accent1")), { name: "Sol Şerit" }),
        deco("rect", w * 0.125, h * 0.525, w * 0.3, 4, solid(S("accent2")), { name: "Çizgi" }),
      ],
      titleAlign: "l",
      titleSlideColor: S("lt1"),
      titleSlideSubColor: S("lt1", { lumMod: 0.75 }),
    },
  },
  {
    id: "kagit",
    name: "Kâğıt",
    fonts: { major: "Book Antiqua", minor: "Garamond" },
    variants: [
      { name: "Kâğıt", colors: cs("#FFFFFF", "#2F2A24", "#F6F0E4", "#4A3F35", ["#A0522D", "#6B8E23", "#B8860B", "#4682B4", "#8B4513", "#708090"]) },
      { name: "Parşömen", colors: cs("#FFFFFF", "#3B2F2F", "#F3EAD3", "#5C4033", ["#8C6A4F", "#A68A64", "#7F5539", "#B08968", "#9C6644", "#606C38"]) },
      { name: "Defter", colors: cs("#FFFFFF", "#22313F", "#EEF3F7", "#34495E", ["#2C7BB6", "#D7191C", "#FDAE61", "#ABD9E9", "#1A9641", "#7B3294"]) },
      { name: "Lavanta", colors: cs("#FFFFFF", "#2D2A32", "#F3F0F7", "#4B4453", ["#7E6B8F", "#C4A287", "#96E6B3", "#DA3E52", "#A3D9FF", "#6C4E31"]) },
    ],
    style: {
      bg: solid(S("lt2")),
      titleColor: S("accent1", { lumMod: 0.75 }),
      area: { l: 120 },
      deco: (w, h) => [
        deco("rect", 0, 0, w * 0.055, h, solid(S("accent1", { lumMod: 0.75 })), { name: "Sol Şerit" }),
        deco("rect", w * 0.062, 0, 2, h, solid(S("accent1", { lumMod: 0.75 })), { name: "İnce Çizgi" }),
      ],
      titleDeco: (w, h) => [
        deco("rect", 0, 0, w, h, solid(S("lt2")), { name: "Arka Plan" }),
        deco("frame", w * 0.04, h * 0.07, w * 0.92, h * 0.86, solid(S("accent1", { lumMod: 0.75 })), { name: "Çerçeve", adj: { adj1: 1200 } }),
      ],
    },
  },
  {
    id: "geometri",
    name: "Geometri",
    fonts: { major: "Calibri Light", minor: "Calibri" },
    variants: [
      { name: "Geometri", colors: cs("#FFFFFF", "#262626", "#EDEFF2", "#2E3A59", ["#3F72AF", "#F9A826", "#E84545", "#52B788", "#8338EC", "#00B4D8"]) },
      { name: "Kübik", colors: cs("#FFFFFF", "#1F1F1F", "#F0F0F0", "#3A3A3A", ["#FF6B35", "#004E89", "#1A936F", "#C6DABF", "#F7C548", "#88498F"]) },
      { name: "Prizma", colors: cs("#FFFFFF", "#202020", "#F2F4F8", "#283593", ["#5C6BC0", "#26A69A", "#EF5350", "#FFCA28", "#AB47BC", "#8D6E63"]) },
      { name: "Mozaik", colors: cs("#FFFFFF", "#252422", "#F5F3EF", "#403D39", ["#EB5E28", "#CCC5B9", "#2A9D8F", "#E9C46A", "#264653", "#E76F51"]) },
    ],
    style: {
      bg: solid(S("lt1")),
      titleColor: S("dk2"),
      deco: (w, h) => [
        deco("rtTriangle", w * 0.9, 0, w * 0.1, h * 0.18, solid(S("accent1")), { name: "Üçgen 1", flipH: true, flipV: true }),
        deco("rtTriangle", w * 0.94, 0, w * 0.06, h * 0.32, solid(S("accent2", { alpha: 0.8 })), { name: "Üçgen 2", flipH: true, flipV: true }),
      ],
      titleDeco: (w, h) => [
        deco("rtTriangle", w * 0.55, 0, w * 0.45, h, solid(S("accent1")), { name: "Üçgen 1", flipH: true, flipV: true }),
        deco("rtTriangle", w * 0.75, 0, w * 0.25, h * 0.6, solid(S("accent2", { alpha: 0.85 })), { name: "Üçgen 2", flipH: true, flipV: true }),
        deco("rtTriangle", 0, h * 0.75, w * 0.18, h * 0.25, solid(S("accent3", { alpha: 0.85 })), { name: "Üçgen 3" }),
      ],
      titleAlign: "l",
      titleSlideColor: S("dk2"),
    },
  },
  {
    id: "geceyarisi",
    name: "Gece Yarısı",
    fonts: { major: "Aptos Display", minor: "Aptos" },
    variants: [
      { name: "Gece Yarısı", colors: cs("#FFFFFF", "#0A1128", "#E5ECF6", "#0A1128", ["#1282A2", "#FEFCFB", "#00D1FF", "#FFB703", "#FB8500", "#8ECAE6"], "#8ECAE6", "#CDB4DB") },
      { name: "Kutup", colors: cs("#FFFFFF", "#081C15", "#E3F2EC", "#081C15", ["#52B788", "#D8F3DC", "#74C69D", "#FFD166", "#118AB2", "#EF476F"], "#8ECAE6", "#CDB4DB") },
      { name: "Galaksi", colors: cs("#FFFFFF", "#10002B", "#EEE6F8", "#10002B", ["#9D4EDD", "#E0AAFF", "#C77DFF", "#FFD60A", "#00F5D4", "#FF5D8F"], "#8ECAE6", "#CDB4DB") },
      { name: "Yakut", colors: cs("#FFFFFF", "#1B0A0F", "#F7E7EA", "#2B0A12", ["#E63946", "#F1FAEE", "#A8DADC", "#FFB703", "#457B9D", "#F4A261"], "#8ECAE6", "#CDB4DB") },
    ],
    style: {
      bg: { type: "gradient", kind: "radial", angle: 0, stops: [{ pos: 0, color: S("dk2", { lumMod: 0.75, lumOff: 0.25 }) }, { pos: 1, color: S("dk2") }] },
      titleColor: S("lt1"),
      bodyColor: S("lt1", { lumMod: 0.9 }),
      deco: (w, h) => [
        deco("line", w * 0.069, h * 0.93, w * 0.862, 0, { type: "none" }, { name: "Çizgi", line: { fill: { type: "solid", color: S("accent1") }, width: 1.5, dash: "solid" } }),
      ],
      titleDeco: (w, h) => [
        deco("ellipse", w * 0.15, h * 0.08, w * 0.7, w * 0.7, { type: "none" }, { name: "Halka", line: { fill: { type: "solid", color: S("accent1", { alpha: 0.6 }) }, width: 1.5, dash: "solid" } }),
        deco("ellipse", w * 0.25, h * 0.26, w * 0.5, w * 0.5, { type: "none" }, { name: "Halka 2", line: { fill: { type: "solid", color: S("accent3", { alpha: 0.35 }) }, width: 1, dash: "sysDash" } }),
      ],
      titleBg: { type: "gradient", kind: "radial", angle: 0, stops: [{ pos: 0, color: S("accent1", { lumMod: 0.5 }) }, { pos: 1, color: S("dk2") }] },
      titleSlideColor: S("lt1"),
      titleSlideSubColor: S("accent3"),
    },
  },
  {
    id: "cizgi",
    name: "Çizgi",
    fonts: { major: "Arial", minor: "Arial" },
    variants: [
      { name: "Çizgi", colors: cs("#FFFFFF", "#1A1A1A", "#EFEFEF", "#333333", ["#E63946", "#1D3557", "#457B9D", "#A8DADC", "#F4A261", "#2A9D8F"]) },
      { name: "Okul", colors: cs("#FFFFFF", "#1C1C1C", "#F1F3F5", "#212529", ["#228BE6", "#40C057", "#FAB005", "#FA5252", "#7950F2", "#15AABF"]) },
      { name: "Kurumsal", colors: cs("#FFFFFF", "#202124", "#EEF1F4", "#1F3A5F", ["#1F6FB2", "#5C9EAD", "#E4A11B", "#7A7A7A", "#2E8B57", "#B23A48"]) },
      { name: "Yeşil", colors: cs("#FFFFFF", "#1B1B1B", "#EDF6EF", "#1E3A2B", ["#2D6A4F", "#95D5B2", "#F4A261", "#264653", "#E76F51", "#6C757D"]) },
    ],
    style: {
      bg: solid(S("lt1")),
      titleColor: S("dk2"),
      titleBold: true,
      titleSize: 40,
      area: { l: 104 },
      deco: (w, h) => [
        deco("rect", w * 0.055, h * 0.085, w * 0.012, h * 0.15, solid(S("accent1")), { name: "Başlık Çubuğu" }),
        deco("rect", 0, h * 0.985, w, h * 0.015, solid(S("accent1")), { name: "Alt Çizgi" }),
      ],
      titleDeco: (w, h) => [
        deco("rect", 0, h * 0.52, w, h * 0.48, solid(S("accent1")), { name: "Alt Bant" }),
        deco("rect", 0, h * 0.5, w, h * 0.02, solid(S("accent2")), { name: "Çizgi" }),
      ],
      titleSlideSubColor: S("lt1"),
      secHeadUsesMaster: true,
    },
  },
  {
    id: "pastel",
    name: "Pastel",
    fonts: { major: "Verdana", minor: "Tahoma" },
    variants: [
      { name: "Pastel", colors: cs("#FFFFFF", "#3D3B4F", "#FBF5F3", "#5B5277", ["#E8A0BF", "#BA90C6", "#C0DBEA", "#FFD8A8", "#A3D2CA", "#F7A4A4"]) },
      { name: "Nane", colors: cs("#FFFFFF", "#2F3E46", "#F1FAF6", "#354F52", ["#84A98C", "#CAD2C5", "#F6BD60", "#F28482", "#84A59D", "#F5CAC3"]) },
      { name: "Şeker", colors: cs("#FFFFFF", "#3A3042", "#FFF5FA", "#5E4B56", ["#FF99C8", "#FCF6BD", "#D0F4DE", "#A9DEF9", "#E4C1F9", "#FFB5A7"]) },
      { name: "Bulut", colors: cs("#FFFFFF", "#2B2D42", "#F4F6FB", "#4A4E69", ["#9A8C98", "#C9ADA7", "#8ECAE6", "#FFB4A2", "#B5E48C", "#CDB4DB"]) },
    ],
    style: {
      bg: solid(S("lt2")),
      titleColor: S("dk2"),
      deco: (w, h) => [
        deco("ellipse", w * 0.88, -h * 0.1, w * 0.18, w * 0.18, solid(S("accent1", { alpha: 0.45 })), { name: "Daire 1" }),
        deco("ellipse", w * 0.93, h * 0.17, w * 0.08, w * 0.08, solid(S("accent3", { alpha: 0.6 })), { name: "Daire 2" }),
      ],
      titleDeco: (w, h) => [
        deco("rect", 0, 0, w, h, solid(S("lt2")), { name: "Arka Plan" }),
        deco("ellipse", -w * 0.08, -h * 0.2, w * 0.32, w * 0.32, solid(S("accent1", { alpha: 0.5 })), { name: "Daire 1" }),
        deco("ellipse", w * 0.8, h * 0.62, w * 0.28, w * 0.28, solid(S("accent3", { alpha: 0.6 })), { name: "Daire 2" }),
        deco("ellipse", w * 0.7, h * 0.08, w * 0.07, w * 0.07, solid(S("accent4", { alpha: 0.8 })), { name: "Daire 3" }),
        deco("ellipse", w * 0.12, h * 0.78, w * 0.05, w * 0.05, solid(S("accent2", { alpha: 0.8 })), { name: "Daire 4" }),
      ],
      titleSlideColor: S("dk2"),
    },
  },
];

export function getThemeDef(id: string | undefined): ThemeDef | undefined {
  return THEMES.find((t) => t.id === id);
}

