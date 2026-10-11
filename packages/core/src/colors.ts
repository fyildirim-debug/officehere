/**
 * Renk seçicilerde kullanılan tema ve standart renkler.
 * Tema sütunları ECMA-376 DrawingML tema renk yuvalarına karşılık gelir
 * (lt1, dk1, lt2, dk2, accent1..6); tonlar aynı yüzdelerle türetilir.
 */

export interface ThemePalette {
  name: string;
  /** lt1, dk1, lt2, dk2, accent1..accent6 */
  colors: [string, string, string, string, string, string, string, string, string, string];
  majorFont: string;
  minorFont: string;
}

export const DEFAULT_THEME: ThemePalette = {
  name: "OfficeHere",
  colors: ["#FFFFFF", "#000000", "#E8E8E8", "#0E2841", "#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"],
  majorFont: "Aptos Display",
  minorFont: "Aptos",
};

export const THEMES: ThemePalette[] = [
  DEFAULT_THEME,
  {
    name: "Klasik",
    colors: ["#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
    majorFont: "Calibri Light",
    minorFont: "Calibri",
  },
  {
    name: "Gece",
    colors: ["#FFFFFF", "#000000", "#D5DCE4", "#1F2A44", "#2F5597", "#C55A11", "#548235", "#BF9000", "#7030A0", "#2E75B6"],
    majorFont: "Georgia",
    minorFont: "Georgia",
  },
  {
    name: "Toprak",
    colors: ["#FFFFFF", "#000000", "#EEECE1", "#4A3B2A", "#9C6B30", "#B5562C", "#7A8C3A", "#C9A227", "#6E5A4B", "#3E7C7B"],
    majorFont: "Cambria",
    minorFont: "Calibri",
  },
  {
    name: "Okyanus",
    colors: ["#FFFFFF", "#000000", "#DCEAF2", "#0B3954", "#087E8B", "#FF5A5F", "#3C91E6", "#9FD356", "#342E37", "#FA824C"],
    majorFont: "Trebuchet MS",
    minorFont: "Trebuchet MS",
  },
];

export const STANDARD_COLORS = [
  "#C00000",
  "#FF0000",
  "#FFC000",
  "#FFFF00",
  "#92D050",
  "#00B050",
  "#00B0F0",
  "#0070C0",
  "#002060",
  "#7030A0",
];

export const HIGHLIGHT_COLORS = [
  "#FFFF00",
  "#00FF00",
  "#00FFFF",
  "#FF00FF",
  "#0000FF",
  "#FF0000",
  "#000080",
  "#008080",
  "#008000",
  "#800080",
  "#800000",
  "#808000",
  "#808080",
  "#C0C0C0",
  "#000000",
];

function clamp(v: number) {
  return Math.max(0, Math.min(255, Math.round(v)));
}

export function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h.padEnd(6, "0").slice(0, 6);
  return [parseInt(full.slice(0, 2), 16), parseInt(full.slice(2, 4), 16), parseInt(full.slice(4, 6), 16)];
}

export function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((v) => clamp(v).toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

/** Rengi açar (amount > 0) ya da koyulaştırır (amount < 0); -1..1. */
export function tint(hex: string, amount: number): string {
  const [r, g, b] = hexToRgb(hex);
  if (amount >= 0) return rgbToHex(r + (255 - r) * amount, g + (255 - g) * amount, b + (255 - b) * amount);
  const k = 1 + amount;
  return rgbToHex(r * k, g * k, b * k);
}

/** Renk seçicideki tema sütunu: ana renk + 5 ton (Office'teki %80, %60, %40 açık; %25, %50 koyu). */
export function themeShades(hex: string): string[] {
  const [r, g, b] = hexToRgb(hex);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  if (lum > 0.95) return [hex, tint(hex, -0.05), tint(hex, -0.15), tint(hex, -0.25), tint(hex, -0.35), tint(hex, -0.5)];
  if (lum < 0.05) return [hex, tint(hex, 0.5), tint(hex, 0.35), tint(hex, 0.25), tint(hex, 0.15), tint(hex, 0.05)];
  return [hex, tint(hex, 0.8), tint(hex, 0.6), tint(hex, 0.4), tint(hex, -0.25), tint(hex, -0.5)];
}

/** Okunabilir metin rengi (koyu zemin için beyaz). */
export function contrastText(hex: string): string {
  const [r, g, b] = hexToRgb(hex);
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#000000" : "#FFFFFF";
}
