import type { Color, SchemeColorName, Theme } from "./types";

/** tx1/bg1/tx2/bg2 eşlemesi (standart clrMap). */
export function mapSchemeAlias(name: string): SchemeColorName | undefined {
  switch (name) {
    case "tx1":
      return "dk1";
    case "bg1":
      return "lt1";
    case "tx2":
      return "dk2";
    case "bg2":
      return "lt2";
    case "dk1":
    case "lt1":
    case "dk2":
    case "lt2":
    case "accent1":
    case "accent2":
    case "accent3":
    case "accent4":
    case "accent5":
    case "accent6":
    case "hlink":
    case "folHlink":
      return name;
    default:
      return undefined;
  }
}

export function hexToRgb(hex: string): [number, number, number] {
  let h = hex.replace("#", "").trim();
  if (h.length === 3) h = h.split("").map((c) => c + c).join("");
  h = h.padEnd(6, "0").slice(0, 6);
  return [parseInt(h.slice(0, 2), 16) || 0, parseInt(h.slice(2, 4), 16) || 0, parseInt(h.slice(4, 6), 16) || 0];
}

const clamp255 = (v: number) => Math.max(0, Math.min(255, Math.round(v)));

export function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((v) => clamp255(v).toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Rengi tema ile çözer ve "#RRGGBB" döndürür (saydamlık hariç). */
export function resolveHex(c: Color | undefined, theme: Theme, fallback = "#000000"): string {
  if (!c) return fallback;
  let base = c.rgb ?? (c.scheme ? theme.colors[c.scheme] : undefined) ?? fallback;
  let [r, g, b] = hexToRgb(base);
  if (c.tint !== undefined) {
    r = r * c.tint + 255 * (1 - c.tint);
    g = g * c.tint + 255 * (1 - c.tint);
    b = b * c.tint + 255 * (1 - c.tint);
  }
  if (c.shade !== undefined) {
    r *= c.shade;
    g *= c.shade;
    b *= c.shade;
  }
  if (c.lumMod !== undefined || c.lumOff !== undefined) {
    const [h, s, l] = rgbToHsl(r, g, b);
    const nl = Math.max(0, Math.min(1, l * (c.lumMod ?? 1) + (c.lumOff ?? 0)));
    [r, g, b] = hslToRgb(h, s, nl);
  }
  base = rgbToHex(r, g, b);
  return base;
}

/** CSS renk dizgesi (saydamlık dahil). */
export function cssColor(c: Color | undefined, theme: Theme, fallback = "#000000"): string {
  const hex = resolveHex(c, theme, fallback);
  const a = c?.alpha;
  if (a === undefined || a >= 1) return hex;
  const [r, g, b] = hexToRgb(hex);
  return `rgba(${r},${g},${b},${Math.max(0, a).toFixed(3)})`;
}

export function luminance(hex: string): number {
  const [r, g, b] = hexToRgb(hex);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

export const rgb = (hex: string): Color => ({ rgb: hex.toUpperCase() });
export const scheme = (name: SchemeColorName, mods?: Partial<Color>): Color => ({ scheme: name, ...mods });

/** Renk seçicideki tema tonları (Office ile aynı oranlar). */
export function themeShadeMods(baseHex: string): Partial<Color>[] {
  const l = luminance(baseHex);
  if (l > 0.95)
    return [{ lumMod: 0.95 }, { lumMod: 0.85 }, { lumMod: 0.75 }, { lumMod: 0.65 }, { lumMod: 0.5 }];
  if (l < 0.05)
    return [{ lumMod: 0.5, lumOff: 0.5 }, { lumMod: 0.65, lumOff: 0.35 }, { lumMod: 0.75, lumOff: 0.25 }, { lumMod: 0.85, lumOff: 0.15 }, { lumMod: 0.95, lumOff: 0.05 }];
  return [{ lumMod: 0.2, lumOff: 0.8 }, { lumMod: 0.4, lumOff: 0.6 }, { lumMod: 0.6, lumOff: 0.4 }, { lumMod: 0.75 }, { lumMod: 0.5 }];
}

export function colorEquals(a: Color | undefined, b: Color | undefined): boolean {
  if (!a || !b) return a === b;
  return a.rgb === b.rgb && a.scheme === b.scheme && a.lumMod === b.lumMod && a.lumOff === b.lumOff && a.alpha === b.alpha && a.shade === b.shade && a.tint === b.tint;
}

const SCHEME_LABELS: Record<SchemeColorName, string> = {
  lt1: "Beyaz, Arka Plan 1",
  dk1: "Siyah, Metin 1",
  lt2: "Arka Plan 2",
  dk2: "Metin 2",
  accent1: "Vurgu 1",
  accent2: "Vurgu 2",
  accent3: "Vurgu 3",
  accent4: "Vurgu 4",
  accent5: "Vurgu 5",
  accent6: "Vurgu 6",
  hlink: "Köprü",
  folHlink: "İzlenen Köprü",
};

export function colorLabel(c: Color): string {
  if (c.scheme) {
    let s = SCHEME_LABELS[c.scheme];
    if (c.lumOff) s += `, %${Math.round(c.lumOff * 100)} Daha Açık`;
    else if (c.lumMod !== undefined && c.lumMod < 1) s += `, %${Math.round((1 - c.lumMod) * 100)} Daha Koyu`;
    return s;
  }
  return c.rgb ?? "";
}
