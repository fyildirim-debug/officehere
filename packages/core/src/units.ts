/**
 * Ofis dosya biçimlerinde kullanılan birimler arası dönüşümler.
 * ECMA-376: 1 inç = 914400 EMU = 1440 twip = 72 pt; ekranda 1 inç = 96 px.
 */

export const EMU_PER_INCH = 914400;
export const EMU_PER_PX = EMU_PER_INCH / 96;
export const EMU_PER_PT = EMU_PER_INCH / 72;
export const TWIPS_PER_INCH = 1440;
export const PX_PER_INCH = 96;
export const PX_PER_CM = 96 / 2.54;

export const emuToPx = (emu: number) => emu / EMU_PER_PX;
export const pxToEmu = (px: number) => Math.round(px * EMU_PER_PX);
export const emuToPt = (emu: number) => emu / EMU_PER_PT;
export const ptToEmu = (pt: number) => Math.round(pt * EMU_PER_PT);
export const twipsToPx = (twips: number) => (twips / TWIPS_PER_INCH) * PX_PER_INCH;
export const pxToTwips = (px: number) => Math.round((px / PX_PER_INCH) * TWIPS_PER_INCH);
export const ptToPx = (pt: number) => (pt * 96) / 72;
export const pxToPt = (px: number) => (px * 72) / 96;
export const cmToPx = (cm: number) => cm * PX_PER_CM;
export const pxToCm = (px: number) => px / PX_PER_CM;

/** Sayfa boyutları (px, 96 dpi). */
export const PAGE_SIZES = {
  A4: { name: "A4", width: cmToPx(21), height: cmToPx(29.7) },
  A3: { name: "A3", width: cmToPx(29.7), height: cmToPx(42) },
  A5: { name: "A5", width: cmToPx(14.8), height: cmToPx(21) },
  Letter: { name: "Letter", width: 8.5 * 96, height: 11 * 96 },
  Legal: { name: "Legal", width: 8.5 * 96, height: 14 * 96 },
} as const;

export type PageSizeName = keyof typeof PAGE_SIZES;
