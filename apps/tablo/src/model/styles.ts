/**
 * Hücre stilleri: düz nesne olarak tutulur ve tabloda tekilleştirilir
 * (her hücre yalnız stil numarası saklar).
 */

export type BorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "slantDashDot";

export interface BorderSide {
  style: BorderStyle;
  color?: string;
}

export type HAlign = "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
export type VAlign = "top" | "middle" | "bottom" | "justify" | "distributed";

export interface CellStyle {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "single" | "double";
  strike?: boolean;
  color?: string;
  fill?: string;
  /** Desen dolgusu (Hücreleri Biçimlendir > Dolgu). */
  pattern?: { type: string; color: string };
  hAlign?: HAlign;
  vAlign?: VAlign;
  wrap?: boolean;
  shrink?: boolean;
  indent?: number;
  /** Derece (-90..90) ya da 255 (dikey metin). */
  rotation?: number;
  numFmt?: string;
  bTop?: BorderSide;
  bBottom?: BorderSide;
  bLeft?: BorderSide;
  bRight?: BorderSide;
  locked?: boolean;
  hidden?: boolean;
  /** Hücre içi onay kutusu (Ekle > Onay Kutusu): değer DOĞRU/YANLIŞ, metin yerine kutu çizilir. */
  checkbox?: boolean;
}

/** Tema gövde yazı tipi (yazı tipi belirtilmemiş hücreler); tema değişince güncellenir. */
export let DEFAULT_FONT = "Calibri";

export function setDefaultFont(font: string): void {
  DEFAULT_FONT = font;
}
export const DEFAULT_SIZE = 11;

export class StyleTable {
  private styles: CellStyle[] = [{}];
  private index = new Map<string, number>([["{}", 0]]);

  get(id: number | undefined): CellStyle {
    return (id !== undefined && this.styles[id]) || this.styles[0];
  }

  intern(style: CellStyle): number {
    const clean = normalize(style);
    const key = styleKey(clean);
    let id = this.index.get(key);
    if (id === undefined) {
      id = this.styles.length;
      this.styles.push(clean);
      this.index.set(key, id);
    }
    return id;
  }

  /** Stile yama uygular ve yeni numarayı döndürür. */
  patch(id: number | undefined, patch: Partial<CellStyle>): number {
    const base = this.get(id);
    const next: CellStyle = { ...base };
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined || v === null) delete (next as Record<string, unknown>)[k];
      else (next as Record<string, unknown>)[k] = v;
    }
    return this.intern(next);
  }

  get size(): number {
    return this.styles.length;
  }

  all(): CellStyle[] {
    return this.styles;
  }
}

function normalize(s: CellStyle): CellStyle {
  const out: CellStyle = {};
  for (const [k, v] of Object.entries(s)) {
    // "locked: false" anlamlıdır (varsayılan kilitli)
    if (v === undefined || v === null || (v === false && k !== "locked")) continue;
    if (k === "locked" && v === true) continue;
    if (k === "numFmt" && v === "General") continue;
    if (k === "hAlign" && v === "general") continue;
    if (k === "indent" && v === 0) continue;
    if (k === "rotation" && v === 0) continue;
    if (k === "font" && v === DEFAULT_FONT) continue;
    if (k === "size" && v === DEFAULT_SIZE) continue;
    (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

function styleKey(s: CellStyle): string {
  const keys = Object.keys(s).sort();
  return JSON.stringify(s, keys.length ? [...keys, "style", "color", "type"] : undefined);
}

export function fontCss(s: CellStyle, zoom = 1, fontStack: (n: string) => string): string {
  const size = (s.size ?? DEFAULT_SIZE) * (96 / 72) * zoom;
  return `${s.italic ? "italic " : ""}${s.bold ? "bold " : ""}${size.toFixed(2)}px ${fontStack(s.font ?? DEFAULT_FONT)}`;
}

/** Hazır hücre stilleri (Giriş > Hücre Stilleri). */
export interface NamedCellStyle {
  id: string;
  label: string;
  group: string;
  style: CellStyle;
}

export const CELL_STYLES: NamedCellStyle[] = [
  { id: "normal", label: "Normal", group: "İyi, Kötü ve Nötr", style: {} },
  { id: "bad", label: "Kötü", group: "İyi, Kötü ve Nötr", style: { fill: "#FFC7CE", color: "#9C0006" } },
  { id: "good", label: "İyi", group: "İyi, Kötü ve Nötr", style: { fill: "#C6EFCE", color: "#006100" } },
  { id: "neutral", label: "Nötr", group: "İyi, Kötü ve Nötr", style: { fill: "#FFEB9C", color: "#9C5700" } },
  { id: "calc", label: "Hesaplama", group: "Veri ve Model", style: { fill: "#F2F2F2", color: "#FA7D00", bold: true, bTop: { style: "thin", color: "#7F7F7F" }, bBottom: { style: "thin", color: "#7F7F7F" }, bLeft: { style: "thin", color: "#7F7F7F" }, bRight: { style: "thin", color: "#7F7F7F" } } },
  { id: "check", label: "Denetim Hücresi", group: "Veri ve Model", style: { fill: "#A5A5A5", color: "#FFFFFF", bold: true, bTop: { style: "double", color: "#3F3F3F" }, bBottom: { style: "double", color: "#3F3F3F" }, bLeft: { style: "double", color: "#3F3F3F" }, bRight: { style: "double", color: "#3F3F3F" } } },
  { id: "explain", label: "Açıklayıcı Metin", group: "Veri ve Model", style: { italic: true, color: "#7F7F7F" } },
  { id: "input", label: "Giriş", group: "Veri ve Model", style: { fill: "#FFCC99", color: "#3F3F76", bTop: { style: "thin", color: "#7F7F7F" }, bBottom: { style: "thin", color: "#7F7F7F" }, bLeft: { style: "thin", color: "#7F7F7F" }, bRight: { style: "thin", color: "#7F7F7F" } } },
  { id: "linked", label: "Bağlı Hücre", group: "Veri ve Model", style: { color: "#FA7D00", bBottom: { style: "double", color: "#FF8001" } } },
  { id: "note", label: "Not", group: "Veri ve Model", style: { fill: "#FFFFCC", bTop: { style: "thin", color: "#B2B2B2" }, bBottom: { style: "thin", color: "#B2B2B2" }, bLeft: { style: "thin", color: "#B2B2B2" }, bRight: { style: "thin", color: "#B2B2B2" } } },
  { id: "output", label: "Çıkış", group: "Veri ve Model", style: { fill: "#F2F2F2", color: "#3F3F3F", bold: true, bTop: { style: "thin", color: "#3F3F3F" }, bBottom: { style: "thin", color: "#3F3F3F" }, bLeft: { style: "thin", color: "#3F3F3F" }, bRight: { style: "thin", color: "#3F3F3F" } } },
  { id: "warning", label: "Uyarı Metni", group: "Veri ve Model", style: { color: "#FF0000" } },
  { id: "title", label: "Konu Başlığı", group: "Başlıklar", style: { size: 18, color: "#0E2841", font: "Calibri Light" } },
  { id: "h1", label: "Başlık 1", group: "Başlıklar", style: { size: 15, bold: true, color: "#0E2841", bBottom: { style: "thick", color: "#156082" } } },
  { id: "h2", label: "Başlık 2", group: "Başlıklar", style: { size: 13, bold: true, color: "#0E2841", bBottom: { style: "thick", color: "#8AB8D0" } } },
  { id: "h3", label: "Başlık 3", group: "Başlıklar", style: { bold: true, color: "#0E2841", bBottom: { style: "medium", color: "#94C5DE" } } },
  { id: "h4", label: "Başlık 4", group: "Başlıklar", style: { bold: true, color: "#0E2841" } },
  { id: "total", label: "Toplam", group: "Başlıklar", style: { bold: true, bTop: { style: "thin", color: "#156082" }, bBottom: { style: "double", color: "#156082" } } },
  { id: "accent1", label: "Vurgu1", group: "Tema Hücre Stilleri", style: { fill: "#156082", color: "#FFFFFF" } },
  { id: "accent2", label: "Vurgu2", group: "Tema Hücre Stilleri", style: { fill: "#E97132", color: "#FFFFFF" } },
  { id: "accent3", label: "Vurgu3", group: "Tema Hücre Stilleri", style: { fill: "#196B24", color: "#FFFFFF" } },
  { id: "accent4", label: "Vurgu4", group: "Tema Hücre Stilleri", style: { fill: "#0F9ED5", color: "#FFFFFF" } },
  { id: "accent5", label: "Vurgu5", group: "Tema Hücre Stilleri", style: { fill: "#A02B93", color: "#FFFFFF" } },
  { id: "accent6", label: "Vurgu6", group: "Tema Hücre Stilleri", style: { fill: "#4EA72E", color: "#FFFFFF" } },
  { id: "accent1-20", label: "%20 - Vurgu1", group: "Tema Hücre Stilleri", style: { fill: "#C0E4F5" } },
  { id: "accent2-20", label: "%20 - Vurgu2", group: "Tema Hücre Stilleri", style: { fill: "#FBE2D5" } },
  { id: "accent3-20", label: "%20 - Vurgu3", group: "Tema Hücre Stilleri", style: { fill: "#C1F0C8" } },
  { id: "accent4-20", label: "%20 - Vurgu4", group: "Tema Hücre Stilleri", style: { fill: "#CAEDFB" } },
  { id: "accent5-20", label: "%20 - Vurgu5", group: "Tema Hücre Stilleri", style: { fill: "#F2CEEF" } },
  { id: "accent6-20", label: "%20 - Vurgu6", group: "Tema Hücre Stilleri", style: { fill: "#DAF2D0" } },
  { id: "comma", label: "Virgül", group: "Sayı Biçimi", style: { numFmt: '_-* #,##0.00 _₺_-;-* #,##0.00 _₺_-;_-* "-"?? _₺_-;_-@_-' } },
  { id: "currency", label: "Para Birimi", group: "Sayı Biçimi", style: { numFmt: '_-* #,##0.00 "₺"_-;-* #,##0.00 "₺"_-;_-* "-"?? "₺"_-;_-@_-' } },
  { id: "percent", label: "Yüzde", group: "Sayı Biçimi", style: { numFmt: "0%" } },
];

/** Sayı biçimi açılan kutusu seçenekleri (Giriş > Sayı). */
export const NUMBER_FORMATS: { id: string; label: string; code: string }[] = [
  { id: "general", label: "Genel", code: "General" },
  { id: "number", label: "Sayı", code: "0.00" },
  { id: "currency", label: "Para Birimi", code: '#,##0.00 "₺"' },
  { id: "accounting", label: "Finansal", code: '_-* #,##0.00 "₺"_-;-* #,##0.00 "₺"_-;_-* "-"?? "₺"_-;_-@_-' },
  { id: "shortDate", label: "Kısa Tarih", code: "dd.mm.yyyy" },
  { id: "longDate", label: "Uzun Tarih", code: "d mmmm yyyy dddd" },
  { id: "time", label: "Saat", code: "hh:mm:ss" },
  { id: "percent", label: "Yüzde", code: "0.00%" },
  { id: "fraction", label: "Kesir", code: "# ?/?" },
  { id: "scientific", label: "Bilimsel", code: "0.00E+00" },
  { id: "text", label: "Metin", code: "@" },
];

export const ACCOUNTING_FORMAT = NUMBER_FORMATS[3].code;
export const COMMA_FORMAT = '_-* #,##0.00 _₺_-;-* #,##0.00 _₺_-;_-* "-"?? _₺_-;_-@_-';
