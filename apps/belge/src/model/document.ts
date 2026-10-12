// Belge modeli: ProseMirror içeriği + sayfa/bölüm ayarları (OOXML sectPr karşılığı).
import type { JSONContent } from "@tiptap/core";

/** 1 inç = 1440 twip; ekranda 96 px → 15 twip/px. */
export const TWIPS_PER_PX = 15;
export const twToPx = (tw: number) => tw / TWIPS_PER_PX;
export const pxToTw = (px: number) => Math.round(px * TWIPS_PER_PX);
export const cmToTw = (cm: number) => Math.round((cm / 2.54) * 1440);
export const twToCm = (tw: number) => (tw / 1440) * 2.54;
export const ptToPx = (pt: number) => (pt * 4) / 3;
export const pxToPt = (px: number) => (px * 3) / 4;

export interface PageSizeDef {
  id: string;
  label: string;
  /** Dikey yönde genişlik/yükseklik (twip). */
  w: number;
  h: number;
}

export const PAGE_SIZE_LIST: PageSizeDef[] = [
  { id: "A4", label: "A4", w: 11906, h: 16838 },
  { id: "A3", label: "A3", w: 16838, h: 23811 },
  { id: "A5", label: "A5", w: 8391, h: 11906 },
  { id: "Letter", label: "Letter", w: 12240, h: 15840 },
  { id: "Legal", label: "Legal", w: 12240, h: 20160 },
];

export interface Margins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
  gutter: number;
}

export interface MarginPreset {
  id: string;
  label: string;
  m: Pick<Margins, "top" | "bottom" | "left" | "right">;
}

export const MARGIN_PRESETS: MarginPreset[] = [
  { id: "normal", label: "Normal", m: { top: 1417, bottom: 1417, left: 1417, right: 1417 } },
  { id: "narrow", label: "Dar", m: { top: 720, bottom: 720, left: 720, right: 720 } },
  { id: "moderate", label: "Orta", m: { top: 1440, bottom: 1440, left: 1080, right: 1080 } },
  { id: "wide", label: "Geniş", m: { top: 1440, bottom: 1440, left: 2880, right: 2880 } },
];

export interface PageSetup {
  /** Gerçek genişlik/yükseklik (yönlendirme uygulanmış, twip). */
  width: number;
  height: number;
  orientation: "portrait" | "landscape";
  margin: Margins;
  columns: number;
  /** Sütun aralığı (twip). */
  columnSpace: number;
}

export interface CommentData {
  id: string;
  author: string;
  initials: string;
  date: string;
  text: string;
  resolved?: boolean;
}

export interface Watermark {
  text: string;
  color: string;
  layout: "diagonal" | "horizontal";
  font: string;
  semitransparent: boolean;
}

export interface PageBorder {
  style: "single" | "double" | "dotted" | "dashed" | "thick";
  color: string;
  /** Kalınlık (pt). */
  width: number;
}

export interface DocProperties {
  title: string;
  subject: string;
  author: string;
  keywords: string;
  created: string;
  modified: string;
}

export interface DocSettings {
  page: PageSetup;
  header: JSONContent | null;
  footer: JSONContent | null;
  watermark: Watermark | null;
  pageColor: string | null;
  pageBorder: PageBorder | null;
  /** Tasarım > Belge Biçimlendirme stil kümesi. */
  styleSet: string;
  /** Tema adı (core THEMES): renkler ve yazı tipleri. */
  theme: string;
  /** Tasarım > Yazı Tipleri ile seçilen yazı tipi teması (yoksa temanınki). */
  fontTheme: string | null;
  /** Tasarım > Paragraf Aralığı önayarı. */
  paragraphSpacing: string;
  comments: Record<string, CommentData>;
  properties: DocProperties;
}

export interface BelgeDocument {
  content: JSONContent;
  settings: DocSettings;
}

export function defaultPage(): PageSetup {
  return {
    width: 11906,
    height: 16838,
    orientation: "portrait",
    margin: { top: 1417, bottom: 1417, left: 1417, right: 1417, header: 708, footer: 708, gutter: 0 },
    columns: 1,
    columnSpace: 708,
  };
}

export function defaultSettings(): DocSettings {
  const now = new Date().toISOString();
  return {
    page: defaultPage(),
    header: null,
    footer: null,
    watermark: null,
    pageColor: null,
    pageBorder: null,
    styleSet: "default",
    theme: "OfficeHere",
    fontTheme: null,
    paragraphSpacing: "default",
    comments: {},
    properties: { title: "", subject: "", author: "", keywords: "", created: now, modified: now },
  };
}

export function emptyDoc(): JSONContent {
  return { type: "doc", content: [{ type: "paragraph" }] };
}

/** Sayfa boyutunun adını bulur (yönlendirmeden bağımsız). */
export function pageSizeId(p: PageSetup): string {
  const w = Math.min(p.width, p.height);
  const h = Math.max(p.width, p.height);
  const hit = PAGE_SIZE_LIST.find((s) => Math.abs(s.w - w) < 30 && Math.abs(s.h - h) < 30);
  return hit?.id ?? "custom";
}

export function marginPresetId(p: PageSetup): string {
  const m = p.margin;
  const hit = MARGIN_PRESETS.find((x) => x.m.top === m.top && x.m.bottom === m.bottom && x.m.left === m.left && x.m.right === m.right);
  return hit?.id ?? "custom";
}

/** Sayfa boyutları (px). */
export function pageMetrics(p: PageSetup) {
  const w = twToPx(p.width);
  const h = twToPx(p.height);
  const mt = twToPx(p.margin.top);
  const mb = twToPx(p.margin.bottom);
  const ml = twToPx(p.margin.left + p.margin.gutter);
  const mr = twToPx(p.margin.right);
  return { w, h, mt, mb, ml, mr, contentW: w - ml - mr, contentH: h - mt - mb, headerY: twToPx(p.margin.header), footerY: twToPx(p.margin.footer) };
}

/** cm değerini Türkçe biçimde gösterir ("2,5 cm"). */
export function fmtCm(tw: number, digits = 2): string {
  return `${Number(twToCm(tw).toFixed(digits)).toLocaleString("tr-TR")} cm`;
}

export function parseCm(text: string): number | null {
  const n = parseFloat(text.replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : null;
}
