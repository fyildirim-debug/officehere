/**
 * Yazdırma düzeni (saf hesap): yazdırılacak slaytlar, kâğıt boyutu ve sayfalar —
 * tam sayfa slaytlar, not sayfaları, ana hat ve dinleyici notları (1, 2, 3, 4, 6, 9).
 * Ölçüler 96 dpi pikseldir; not/dinleyici/ana hat sayfaları A4'tür.
 */
import type { PdfExportOptions } from "@officehere/core";
import { bodyShape, paraText, titleShape } from "../model/outline";
import type { Presentation, PrintSettings, PrintWhat } from "../model/types";

export type PrintScope = { kind: "all" } | { kind: "current" } | { kind: "selection" } | { kind: "custom"; text: string };

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Dinleyici notu hücresi: slayt kutusu ve (3'lü düzende) not satırları alanı. */
export interface PrintCell {
  index: number;
  slide: Box;
  lines?: Box;
}

export interface OutlineBlock {
  index: number;
  title: string;
  lines: { level: number; text: string }[];
}

export type PrintPageDesc =
  | { kind: "slide"; index: number }
  | { kind: "notes"; index: number }
  | { kind: "handout"; cells: PrintCell[] }
  | { kind: "outline"; blocks: OutlineBlock[] };

/** A4 (21 × 29,7 cm) 96 dpi piksel. */
export const A4 = { w: (21 / 2.54) * 96, h: (29.7 / 2.54) * 96 };

/** Sayfa kenar boşlukları: üst/alt bilgi şeridi üstte ve altta. */
export const PAGE_MARGIN = { x: 48, top: 64, bottom: 64 };

export const HANDOUT_COUNTS: Record<string, number> = { handouts1: 1, handouts2: 2, handouts3: 3, handouts4: 4, handouts6: 6, handouts9: 9 };

/**
 * "1-3, 5; 8" biçimindeki slayt aralığı → 0 tabanlı, sıralı ve tekil dizinler.
 * Sınır dışı numaralar atılır; tanınmayan parça varsa null.
 */
export function parseSlideRange(text: string, count: number): number[] | null {
  const parts = text
    .split(/[,;]/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!parts.length) return null;
  const out = new Set<number>();
  for (const p of parts) {
    const m = /^(\d+)\s*(?:[-–]\s*(\d+))?$/.exec(p);
    if (!m) return null;
    let a = parseInt(m[1], 10);
    let b = m[2] ? parseInt(m[2], 10) : a;
    if (a > b) [a, b] = [b, a];
    for (let i = Math.max(1, a); i <= Math.min(count, b); i++) out.add(i - 1);
  }
  return [...out].sort((x, y) => x - y);
}

/** Yazdırılacak slayt dizinleri; gizli slaytlar ayar açık değilse atlanır (geçerli slayt dışında). */
export function printSlideIndices(pres: Presentation, settings: PrintSettings, scope: PrintScope, current: number, selection: number[]): number[] {
  const n = pres.slides.length;
  let idx: number[];
  switch (scope.kind) {
    case "current":
      return current >= 0 && current < n ? [current] : [];
    case "selection":
      idx = [...new Set(selection)].filter((i) => i >= 0 && i < n).sort((a, b) => a - b);
      break;
    case "custom":
      idx = parseSlideRange(scope.text, n) ?? [];
      break;
    default:
      idx = pres.slides.map((_, i) => i);
  }
  return settings.hiddenSlides ? idx : idx.filter((i) => !pres.slides[i].hidden);
}

/** Kâğıt boyutu: tam sayfa slaytlarda slayt boyutu, diğerlerinde A4 (yönlendirmeye göre). */
export function paperSize(pres: Presentation, what: PrintWhat, orientation: PrintSettings["orientation"]): { w: number; h: number } {
  if (what === "slides") return { w: pres.width, h: pres.height };
  return orientation === "landscape" ? { w: A4.h, h: A4.w } : { ...A4 };
}

/** Kutuya oranı korunarak sığdırılmış, ortalanmış slayt kutusu. */
function fitSlide(cell: Box, ratio: number): Box {
  let w = cell.w;
  let h = w / ratio;
  if (h > cell.h) {
    h = cell.h;
    w = h * ratio;
  }
  return { x: cell.x + (cell.w - w) / 2, y: cell.y + (cell.h - h) / 2, w, h };
}

/** Dinleyici notu sayfasındaki hücre yerleşimi (sıra: yatay satır satır, dikey sütun sütun). */
export function handoutSlots(per: number, page: { w: number; h: number }, ratio: number, order: "horz" | "vert" = "horz"): { slide: Box; lines?: Box }[] {
  const portrait = page.h >= page.w;
  const area: Box = { x: PAGE_MARGIN.x, y: PAGE_MARGIN.top, w: page.w - PAGE_MARGIN.x * 2, h: page.h - PAGE_MARGIN.top - PAGE_MARGIN.bottom };
  const gap = 18;
  // 3'lü düzen: her slaydın yanında (dikey) ya da altında (yatay) not satırları.
  if (per === 3) {
    const slots: { slide: Box; lines?: Box }[] = [];
    for (let i = 0; i < 3; i++) {
      if (portrait) {
        const rowH = (area.h - gap * 2) / 3;
        const y = area.y + i * (rowH + gap);
        const half = (area.w - gap) / 2;
        const sl = fitSlide({ x: area.x, y, w: half, h: rowH }, ratio);
        // Not satırları slaydın yanında, slayt yüksekliğince.
        slots.push({ slide: sl, lines: { x: area.x + half + gap, y: sl.y, w: half, h: sl.h } });
      } else {
        const colW = (area.w - gap * 2) / 3;
        const x = area.x + i * (colW + gap);
        const s = fitSlide({ x, y: area.y, w: colW, h: area.h / 2 }, ratio);
        const sTop = { ...s, y: area.y };
        slots.push({ slide: sTop, lines: { x, y: sTop.y + sTop.h + gap, w: colW, h: area.y + area.h - (sTop.y + sTop.h + gap) } });
      }
    }
    return slots;
  }
  const grid: Record<number, [number, number]> = portrait ? { 1: [1, 1], 2: [1, 2], 4: [2, 2], 6: [2, 3], 9: [3, 3] } : { 1: [1, 1], 2: [2, 1], 4: [2, 2], 6: [3, 2], 9: [3, 3] };
  const [cols, rows] = grid[per] ?? [1, 1];
  const cw = (area.w - gap * (cols - 1)) / cols;
  const ch = (area.h - gap * (rows - 1)) / rows;
  const slots: { slide: Box }[] = [];
  for (let k = 0; k < per; k++) {
    const r = order === "vert" ? k % rows : Math.floor(k / cols);
    const c = order === "vert" ? Math.floor(k / rows) : k % cols;
    slots.push({ slide: fitSlide({ x: area.x + c * (cw + gap), y: area.y + r * (ch + gap), w: cw, h: ch }, ratio) });
  }
  return slots;
}

/** Not sayfası: üstte slayt, altta notlar alanı. */
export function notesPageBoxes(page: { w: number; h: number }, ratio: number): { slide: Box; notes: Box } {
  const area: Box = { x: PAGE_MARGIN.x + 24, y: PAGE_MARGIN.top, w: page.w - (PAGE_MARGIN.x + 24) * 2, h: page.h - PAGE_MARGIN.top - PAGE_MARGIN.bottom };
  const slide = fitSlide({ x: area.x, y: area.y + 8, w: area.w, h: area.h * 0.45 }, ratio);
  const top = slide.y + slide.h + 28;
  return { slide, notes: { x: area.x, y: top, w: area.w, h: area.y + area.h - top } };
}

/** Ana hat blokları (başlık + gövde yer tutucusu paragrafları). */
export function outlineBlocks(pres: Presentation, indices: number[]): OutlineBlock[] {
  return indices.map((i) => {
    const s = pres.slides[i];
    const title = titleShape(s)?.text?.paragraphs.map(paraText).join(" ").replace(/\v/g, " ").trim() ?? "";
    const lines = (bodyShape(s)?.text?.paragraphs ?? []).map((p) => ({ level: p.level, text: paraText(p).replace(/\v/g, " ").trim() })).filter((l) => l.text);
    return { index: i, title, lines };
  });
}

/** 3'lü dinleyici notunda not satırı aralığı (px). */
export const NOTE_LINE_GAP = 28;

/** Ana hat satır yüksekliği ve sayfa başına satır tahmini (karakter genişliği ~7,5 px). */
const OUTLINE_LINE = 20;
function blockHeight(b: OutlineBlock, width: number): number {
  const wrap = (t: string, indent: number) => Math.max(1, Math.ceil((t.length * 7.5) / Math.max(80, width - indent)));
  return 10 + wrap(b.title || " ", 28) * (OUTLINE_LINE + 4) + b.lines.reduce((n, l) => n + wrap(l.text, 28 + 24 * (l.level + 1)) * OUTLINE_LINE, 0);
}

/** Yazdırılacak sayfalar. */
export function buildPrintPages(pres: Presentation, settings: PrintSettings, indices: number[]): PrintPageDesc[] {
  const what = settings.what;
  if (what === "slides") return indices.map((index) => ({ kind: "slide", index }));
  if (what === "notes") return indices.map((index) => ({ kind: "notes", index }));
  const page = paperSize(pres, what, settings.orientation);
  if (what === "outline") {
    const avail = page.h - PAGE_MARGIN.top - PAGE_MARGIN.bottom;
    const width = page.w - PAGE_MARGIN.x * 2;
    const pages: PrintPageDesc[] = [];
    let cur: OutlineBlock[] = [];
    let used = 0;
    for (const b of outlineBlocks(pres, indices)) {
      const h = blockHeight(b, width);
      if (cur.length && used + h > avail) {
        pages.push({ kind: "outline", blocks: cur });
        cur = [];
        used = 0;
      }
      cur.push(b);
      used += h;
    }
    if (cur.length) pages.push({ kind: "outline", blocks: cur });
    return pages;
  }
  const per = HANDOUT_COUNTS[what] ?? 6;
  const slots = handoutSlots(per, page, pres.width / pres.height, settings.order);
  const pages: PrintPageDesc[] = [];
  for (let k = 0; k < indices.length; k += per) {
    pages.push({ kind: "handout", cells: indices.slice(k, k + per).map((index, j) => ({ index, ...slots[j] })) });
  }
  return pages;
}

/** Yazdırma/PDF işi: ayarlar ve yazdırılacak slaytlar. pdf: masaüstünde doğrudan PDF'e yazılır. */
export interface PrintJob {
  settings: PrintSettings;
  indices: number[];
  pdf: boolean;
  /** Masaüstü PDF seçenekleri (iletişim kutusundan). */
  pdfOptions?: PdfExportOptions;
}
