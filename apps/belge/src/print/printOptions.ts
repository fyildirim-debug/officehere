// Yazdırma ayarları (Dosya › Yazdır): kopya, sayfa aralığı, tek/elle iki taraflı, harmanlama ve yaprak
// başına sayfa. Saf işlevler: sayfa aralığı ayrıştırma ve yazdırılacak yaprak dizisi.

export type PrintRange = "all" | "current" | "custom";
export type Duplex = "single" | "manual";

export interface PrintOptions {
  copies: number;
  range: PrintRange;
  /** Özel aralık metni ("1-3, 5"). */
  pages: string;
  duplex: Duplex;
  /** Harmanlanmış (1,2,3 1,2,3) ya da harmanlanmamış (1,1,1 2,2,2). */
  collate: boolean;
  /** Yaprak başına sayfa. */
  perSheet: number;
}

export const PER_SHEET = [1, 2, 4, 6, 8, 16] as const;

export function defaultPrintOptions(): PrintOptions {
  return { copies: 1, range: "all", pages: "", duplex: "single", collate: true, perSheet: 1 };
}

/**
 * "1-3, 5" biçimindeki sayfa aralığını 0 tabanlı sayfa sıralarına çevirir (belge sırasıyla, yinelemesiz).
 * Geçersizse hata iletisi döner.
 */
export function parsePageRange(text: string, count: number): number[] | string {
  const parts = text
    .split(/[,;]/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!parts.length) return "Yazdırılacak sayfaları girin (ör. 1-3, 5).";
  const out = new Set<number>();
  for (const part of parts) {
    const m = /^(\d+)\s*(?:-\s*(\d*))?$/.exec(part);
    if (!m) return `Geçersiz sayfa aralığı: "${part}"`;
    const a = Number(m[1]);
    const b = m[2] === undefined ? a : m[2] === "" ? count : Number(m[2]);
    if (a < 1 || b < a) return `Geçersiz sayfa aralığı: "${part}"`;
    for (let p = a; p <= Math.min(b, count); p++) out.add(p - 1);
  }
  if (!out.size) return `Belgede ${count} sayfa var; aralık belgenin dışında.`;
  return [...out].sort((x, y) => x - y);
}

/** Yazdırılacak sayfalar (0 tabanlı), kopyalar hariç. */
export function selectedPages(o: PrintOptions, count: number, current: number): number[] | string {
  if (o.range === "current") return [Math.max(0, Math.min(count - 1, current))];
  if (o.range === "custom") return parsePageRange(o.pages, count);
  return Array.from({ length: count }, (_, i) => i);
}

/**
 * Yazdırma yaprakları: her yaprak 1..perSheet sayfa içerir. Kopyalar harmanlamaya göre dizilir.
 * Elle iki taraflı yazdırmada `pass` 1 tek, 2 çift sıradaki yaprakları verir (arka yüzler ters sırada).
 */
export function printSheets(pages: number[], o: PrintOptions, pass?: 1 | 2): number[][] {
  const n = Math.max(1, o.perSheet);
  const sheetsOf = (list: number[]) => {
    const out: number[][] = [];
    for (let i = 0; i < list.length; i += n) out.push(list.slice(i, i + n));
    return out;
  };
  const one = sheetsOf(pages);
  const copies = Math.max(1, Math.min(999, Math.floor(o.copies) || 1));
  let all: number[][] = [];
  if (o.collate || n > 1) for (let c = 0; c < copies; c++) all.push(...one);
  else for (const s of one) for (let c = 0; c < copies; c++) all.push(s);
  if (o.duplex === "manual" && pass) {
    // Ön yüzler: 1., 3., 5. yapraklar; arka yüzler: 2., 4., 6. (kağıt destesi ters çevrildiği için sondan başa).
    const front = all.filter((_, i) => i % 2 === 0);
    const back = all.filter((_, i) => i % 2 === 1).reverse();
    all = pass === 1 ? front : back;
  }
  return all;
}

/** Yaprak ızgarası (sütun × satır); yazdırmada kağıt yönü sayfaların en büyük görüneceği biçimde seçilir. */
export function sheetGrid(perSheet: number): { cols: number; rows: number } {
  switch (perSheet) {
    case 2:
      return { cols: 2, rows: 1 };
    case 4:
      return { cols: 2, rows: 2 };
    case 6:
      return { cols: 2, rows: 3 };
    case 8:
      return { cols: 4, rows: 2 };
    case 16:
      return { cols: 4, rows: 4 };
    default:
      return { cols: 1, rows: 1 };
  }
}
