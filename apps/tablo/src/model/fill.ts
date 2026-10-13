/**
 * Doldurma tutamacı serileri: sayı ve tarih artışları, Türkçe gün/ay
 * adları, "Ürün 1" gibi metin+sayı, formüllerin göreli kaydırılması.
 */
import { DAYS_TR, DAYS_TR_SHORT, MONTHS_TR, MONTHS_TR_SHORT, MONTHS_EN, DAYS_EN, isDateFormat, serialToDate, dateToSerial, daysInMonth, shiftFormula, type Scalar } from "@officehere/formula";
import type { CellInput } from "./workbook";

const MON_FIRST_TR = [...DAYS_TR.slice(1), DAYS_TR[0]];
const MON_FIRST_TR_SHORT = [...DAYS_TR_SHORT.slice(1), DAYS_TR_SHORT[0]];
const MON_FIRST_EN = [...DAYS_EN.slice(1), DAYS_EN[0]];

export const CUSTOM_LISTS: string[][] = [
  MON_FIRST_TR,
  MON_FIRST_TR_SHORT,
  MONTHS_TR,
  MONTHS_TR_SHORT,
  MON_FIRST_EN,
  MON_FIRST_EN.map((d) => d.slice(0, 3)),
  MONTHS_EN,
  MONTHS_EN.map((m) => m.slice(0, 3)),
];

function findList(s: string): { list: string[]; index: number } | null {
  const u = s.trim().toLocaleUpperCase("tr-TR");
  for (const list of CUSTOM_LISTS) {
    const i = list.findIndex((x) => x.toLocaleUpperCase("tr-TR") === u);
    if (i >= 0) return { list, index: i };
  }
  return null;
}

function matchCase(template: string, word: string): string {
  if (template === template.toLocaleUpperCase("tr-TR") && template !== template.toLocaleLowerCase("tr-TR")) return word.toLocaleUpperCase("tr-TR");
  if (template === template.toLocaleLowerCase("tr-TR")) return word.toLocaleLowerCase("tr-TR");
  return word;
}

/** Doğrusal eğilim (en küçük kareler): y = a + b·x, x = 0..n-1. */
function linearFit(ys: number[]): { a: number; b: number } {
  const n = ys.length;
  if (n === 1) return { a: ys[0], b: 0 };
  let sx = 0;
  let sy = 0;
  let sxx = 0;
  let sxy = 0;
  ys.forEach((y, x) => {
    sx += x;
    sy += y;
    sxx += x * x;
    sxy += x * y;
  });
  const b = (n * sxy - sx * sy) / (n * sxx - sx * sx);
  return { a: (sy - b * sx) / n, b };
}

function clean(n: number): number {
  return Number(n.toPrecision(15));
}

export interface FillSource {
  input: CellInput;
  numFmt?: string;
}

/**
 * Bir doğrultudaki kaynak değerlerden yeni değerler üretir.
 * index: üretilecek konum (kaynak sonrası 0, 1, …); geriye doğru doldurmada negatif adım.
 * formulaOffset(i): i. hedefin kopyalandığı kaynağa göre (satır, sütun) uzaklığı.
 */
export function generateFill(src: FillSource[], count: number, reverse: boolean, copyOnly: boolean, offset: (target: number, srcIndex: number) => [number, number], increment = false): CellInput[] {
  const n = src.length;
  const out: CellInput[] = [];
  const inputs = src.map((s) => s.input);
  const allNumbers = inputs.every((v) => typeof v === "number");
  const isFormula = (v: CellInput): v is { formula: string } => v !== null && typeof v === "object" && "formula" in v;

  // Ctrl ile sürüklenen tek sayı: birer artar (yukarı/sola doğru birer azalır)
  if (increment && n === 1 && allNumbers) {
    const v = inputs[0] as number;
    for (let i = 0; i < count; i++) out.push(clean(v + (reverse ? -(i + 1) : i + 1)));
    return out;
  }

  // Tek sayı: kopyala (Excel davranışı); tarihse gün artır
  if (!copyOnly && allNumbers) {
    const nums = inputs as number[];
    const dateFmt = src[0].numFmt && isDateFormat(src[0].numFmt);
    if (dateFmt) {
      const step = dateStep(nums);
      for (let i = 0; i < count; i++) {
        const k = reverse ? -(i + 1) : n + i;
        out.push(applyDateStep(reverse ? nums[0] : nums[0], step, reverse ? k : k));
      }
      return out;
    }
    if (n >= 2) {
      const { a, b } = linearFit(nums);
      for (let i = 0; i < count; i++) {
        const x = reverse ? -(i + 1) : n + i;
        out.push(clean(a + b * x));
      }
      return out;
    }
  }

  // Metin + sayı ya da liste serileri
  if (!copyOnly && inputs.every((v) => typeof v === "string")) {
    const strs = inputs as string[];
    const lists = strs.map(findList);
    if (lists.every((l) => l && l.list === lists[0]!.list)) {
      const list = lists[0]!.list;
      const idx = lists.map((l) => l!.index);
      const step = n >= 2 ? idx[n - 1] - idx[n - 2] || 1 : 1;
      for (let i = 0; i < count; i++) {
        const base = reverse ? idx[0] - step * (i + 1) : idx[n - 1] + step * (i + 1);
        const word = list[((base % list.length) + list.length) % list.length];
        out.push(matchCase(strs[reverse ? 0 : n - 1], word));
      }
      return out;
    }
    const parts = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s));
    if (parts.every((p) => p && p[1] === parts[0]![1] && p[3] === parts[0]![3])) {
      const nums = parts.map((p) => parseInt(p![2], 10));
      const width = parts[0]![2].length;
      const step = n >= 2 ? nums[n - 1] - nums[n - 2] : 1;
      for (let i = 0; i < count; i++) {
        const v = reverse ? nums[0] - step * (i + 1) : nums[n - 1] + step * (i + 1);
        const abs = Math.abs(v);
        const numStr = parts[0]![2].startsWith("0") ? String(abs).padStart(width, "0") : String(abs);
        out.push(`${parts[0]![1]}${v < 0 ? "-" : ""}${numStr}${parts[0]![3]}`);
      }
      return out;
    }
  }

  // Desen tekrarı (formüller kaydırılır)
  for (let i = 0; i < count; i++) {
    const si = reverse ? (n - 1 - (i % n)) : i % n;
    const v = inputs[si];
    if (isFormula(v)) {
      const [dr, dc] = offset(i, si);
      out.push({ formula: shiftFormula(v.formula, dr, dc) });
    } else if (!copyOnly && typeof v === "number" && n > 1 && !allNumbers) {
      out.push(v);
    } else out.push(v as Scalar);
  }
  return out;
}

type DateStep = { unit: "day" | "month" | "year"; n: number };

function dateStep(nums: number[]): DateStep {
  if (nums.length < 2) return { unit: "day", n: 1 };
  const a = serialToDate(nums[nums.length - 2]);
  const b = serialToDate(nums[nums.length - 1]);
  if (a.d === b.d && (a.m !== b.m || a.y !== b.y)) {
    const months = (b.y - a.y) * 12 + (b.m - a.m);
    if (months % 12 === 0) return { unit: "year", n: months / 12 };
    return { unit: "month", n: months };
  }
  return { unit: "day", n: nums[nums.length - 1] - nums[nums.length - 2] };
}

function applyDateStep(base: number, step: DateStep, k: number): number {
  // base = ilk kaynak; k = ilkten itibaren adım sayısı
  if (step.unit === "day") return base + step.n * k;
  const p = serialToDate(base);
  const months = step.unit === "month" ? step.n * k : step.n * 12 * k;
  let m = p.m - 1 + months;
  const y = p.y + Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return dateToSerial(y, m + 1, Math.min(p.d, daysInMonth(y, m + 1))) + (base - Math.floor(base));
}

/** Seri numarasına tarih birimi ekler. */
export function addDate(serial: number, unit: "day" | "weekday" | "month" | "year", n: number): number {
  if (unit === "day") return serial + n;
  if (unit === "weekday") {
    let s = serial;
    let left = Math.abs(n);
    const d = n >= 0 ? 1 : -1;
    while (left > 0) {
      s += d;
      const wd = (((Math.floor(s) - 1) % 7) + 7) % 7; // 0=Pazar (1900 sistemi)
      if (wd !== 0 && wd !== 6) left--;
    }
    return s;
  }
  const p = serialToDate(serial);
  const tm = unit === "month" ? p.m - 1 + n : p.m - 1;
  const ty = unit === "year" ? p.y + n : p.y + Math.floor(tm / 12);
  const mm = ((tm % 12) + 12) % 12;
  return dateToSerial(ty, mm + 1, Math.min(p.d, daysInMonth(ty, mm + 1))) + (serial - Math.floor(serial));
}
