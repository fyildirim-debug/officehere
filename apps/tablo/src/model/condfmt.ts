/**
 * Koşullu biçimlendirme değerlendirmesi. Kural istatistikleri (ortalama,
 * ilk 10, yinelenenler, min/maks) her hesap sürümünde bir kez çıkarılır.
 */
import { cellKey, isError, jsDateToSerial, shiftFormula, toBool, weekdayOf, type Scalar } from "@officehere/formula";
import { contains } from "./rect";
import type { CellStyle } from "./styles";
import type { CondRule } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

export interface CondResult {
  style?: Partial<CellStyle>;
  bar?: { from: number; to: number; color: string };
  scale?: string;
}

interface RuleStats {
  nums: number[];
  min: number;
  max: number;
  avg: number;
  threshold?: number;
  counts?: Map<string, number>;
  mid?: number;
}

/** Hazır vurgu stilleri (Hücre Vurgulama Kuralları iletişim kutusu). */
export const HIGHLIGHT_PRESETS: { id: string; label: string; style: Partial<CellStyle> }[] = [
  { id: "lightRed", label: "Koyu Kırmızı Metinle Açık Kırmızı Dolgu", style: { fill: "#FFC7CE", color: "#9C0006" } },
  { id: "yellow", label: "Koyu Sarı Metinle Sarı Dolgu", style: { fill: "#FFEB9C", color: "#9C5700" } },
  { id: "green", label: "Koyu Yeşil Metinle Yeşil Dolgu", style: { fill: "#C6EFCE", color: "#006100" } },
  { id: "redFill", label: "Açık Kırmızı Dolgu", style: { fill: "#FFC7CE" } },
  { id: "redText", label: "Kırmızı Metin", style: { color: "#9C0006" } },
  { id: "redBorder", label: "Kırmızı Kenarlık", style: { bTop: { style: "thin", color: "#C00000" }, bBottom: { style: "thin", color: "#C00000" }, bLeft: { style: "thin", color: "#C00000" }, bRight: { style: "thin", color: "#C00000" } } },
];

export const DATA_BAR_COLORS = ["#638EC6", "#63C384", "#FF555A", "#FFB628", "#008AEF", "#D6007B"];
export const COLOR_SCALES: { id: string; label: string; colors: string[] }[] = [
  { id: "gyr", label: "Yeşil - Sarı - Kırmızı Renk Ölçeği", colors: ["#F8696B", "#FFEB84", "#63BE7B"] },
  { id: "ryg", label: "Kırmızı - Sarı - Yeşil Renk Ölçeği", colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  { id: "gwr", label: "Yeşil - Beyaz - Kırmızı Renk Ölçeği", colors: ["#F8696B", "#FCFCFF", "#63BE7B"] },
  { id: "rwg", label: "Kırmızı - Beyaz - Yeşil Renk Ölçeği", colors: ["#63BE7B", "#FCFCFF", "#F8696B"] },
  { id: "bwr", label: "Mavi - Beyaz - Kırmızı Renk Ölçeği", colors: ["#F8696B", "#FCFCFF", "#5A8AC6"] },
  { id: "rwb", label: "Kırmızı - Beyaz - Mavi Renk Ölçeği", colors: ["#5A8AC6", "#FCFCFF", "#F8696B"] },
  { id: "wr", label: "Beyaz - Kırmızı Renk Ölçeği", colors: ["#FCFCFF", "#F8696B"] },
  { id: "gw", label: "Yeşil - Beyaz Renk Ölçeği", colors: ["#63BE7B", "#FCFCFF"] },
  { id: "wg", label: "Beyaz - Yeşil Renk Ölçeği", colors: ["#FCFCFF", "#63BE7B"] },
  { id: "gy", label: "Yeşil - Sarı Renk Ölçeği", colors: ["#63BE7B", "#FFEF9C"] },
];

function keyOf(v: Scalar): string {
  if (typeof v === "string") return "s" + v.toLocaleUpperCase("tr-TR");
  return typeof v + String(v);
}

function lerpColor(a: string, b: string, t: number): string {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  const c = pa.map((x, i) => Math.round(x + (pb[i] - x) * t));
  return `#${c.map((x) => x.toString(16).padStart(2, "0")).join("")}`;
}

export class CondFormatter {
  private stats = new Map<string, RuleStats>();
  private version = -1;

  constructor(private wb: WorkbookModel) {}

  private statsFor(sheet: SheetModel, rule: CondRule): RuleStats {
    const v = this.wb.engine.version;
    if (v !== this.version) {
      this.stats.clear();
      this.version = v;
    }
    let st = this.stats.get(rule.id);
    if (st) return st;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    const ext = this.wb.dataExtent(sheet);
    for (const r of rule.ranges) {
      const r2 = Math.min(r.r2, ext.rows - 1);
      const c2 = Math.min(r.c2, ext.cols - 1);
      for (let i = r.r1; i <= r2; i++)
        for (let j = r.c1; j <= c2; j++) {
          const val = this.wb.getValue(sheet, i, j);
          if (typeof val === "number") nums.push(val);
          if (val !== null && (rule.type === "duplicateValues" || rule.type === "uniqueValues")) {
            const k = keyOf(val);
            counts.set(k, (counts.get(k) ?? 0) + 1);
          }
        }
    }
    const sorted = [...nums].sort((a, b) => a - b);
    const min = sorted[0] ?? 0;
    const max = sorted[sorted.length - 1] ?? 0;
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    st = { nums: sorted, min, max, avg, counts };
    if (rule.type === "top10") {
      const n = rule.percent ? Math.max(1, Math.floor((sorted.length * (rule.rank ?? 10)) / 100)) : rule.rank ?? 10;
      const list = rule.bottom ? sorted : [...sorted].reverse();
      st.threshold = list[Math.min(n, list.length) - 1];
    }
    if (rule.type === "colorScale") st.mid = sorted.length ? sorted[Math.floor((sorted.length - 1) / 2)] * 0.5 + sorted[Math.ceil((sorted.length - 1) / 2)] * 0.5 : 0;
    this.stats.set(rule.id, st);
    return st;
  }

  private formulaValue(sheet: SheetModel, rule: CondRule, idx: number, row: number, col: number): Scalar {
    const raw = rule.values?.[idx] ?? "";
    if (!raw.startsWith("=")) {
      const n = Number(raw.replace(",", "."));
      return raw.trim() !== "" && Number.isFinite(n) ? n : raw;
    }
    const base = rule.ranges[0];
    const f = shiftFormula(raw, row - base.r1, col - base.c1);
    return this.wb.engine.evaluate(sheet.id, row, col, f);
  }

  /** Hücreye uygulanan koşullu biçim sonuçları. */
  evaluate(sheet: SheetModel, row: number, col: number): CondResult | null {
    if (!sheet.condFormats.length) return null;
    let out: CondResult | null = null;
    for (const rule of sheet.condFormats) {
      if (!rule.ranges.some((r) => contains(r, row, col))) continue;
      const v = this.wb.getValue(sheet, row, col);
      const hit = this.match(sheet, rule, v, row, col);
      if (!hit) continue;
      out ??= {};
      if (hit.bar && !out.bar) out.bar = hit.bar;
      if (hit.scale && !out.scale) out.scale = hit.scale;
      if (hit.style) out.style = { ...hit.style, ...(out.style ?? {}) };
      if (rule.stopIfTrue) break;
    }
    return out;
  }

  private match(sheet: SheetModel, rule: CondRule, v: Scalar, row: number, col: number): CondResult | null {
    const style = rule.style ? { style: rule.style } : {};
    switch (rule.type) {
      case "cellIs": {
        if (v === null && rule.operator !== "equal" && rule.operator !== "notEqual") return null;
        const a = this.formulaValue(sheet, rule, 0, row, col);
        const b = this.formulaValue(sheet, rule, 1, row, col);
        const cmp = (x: Scalar, y: Scalar): number => {
          if (typeof x === "number" && typeof y === "number") return x - y;
          const sx = String(x ?? "").toLocaleUpperCase("tr-TR");
          const sy = String(y ?? "").toLocaleUpperCase("tr-TR");
          return typeof x === "number" ? -1 : typeof y === "number" ? 1 : sx < sy ? -1 : sx > sy ? 1 : 0;
        };
        let ok = false;
        switch (rule.operator) {
          case "greaterThan":
            ok = cmp(v, a) > 0;
            break;
          case "lessThan":
            ok = cmp(v, a) < 0;
            break;
          case "greaterThanOrEqual":
            ok = cmp(v, a) >= 0;
            break;
          case "lessThanOrEqual":
            ok = cmp(v, a) <= 0;
            break;
          case "equal":
            ok = cmp(v, a) === 0;
            break;
          case "notEqual":
            ok = cmp(v, a) !== 0;
            break;
          case "between":
            ok = cmp(v, a) >= 0 && cmp(v, b) <= 0;
            break;
          case "notBetween":
            ok = cmp(v, a) < 0 || cmp(v, b) > 0;
            break;
        }
        return ok ? style : null;
      }
      case "containsText":
      case "notContainsText":
      case "beginsWith":
      case "endsWith": {
        const s = v === null ? "" : String(v).toLocaleUpperCase("tr-TR");
        const t = (rule.text ?? "").toLocaleUpperCase("tr-TR");
        const ok = rule.type === "containsText" ? s.includes(t) : rule.type === "notContainsText" ? !s.includes(t) : rule.type === "beginsWith" ? s.startsWith(t) : s.endsWith(t);
        return ok && (v !== null || rule.type === "notContainsText") ? style : null;
      }
      case "duplicateValues":
      case "uniqueValues": {
        if (v === null) return null;
        const st = this.statsFor(sheet, rule);
        const n = st.counts?.get(keyOf(v)) ?? 0;
        return (rule.type === "duplicateValues" ? n > 1 : n === 1) ? style : null;
      }
      case "top10": {
        if (typeof v !== "number") return null;
        const st = this.statsFor(sheet, rule);
        if (st.threshold === undefined) return null;
        return (rule.bottom ? v <= st.threshold : v >= st.threshold) ? style : null;
      }
      case "aboveAverage": {
        if (typeof v !== "number") return null;
        const st = this.statsFor(sheet, rule);
        return (rule.below ? v < st.avg : v > st.avg) ? style : null;
      }
      case "containsBlanks":
        return v === null || (typeof v === "string" && v.trim() === "") ? style : null;
      case "notContainsBlanks":
        return v !== null && !(typeof v === "string" && v.trim() === "") ? style : null;
      case "containsErrors":
        return isError(v) ? style : null;
      case "expression": {
        const r = this.formulaValue(sheet, rule, 0, row, col);
        const b = toBool(r);
        return b === true ? style : null;
      }
      case "timePeriod": {
        if (typeof v !== "number") return null;
        const today = Math.floor(jsDateToSerial(new Date()));
        const d = Math.floor(v);
        const wd = weekdayOf(today);
        const weekStart = today - wd;
        let ok = false;
        switch (rule.timePeriod) {
          case "yesterday":
            ok = d === today - 1;
            break;
          case "today":
            ok = d === today;
            break;
          case "tomorrow":
            ok = d === today + 1;
            break;
          case "last7Days":
            ok = d > today - 7 && d <= today;
            break;
          case "thisWeek":
            ok = d >= weekStart && d < weekStart + 7;
            break;
          case "lastWeek":
            ok = d >= weekStart - 7 && d < weekStart;
            break;
          case "nextWeek":
            ok = d >= weekStart + 7 && d < weekStart + 14;
            break;
          default: {
            const now = new Date();
            const dt = new Date(Date.UTC(1899, 11, 30) + d * 86400000);
            const mDiff = (dt.getUTCFullYear() - now.getFullYear()) * 12 + dt.getUTCMonth() - now.getMonth();
            ok = rule.timePeriod === "thisMonth" ? mDiff === 0 : rule.timePeriod === "lastMonth" ? mDiff === -1 : rule.timePeriod === "nextMonth" ? mDiff === 1 : false;
          }
        }
        return ok ? style : null;
      }
      case "dataBar": {
        if (typeof v !== "number") return null;
        const st = this.statsFor(sheet, rule);
        const lo = Math.min(0, st.min);
        const hi = Math.max(0, st.max);
        const span = hi - lo || 1;
        const zero = (0 - lo) / span;
        const pos = (v - lo) / span;
        return { bar: { from: Math.min(zero, pos), to: Math.max(zero, pos), color: rule.color ?? DATA_BAR_COLORS[0] } };
      }
      case "colorScale": {
        if (typeof v !== "number") return null;
        const st = this.statsFor(sheet, rule);
        const cols = rule.colors ?? COLOR_SCALES[0].colors;
        if (st.max === st.min) return { scale: cols[cols.length > 2 ? 1 : 0] };
        if (cols.length === 2) return { scale: lerpColor(cols[0], cols[1], (v - st.min) / (st.max - st.min)) };
        const mid = st.mid ?? (st.min + st.max) / 2;
        if (v <= mid) return { scale: lerpColor(cols[0], cols[1], mid === st.min ? 1 : (v - st.min) / (mid - st.min)) };
        return { scale: lerpColor(cols[1], cols[2], st.max === mid ? 1 : (v - mid) / (st.max - mid)) };
      }
    }
    return null;
  }
}

export function ruleKey(row: number, col: number): number {
  return cellKey(row, col);
}
