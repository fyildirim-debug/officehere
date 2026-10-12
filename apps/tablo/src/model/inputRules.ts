/** Hücre girişi kuralları: veri doğrulama, parantez tamamlama, gevşek başvuru ayrıştırma, doğrulama listeleri. */
import { MAX_COLS, MAX_ROWS } from "@officehere/formula";
import { rect } from "./rect";
import type { Rect } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

/** Veri doğrulama kuralı ihlali (Hata Uyarısı "Dur" ise ileti). */
export function validateInputRule(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, text: string): { title: string; message: string } | null {
  const v = sheet.validations.find((x) => x.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2));
  if (!v || text === "" || text.startsWith("=")) return null;
  const rule = v.rule;
  if (rule.type === "any" || (rule.errorStyle && rule.errorStyle !== "stop")) return null;
  const fail = { title: rule.errorTitle || "OfficeHere Tablo", message: rule.error || "Bu değer, bu hücre için tanımlanan veri doğrulama kısıtlamalarıyla eşleşmiyor." };
  if (rule.type === "list") {
    const items = validationList(wb, sheet, v.rule.formula1 ?? "", row, col);
    return items.some((i) => i.toLocaleUpperCase("tr-TR") === text.trim().toLocaleUpperCase("tr-TR")) ? null : fail;
  }
  const n = Number(text.replace(/\./g, "").replace(",", "."));
  const evalF = (f?: string) => {
    if (!f) return NaN;
    if (f.startsWith("=")) {
      const r = wb.engine.evaluate(sheet.id, row, col, f);
      return typeof r === "number" ? r : NaN;
    }
    return Number(f.replace(",", "."));
  };
  let x = n;
  if (rule.type === "textLength") x = text.length;
  if (rule.type === "date") {
    const p = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(text.trim());
    x = p ? Math.round((Date.UTC(+p[3], +p[2] - 1, +p[1]) - Date.UTC(1899, 11, 30)) / 86400000) : NaN;
  }
  if (Number.isNaN(x)) return fail;
  if (rule.type === "whole" && !Number.isInteger(x)) return fail;
  if (rule.type === "custom") {
    const r = wb.engine.evaluate(sheet.id, row, col, rule.formula1 ?? "=TRUE");
    return r === true ? null : fail;
  }
  const a = evalF(rule.formula1);
  const b = evalF(rule.formula2);
  let ok = true;
  switch (rule.operator ?? "between") {
    case "between":
      ok = x >= a && x <= b;
      break;
    case "notBetween":
      ok = x < a || x > b;
      break;
    case "equal":
      ok = x === a;
      break;
    case "notEqual":
      ok = x !== a;
      break;
    case "greaterThan":
      ok = x > a;
      break;
    case "lessThan":
      ok = x < a;
      break;
    case "greaterThanOrEqual":
      ok = x >= a;
      break;
    case "lessThanOrEqual":
      ok = x <= a;
      break;
  }
  return ok ? null : fail;
}

/** Kapanmamış parantezleri tamamlar (Excel davranışı). */
export function autoCloseParens(text: string): string {
  let depth = 0;
  let inStr = false;
  for (const ch of text) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === "(") depth++;
    else if (!inStr && ch === ")") depth--;
  }
  return depth > 0 && !inStr ? text + ")".repeat(depth) : text;
}

/** "A1", "a1:b5", "A:A", "3:5" gibi başvuruları gevşek ayrıştırır. */
export function parseRangeLoose(text: string): Rect | null {
  const t = text.trim().replace(/\$/g, "").toUpperCase();
  const m1 = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(t);
  const col = (s: string) => s.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  if (m1) {
    const r1 = +m1[2] - 1;
    const c1 = col(m1[1]);
    const r2 = m1[4] ? +m1[4] - 1 : r1;
    const c2 = m1[3] ? col(m1[3]) : c1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS || c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return rect(r1, c1, r2, c2);
  }
  const m2 = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(t);
  if (m2) return rect(0, col(m2[1]), MAX_ROWS - 1, col(m2[2]));
  const m3 = /^(\d{1,7}):(\d{1,7})$/.exec(t);
  if (m3) return rect(+m3[1] - 1, 0, +m3[2] - 1, MAX_COLS - 1);
  return null;
}

/** Veri doğrulama listesi öğeleri ("a;b;c" ya da "=$A$1:$A$5"). */
export function validationList(wb: WorkbookModel, sheet: SheetModel, formula: string, row: number, col: number): string[] {
  if (!formula) return [];
  if (formula.startsWith("=")) {
    const m = wb.engine.evaluateArray(sheet.id, row, col, formula);
    return m.flat().filter((v) => v !== null && v !== "").map((v) => (typeof v === "number" ? String(v).replace(".", ",") : String(v)));
  }
  return formula
    .split(/[;,]/)
    .map((s) => s.trim().replace(/^"|"$/g, ""))
    .filter(Boolean);
}
