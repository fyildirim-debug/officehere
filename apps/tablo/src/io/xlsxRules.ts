/** Veri doğrulama ve koşullu biçimlendirme kurallarının exceljs modeli ⇄ uygulama modeli dönüşümleri. */
import type { CondRule, Rect, ValidationRule } from "../model/types";
import { fromXStyle, toXStyle, type ColorResolver } from "./xlsxStyles";
import { addXlfnPrefixes, normalizeFileFormula, utcSerial } from "./xlsxFormula";

/* eslint-disable @typescript-eslint/no-explicit-any */

export function readValidation(dv: any): ValidationRule {
  const strip = (f: any) => (f === undefined || f === null ? undefined : f instanceof Date ? String(utcSerial(f)) : String(f));
  let f1 = strip(dv.formulae?.[0]);
  const f2 = strip(dv.formulae?.[1]);
  if (dv.type === "list" && f1) {
    if (f1.startsWith('"')) f1 = f1.slice(1, -1).split(",").join(";");
    else f1 = "=" + f1.replace(/^=/, "");
  } else if (f1 && !/^-?[\d.]+$/.test(f1) && dv.type !== "date") f1 = "=" + f1.replace(/^=/, "");
  return {
    type: dv.type,
    operator: dv.operator,
    formula1: f1,
    formula2: f2,
    allowBlank: dv.allowBlank,
    showDropdown: dv.showDropDown !== true,
    errorTitle: dv.errorTitle,
    error: dv.error,
    errorStyle: dv.errorStyle,
    promptTitle: dv.promptTitle,
    prompt: dv.prompt,
  };
}

/** Tek hücre listesini dikdörtgenlere sıkıştırır (sütun bazında ardışık satırlar). */
export function compressCells(cells: [number, number][]): Rect[] {
  const byCol = new Map<number, number[]>();
  for (const [r, c] of cells) {
    const l = byCol.get(c) ?? [];
    l.push(r);
    byCol.set(c, l);
  }
  const runs: Rect[] = [];
  for (const [c, rows] of byCol) {
    rows.sort((a, b) => a - b);
    let start = rows[0];
    let prev = rows[0];
    for (let i = 1; i <= rows.length; i++) {
      if (i < rows.length && rows[i] === prev + 1) {
        prev = rows[i];
        continue;
      }
      runs.push({ r1: start, c1: c, r2: prev, c2: c });
      start = rows[i];
      prev = rows[i];
    }
  }
  // Aynı satır aralığındaki komşu sütunları birleştir
  runs.sort((a, b) => a.r1 - b.r1 || a.r2 - b.r2 || a.c1 - b.c1);
  const out: Rect[] = [];
  for (const r of runs) {
    const last = out[out.length - 1];
    if (last && last.r1 === r.r1 && last.r2 === r.r2 && last.c2 + 1 === r.c1) last.c2 = r.c2;
    else out.push({ ...r });
  }
  return out;
}

export function readCondRule(rule: any, cr: ColorResolver): Omit<CondRule, "id" | "ranges"> | null {
  const style = rule.style ? fromXStyle(rule.style, cr) : undefined;
  const formulas = (rule.formulae ?? []).map((f: any) => "=" + normalizeFileFormula(String(f)));
  switch (rule.type) {
    case "cellIs":
      return { type: "cellIs", operator: rule.operator, values: formulas, style };
    case "containsText": {
      const op = rule.operator;
      if (op === "containsBlanks") return { type: "containsBlanks", style };
      if (op === "notContainsBlanks") return { type: "notContainsBlanks", style };
      if (op === "containsErrors") return { type: "containsErrors", style };
      return { type: op === "notContains" ? "notContainsText" : op === "beginsWith" ? "beginsWith" : op === "endsWith" ? "endsWith" : "containsText", text: rule.text, style };
    }
    case "top10":
      return { type: "top10", rank: rule.rank, percent: !!rule.percent, bottom: !!rule.bottom, style };
    case "aboveAverage":
      return { type: "aboveAverage", below: rule.aboveAverage === false, style };
    case "duplicateValues":
    case "uniqueValues":
      return { type: rule.type, style };
    case "expression":
      return { type: "expression", values: formulas, style };
    case "timePeriod":
      return { type: "timePeriod", timePeriod: rule.timePeriod, style };
    case "dataBar":
      return { type: "dataBar", color: cr.resolve(rule.color) ?? "#638EC6" };
    case "colorScale":
      return { type: "colorScale", colors: (rule.color ?? []).map((c: any) => cr.resolve(c) ?? "#FFFFFF") };
  }
  return null;
}


export function writeValidation(r: ValidationRule): any {
  const out: any = { type: r.type === "any" ? "any" : r.type, allowBlank: r.allowBlank !== false, showErrorMessage: true, showInputMessage: !!r.prompt };
  if (r.operator && r.type !== "list" && r.type !== "custom") out.operator = r.operator;
  const f = (s?: string) => (s === undefined ? undefined : s.startsWith("=") ? s.slice(1) : s);
  if (r.type === "list") {
    const f1 = r.formula1 ?? "";
    out.formulae = [f1.startsWith("=") ? f1.slice(1) : `"${f1.split(";").map((x) => x.trim()).join(",")}"`];
    if (r.showDropdown === false) out.showDropDown = true;
  } else {
    out.formulae = [f(r.formula1), f(r.formula2)].filter((x) => x !== undefined).map((x) => (x !== undefined && /^-?\d+([.,]\d+)?$/.test(x) ? Number(x.replace(",", ".")) : x));
  }
  if (r.error) out.error = r.error;
  if (r.errorTitle) out.errorTitle = r.errorTitle;
  if (r.errorStyle) out.errorStyle = r.errorStyle;
  if (r.prompt) out.prompt = r.prompt;
  if (r.promptTitle) out.promptTitle = r.promptTitle;
  return out;
}

export function writeCondRule(r: CondRule, priority: number): any {
  const style = r.style ? toXStyle(r.style) : undefined;
  if (style) {
    // Koşullu biçimde yazı tipi adı/boyutu yazılmaz
    if (style.font) {
      delete style.font.name;
      delete style.font.size;
      delete style.font.family;
      delete style.font.scheme;
      if (!Object.keys(style.font).length) delete style.font;
    }
    if (style.fill) style.fill = { type: "pattern", pattern: "solid", bgColor: style.fill.fgColor };
  }
  const fx = (v: string) => addXlfnPrefixes(v.startsWith("=") ? v.slice(1) : v);
  switch (r.type) {
    case "cellIs":
      return { type: "cellIs", operator: r.operator, formulae: (r.values ?? []).map(fx), priority, style };
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith":
      return { type: "containsText", operator: r.type === "containsText" ? "containsText" : r.type === "notContainsText" ? "notContains" : r.type, text: r.text, priority, style };
    case "containsBlanks":
    case "notContainsBlanks":
    case "containsErrors":
      return { type: "containsText", operator: r.type, priority, style };
    case "duplicateValues":
    case "uniqueValues":
      return { type: r.type, priority, style };
    case "top10":
      return { type: "top10", rank: r.rank ?? 10, percent: !!r.percent, bottom: !!r.bottom, priority, style };
    case "aboveAverage":
      return { type: "aboveAverage", aboveAverage: !r.below, priority, style };
    case "expression":
      return { type: "expression", formulae: (r.values ?? []).map(fx), priority, style };
    case "timePeriod":
      return { type: "timePeriod", timePeriod: r.timePeriod, priority, style };
    case "dataBar":
      return { type: "dataBar", priority, gradient: true, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF" + (r.color ?? "#638EC6").slice(1) } };
    case "colorScale": {
      const cols = r.colors ?? ["#F8696B", "#FFEB84", "#63BE7B"];
      const cfvo = cols.length === 2 ? [{ type: "min" }, { type: "max" }] : [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }];
      return { type: "colorScale", priority, cfvo, color: cols.map((c) => ({ argb: "FF" + c.slice(1) })) };
    }
  }
  return null;
}

