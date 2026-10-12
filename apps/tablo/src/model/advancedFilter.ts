/**
 * Gelişmiş Filtre: ölçüt aralığının ilk satırı liste sütun başlıklarıdır; aynı
 * satırdaki koşullar VE, farklı satırlar VEYA ile birleşir. Koşul yazımı:
 * "Ege" (ile başlar, * ? joker), "=Ege" (tam eşit), ">100", "<>Ege", ">=01.01.2026"
 * ve boş hücre (koşul yok). "Yalnızca benzersiz kayıtlar" yinelenen satırları atar.
 */
import { hasWildcards, parseCellInput, wildcardToRegExp, type Scalar } from "@officehere/formula";
import { displayText } from "./display";
import type { Rect } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

type Cond = { col: number; op: "=" | "<>" | ">" | "<" | ">=" | "<=" | "begins"; operand: string };

const up = (s: string) => s.trim().toLocaleUpperCase("tr-TR");

function operandValue(s: string): Scalar {
  const p = parseCellInput(s);
  return p.kind === "value" ? (p.value ?? s) : s;
}

function match(c: Cond, v: Scalar, text: string): boolean {
  if (c.op === "begins") {
    if (hasWildcards(c.operand)) return wildcardToRegExp(c.operand + "*", true).test(text);
    return up(text).startsWith(up(c.operand));
  }
  if (c.operand === "") {
    const blank = v === null || text === "";
    return c.op === "=" ? blank : c.op === "<>" ? !blank : false;
  }
  const o = operandValue(c.operand);
  if (typeof o === "number") {
    if (typeof v !== "number") return c.op === "<>";
    switch (c.op) {
      case "=":
        return v === o;
      case "<>":
        return v !== o;
      case ">":
        return v > o;
      case "<":
        return v < o;
      case ">=":
        return v >= o;
      case "<=":
        return v <= o;
    }
  }
  const a = up(text);
  const b = up(String(o));
  if ((c.op === "=" || c.op === "<>") && hasWildcards(c.operand)) {
    const hit = wildcardToRegExp(c.operand, true).test(text);
    return c.op === "=" ? hit : !hit;
  }
  const cmp = a.localeCompare(b, "tr");
  switch (c.op) {
    case "=":
      return cmp === 0;
    case "<>":
      return cmp !== 0;
    case ">":
      return cmp > 0;
    case "<":
      return cmp < 0;
    case ">=":
      return cmp >= 0;
    case "<=":
      return cmp <= 0;
  }
  return false;
}

/** Ölçüt aralığından koşul satırları (liste sütunlarına eşlenmiş). */
export function parseCriteria(wb: WorkbookModel, sheet: SheetModel, list: Rect, criteria: Rect): Cond[][] {
  const headers = new Map<string, number>();
  for (let c = list.c1; c <= list.c2; c++) headers.set(up(displayText(wb, sheet, list.r1, c)), c);
  const rows: Cond[][] = [];
  for (let r = criteria.r1 + 1; r <= criteria.r2; r++) {
    const conds: Cond[] = [];
    for (let c = criteria.c1; c <= criteria.c2; c++) {
      const col = headers.get(up(displayText(wb, sheet, criteria.r1, c)));
      const raw = displayText(wb, sheet, r, c);
      if (col === undefined || raw === "") continue;
      const m = /^(<>|>=|<=|=|>|<)(.*)$/.exec(raw);
      conds.push(m ? { col, op: m[1] as Cond["op"], operand: m[2] } : { col, op: "begins", operand: raw });
    }
    rows.push(conds);
  }
  return rows;
}

/** Listenin ölçütlere uyan veri satırları (başlık hariç). */
export function advancedMatches(wb: WorkbookModel, sheet: SheetModel, list: Rect, criteria: Rect | null, unique: boolean): number[] {
  const rows = criteria ? parseCriteria(wb, sheet, list, criteria) : [];
  // Boş ölçüt satırı her şeyi kapsar (Excel'deki gibi)
  const any = !rows.length || rows.some((r) => !r.length);
  const seen = new Set<string>();
  const out: number[] = [];
  for (let r = list.r1 + 1; r <= list.r2; r++) {
    const ok = any || rows.some((conds) => conds.every((c) => match(c, wb.getValue(sheet, r, c.col), displayText(wb, sheet, r, c.col))));
    if (!ok) continue;
    if (unique) {
      const key: string[] = [];
      for (let c = list.c1; c <= list.c2; c++) key.push(up(displayText(wb, sheet, r, c)));
      const k = key.join("\u0001");
      if (seen.has(k)) continue;
      seen.add(k);
    }
    out.push(r);
  }
  return out;
}
