/**
 * Formül sonucunun sayı biçimi (Genel biçimli hücreye girildiğinde):
 * BUGÜN/TARİH → tarih, ŞİMDİ → tarih ve saat, ZAMAN → saat; aksi halde
 * formülün başvurduğu ilk biçimli hücrenin biçimi (=TOPLA(C2:C7) para biçimini alır).
 */
import { EN_SYNTAX, tokenize } from "@officehere/formula";
import type { SheetModel, WorkbookModel } from "./workbook";

const FN_FORMATS: Record<string, string> = {
  TODAY: "dd.mm.yyyy",
  DATE: "dd.mm.yyyy",
  DATEVALUE: "dd.mm.yyyy",
  EDATE: "dd.mm.yyyy",
  EOMONTH: "dd.mm.yyyy",
  WORKDAY: "dd.mm.yyyy",
  "WORKDAY.INTL": "dd.mm.yyyy",
  NOW: "dd.mm.yyyy hh:mm",
  TIME: "hh:mm",
  TIMEVALUE: "hh:mm",
};

/** Biçim taşımayan (sonucu sayı/metin olan) fonksiyonlar: başvurunun biçimi devralınmaz. */
const NO_INHERIT = new Set(["COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "COUNTBLANK", "ROWS", "COLUMNS", "ROW", "COLUMN", "LEN", "MATCH", "XMATCH", "RANK", "RANK.EQ", "RANK.AVG", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "WEEKDAY", "WEEKNUM", "ISOWEEKNUM", "DATEDIF", "DAYS", "NETWORKDAYS", "TEXT", "CONCAT", "CONCATENATE", "TEXTJOIN", "IF", "IFS", "AND", "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT", "SUBTOTAL"]);

export function inheritedFormat(wb: WorkbookModel, sheet: SheetModel, formula: string): string | undefined {
  const body = formula.startsWith("=") ? formula.slice(1) : formula;
  const toks = tokenize(body, EN_SYNTAX);
  const firstFn = toks.find((t) => t.type === "func");
  const fnName = firstFn ? String(firstFn.value).toUpperCase().replace(/^_XLFN\./, "") : undefined;
  if (fnName && FN_FORMATS[fnName]) return FN_FORMATS[fnName];
  if (fnName && NO_INHERIT.has(fnName)) return undefined;
  for (const t of toks) {
    if (t.type !== "ref" || !t.ref || t.ref.refError || t.ref.kind === "cols" || t.ref.kind === "rows") continue;
    const target = t.ref.sheet ? wb.sheetByName(t.ref.sheet) : sheet;
    if (!target) continue;
    const st = wb.styleAt(target, t.ref.a.row, t.ref.a.col);
    if (st.numFmt && st.numFmt !== "General" && st.numFmt !== "@") return st.numFmt;
    // Aralık: ilk hücre biçimsizse ikinci hücreye de bak (başlık satırı içeren aralıklar)
    if (t.ref.b) {
      const st2 = wb.styleAt(target, Math.min(t.ref.a.row + 1, t.ref.b.row), t.ref.a.col);
      if (st2.numFmt && st2.numFmt !== "General" && st2.numFmt !== "@") return st2.numFmt;
    }
    return undefined;
  }
  return undefined;
}
