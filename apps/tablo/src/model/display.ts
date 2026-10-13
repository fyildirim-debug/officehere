/** Hücrenin ekranda nasıl görüneceği: biçimlenmiş metin, renkler, hizalama. */
import { formatValue, isError, localErrorName, toLocalFormula, type Scalar } from "@officehere/formula";
import type { CondFormatter, CondResult } from "./condfmt";
import type { CellStyle } from "./styles";
import { tableCellStyle } from "./tablestyles";
import type { SheetModel, WorkbookModel } from "./workbook";

export interface CellView {
  value: Scalar;
  text: string;
  /** Metin rengi (sayı biçimi rengi > koşullu > stil). */
  color: string;
  fill?: string;
  align: "left" | "center" | "right";
  style: CellStyle;
  isNumber: boolean;
  fillChar?: { index: number; char: string };
  bar?: CondResult["bar"];
  icon?: CondResult["icon"];
  /** Hücre içi onay kutusu: işaretli mi (değer DOĞRU). */
  checkbox?: boolean;
}

const EMPTY_STYLE: CellStyle = {};

export function computeView(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, cf: CondFormatter | null): CellView {
  const value = wb.getValue(sheet, row, col);
  let style = wb.styleAt(sheet, row, col);
  // Tablo stili (açık biçimlerin altında)
  if (sheet.tables.length) {
    const t = sheet.tableAt(row, col);
    if (t) {
      const ts = tableCellStyle(t, row, col);
      const merged: CellStyle = { ...ts, ...style };
      if (style.fill === undefined && ts.fill) merged.fill = ts.fill;
      if (style.color === undefined && ts.color) merged.color = ts.color;
      style = merged;
    }
  }
  let condRes: CondResult | null = null;
  if (cf && sheet.condFormats.length) {
    condRes = cf.evaluate(sheet, row, col);
    if (condRes?.style) style = { ...style, ...condRes.style };
  }
  let text = "";
  let numColor: string | undefined;
  let fillChar: CellView["fillChar"];
  let isNumber = false;
  if (sheet.showFormulas) {
    const f = wb.getFormula(sheet, row, col);
    if (f) text = toLocalFormula(f, "tr");
    else if (value !== null) text = typeof value === "number" ? String(value).replace(".", ",") : formatScalar(value);
  } else if (value !== null) {
    if (typeof value === "number") {
      isNumber = true;
      const r = formatValue(value, style.numFmt);
      text = r.text;
      numColor = r.color;
      fillChar = r.fill;
    } else if (typeof value === "string") {
      if (style.numFmt && style.numFmt !== "General" && style.numFmt !== "@") {
        const r = formatValue(value, style.numFmt);
        text = r.text;
        numColor = r.color;
      } else text = value;
    } else text = formatScalar(value);
  }
  // Simge kümesi "Yalnızca Simgeyi Göster"
  if (condRes?.icon?.only) text = "";
  // Onay kutusu: DOĞRU/YANLIŞ metni yerine kutu çizilir (formüller gösterilirken metin)
  const checkbox = style.checkbox && !sheet.showFormulas ? value === true : undefined;
  if (checkbox !== undefined) text = "";
  let align: CellView["align"] = "left";
  const h = style.hAlign;
  if (h === "center" || h === "centerContinuous" || h === "distributed") align = "center";
  else if (h === "right") align = "right";
  else if (h === "left" || h === "justify" || h === "fill") align = "left";
  else if (sheet.showFormulas) align = "left";
  else if (typeof value === "number") align = "right";
  else if (typeof value === "boolean" || isError(value)) align = "center";
  return {
    value,
    text,
    color: numColor ?? style.color ?? "#000000",
    fill: condRes?.scale ?? style.fill,
    align,
    style: style ?? EMPTY_STYLE,
    isNumber,
    fillChar,
    bar: condRes?.bar,
    icon: condRes?.icon,
    checkbox,
  };
}

export function formatScalar(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "DOĞRU" : "YANLIŞ";
  if (isError(v)) return localErrorName(v.code);
  return String(v);
}

/** Durum çubuğu ve kopyalama için hücrenin görünen metni. */
export function displayText(wb: WorkbookModel, sheet: SheetModel, row: number, col: number): string {
  const v = wb.getValue(sheet, row, col);
  if (v === null) return "";
  if (typeof v === "number") return formatValue(v, wb.styleAt(sheet, row, col).numFmt).text;
  return formatScalar(v);
}
