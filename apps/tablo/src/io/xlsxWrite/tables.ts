/**
 * Tablolar (tableN.xml) ve tablonun hücrelere etkisi. Sütun listesi, hesaplanmış sütunlar, toplam
 * satırı ve otomatik filtre sonradan (xlsxExtras.writeSheetExtras) modelden yeniden yazılır; burada
 * tablo öğesi, aralık ve stil bilgisi üretilir. Veri satırı olmayan tabloya boş bir veri satırı
 * eklenir (tablo en az bir veri satırı içerir); başlık satırı boşsa sütun adı, toplam satırı boşsa
 * "Toplam" etiketi ve son sütunda ALTTOPLAM formülü yazılır.
 */
import { colToName } from "@officehere/formula";
import type { TableDef } from "../../model/types";
import { tableFileName } from "../xlsxPackage";
import { esc, XML_DECL } from "./xml";

const enc = (r1: number, c1: number, r2: number, c2: number) => `${colToName(c1)}${r1 + 1}:${colToName(c2)}${r2 + 1}`;

export interface TablePlan {
  def: TableDef;
  name: string;
  /** Kitap genelinde sıra numarası (id ve dosya adı). */
  id: number;
  path: string;
  /** Başlık satırı (sütun adları) ya da yok. */
  headerRow: number | null;
  /** Veri satırları [ilk, son] ve doldurulan sütun sayısı. */
  dataRows: [number, number];
  dataWidth: number;
  /** Toplam satırı ya da yok. */
  totalRow: number | null;
  ref: string;
}

/** Tablo planı; geçersiz tabloda (boş sütun adı, aralıktan az sütun) null: tablo ve hücreleri yazılmaz. */
export function planTable(t: TableDef, id: number): TablePlan | null {
  if (!t.columns.length || t.columns.some((c) => !c)) return null;
  const dataWidth = t.ref.c2 - t.ref.c1 + 1;
  if (dataWidth > t.columns.length) return null;
  const start = t.headerRow ? t.ref.r1 + 1 : t.ref.r1;
  const end = t.totalRow ? t.ref.r2 - 1 : t.ref.r2;
  const rows = Math.max(1, end - start + 1);
  const first = t.ref.r1 + (t.headerRow ? 1 : 0);
  const height = rows + (t.headerRow ? 1 : 0) + (t.totalRow ? 1 : 0);
  return {
    def: t,
    name: tableFileName(t.name),
    id,
    path: `xl/tables/table${id}.xml`,
    headerRow: t.headerRow ? t.ref.r1 : null,
    dataRows: [first, first + rows - 1],
    dataWidth,
    totalRow: t.totalRow ? first + rows : null,
    ref: enc(t.ref.r1, t.ref.c1, t.ref.r1 + height - 1, t.ref.c1 + t.columns.length - 1),
  };
}

/** Toplam satırının son sütunundaki formül (sütun birden çoksa). */
export function totalFormula(p: TablePlan): string | null {
  const cols = p.def.columns;
  return cols.length > 1 ? `SUBTOTAL(109,${p.name}[${cols[cols.length - 1]}])` : null;
}

export function tableXml(p: TablePlan): string {
  const t = p.def;
  const head =
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"` +
    ` id="${p.id}" name="${esc(p.name)}" displayName="${esc(p.name)}" ref="${p.ref}"` +
    (t.totalRow ? ' totalsRowCount="1"' : ' totalsRowShown="1"') +
    ` headerRowCount="${t.headerRow ? 1 : 0}">`;
  const afEnd = p.totalRow !== null ? p.totalRow - 1 : p.dataRows[1];
  const af = `<autoFilter ref="${enc(t.ref.r1, t.ref.c1, afEnd, t.ref.c1 + t.columns.length - 1)}"/>`;
  const cols = `<tableColumns count="${t.columns.length}">${t.columns.map((n, i) => `<tableColumn id="${i + 1}" name="${esc(n)}"/>`).join("")}</tableColumns>`;
  const theme = t.style === undefined ? "TableStyleMedium2" : t.style;
  const info = `<tableStyleInfo${theme ? ` name="${esc(theme)}"` : ""} showFirstColumn="${t.firstColumn ? 1 : 0}" showLastColumn="${t.lastColumn ? 1 : 0}" showRowStripes="${t.bandedRows ? 1 : 0}" showColumnStripes="${t.bandedCols ? 1 : 0}"/>`;
  return `${XML_DECL}${head}${af}${cols}${info}</table>`;
}
