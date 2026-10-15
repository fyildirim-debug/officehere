/** Sayfa yapısı: kâğıt kodları, üst/alt bilgi kodları ve yazdırma öğeleri (printOptions … headerFooter). */
import { colToName, formatRange } from "@officehere/formula";
import type { PageSetup } from "../../model/types";
import { attrs, esc } from "./xml";

export const PAPER: Record<number, PageSetup["paper"]> = { 1: "Letter", 5: "Legal", 8: "A3", 9: "A4", 11: "A5" };
export const PAPER_ID: Record<string, number> = { Letter: 1, Legal: 5, A3: 8, A4: 9, A5: 11 };

/** Üst/alt bilgi kodları: dosya (&P, &N, &A, &F, &D) ↔ arayüz (&[Sayfa] …). */
const HF_CODES: [string, string][] = [
  ["&P", "&[Sayfa]"],
  ["&N", "&[Sayfalar]"],
  ["&A", "&[Sekme]"],
  ["&F", "&[Dosya]"],
  ["&D", "&[Tarih]"],
];

export function stripHF(s: string): string {
  let t = s.replace(/&[LCR]/g, "").replace(/&"[^"]*"/g, "").replace(/&\d+/g, "");
  for (const [file, ui] of HF_CODES) t = t.split(file).join(ui);
  return t.replace(/&(?!\[)[A-Z]/g, "").replace(/&&/g, "&").trim();
}

export function hfToFile(s: string): string {
  let t = s.replace(/&(?!\[)/g, "&&");
  for (const [file, ui] of HF_CODES) t = t.split(ui).join(file);
  return "&C" + t;
}

/** Sayfaya sığdırma: [fitToWidth, fitToHeight] ya da yok. */
export function fitOf(p: PageSetup): [number, number] | null {
  if (p.fitMode === "sheet") return [1, 1];
  if (p.fitMode === "rows") return [0, 1];
  if (p.fitMode === "cols" || p.fitToWidth) return [1, 0];
  return null;
}

/** printOptions, pageMargins, pageSetup ve headerFooter öğeleri (şema sırasıyla). */
export function pageXml(p: PageSetup): string {
  let out = "";
  const po = attrs([
    ["headings", p.printHeadings ? "1" : undefined],
    ["gridLines", p.printGridlines ? "1" : undefined],
  ]);
  if (po) out += `<printOptions${po}/>`;
  const m = p.margins;
  if (m) {
    const ma = attrs([
      ["left", m.left],
      ["right", m.right],
      ["top", m.top],
      ["bottom", m.bottom],
      ["header", m.header],
      ["footer", m.footer],
    ]);
    if (ma) out += `<pageMargins${ma}/>`;
  }
  const fit = fitOf(p);
  out += `<pageSetup${attrs([
    ["paperSize", PAPER_ID[p.paper] ?? 9],
    ["orientation", p.orientation],
    ["pageOrder", p.pageOrder === "overThenDown" ? "overThenDown" : undefined],
    ["blackAndWhite", p.blackAndWhite ? "1" : undefined],
    ["draft", p.draft ? "1" : undefined],
    ["fitToWidth", fit?.[0]],
    ["fitToHeight", fit?.[1]],
  ])}/>`;
  if (p.header || p.footer) out += `<headerFooter>${p.header ? `<oddHeader>${esc(hfToFile(p.header))}</oddHeader>` : ""}${p.footer ? `<oddFooter>${esc(hfToFile(p.footer))}</oddFooter>` : ""}</headerFooter>`;
  return out;
}

const quoteSheet = (name: string) => `'${name.replace(/'/g, "''")}'`;
const abs = (cell: string) => cell.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");

/** Yazdırma alanı ve başlıkları: workbook.xml'deki yerleşik adlar (_xlnm.Print_Area, _xlnm.Print_Titles). */
export function printNames(p: PageSetup, sheetName: string, sheetIndex: number): string[] {
  const out: string[] = [];
  const q = quoteSheet(sheetName);
  if (p.printArea) {
    // Tek hücre de aralık olarak (A1:A1) yazılır: okuyucular yerleşik adı aralık bekler
    const [a, b = a] = formatRange(p.printArea).split(":");
    const ref = /^\d+$/.test(a) || /^[A-Z]+$/.test(a) ? `$${a}:$${b}` : `${abs(a)}:${abs(b)}`;
    out.push(`<definedName name="_xlnm.Print_Area" localSheetId="${sheetIndex}">${esc(`${q}!${ref}`)}</definedName>`);
  }
  if (p.titleRows || p.titleCols) {
    const ranges: string[] = [];
    if (p.titleCols) ranges.push(`${q}!$${colToName(p.titleCols.c1)}:$${colToName(p.titleCols.c2)}`);
    if (p.titleRows) ranges.push(`${q}!$${p.titleRows.r1 + 1}:$${p.titleRows.r2 + 1}`);
    out.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${sheetIndex}">${esc(ranges.join(","))}</definedName>`);
  }
  return out;
}
