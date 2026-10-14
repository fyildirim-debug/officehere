/**
 * XLSX tabloları (ListObject, ECMA-376 §18.5) doğrudan paket XML'inden okunur.
 * exceljs'in tablo okuyucusu güvenilmez: <tableColumn> içinde
 * <calculatedColumnFormula>/<totalsRowFormula> olunca sonraki sütunları
 * düşürür, autoFilter sütunu sayısı tutmayınca "filterButton" hatasıyla
 * dosyanın tamamını açılamaz kılar ve headerRowCount yazılmadığında (Excel'in
 * varsayılanı) başlık satırını yok sayar. Bu yüzden tablo parçaları exceljs'e
 * verilen kopyadan çıkarılır (xlsxGuard) ve burada okunur.
 */
import { parseRangeAddress } from "@officehere/formula";
import type { TableDef } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { relsPathOf, resolveTarget, sheetPaths, unescapeXml } from "./ooxmlZip";
import { loadZip } from "./zipCache";

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unescapeXml(m[1]) : undefined;
}

/** ST_Xstring kaçışları: "_x000a_" → satır sonu. */
export function decodeXString(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}

const isTrue = (v: string | undefined, def: boolean) => (v === undefined ? def : v === "1" || v === "true");

/** Tek tablo parçasından model tanımı (geçersizse null). */
export function parseTableXml(xml: string, newId: () => string): TableDef | null {
  const head = /<(?:\w+:)?table\b[^>]*>/.exec(xml)?.[0];
  if (!head) return null;
  const ref = parseRangeAddress((attr(head, "ref") ?? "").replace(/\$/g, ""));
  if (!ref) return null;
  const headerRow = Number(attr(head, "headerRowCount") ?? 1) !== 0;
  const totalRow = Number(attr(head, "totalsRowCount") ?? 0) > 0;
  const columns: string[] = [];
  const colsXml = /<(?:\w+:)?tableColumns\b[^>]*>([\s\S]*?)<\/(?:\w+:)?tableColumns>/.exec(xml)?.[1] ?? "";
  for (const m of colsXml.matchAll(/<(?:\w+:)?tableColumn\b[^>]*>/g)) columns.push(decodeXString(attr(m[0], "name") ?? ""));
  const width = ref.c2 - ref.c1 + 1;
  // Sütun sayısı aralıkla tutmazsa aralık esas alınır (eksik adlar "SütunN")
  if (columns.length > width) columns.length = width;
  for (let i = columns.length; i < width; i++) columns.push(`Sütun${i + 1}`);
  const seen = new Set<string>();
  for (let i = 0; i < columns.length; i++) {
    let n = columns[i].trim() || `Sütun${i + 1}`;
    let k = 2;
    while (seen.has(n.toLocaleLowerCase("tr-TR"))) n = `${columns[i].trim() || "Sütun"}${k++}`;
    seen.add(n.toLocaleLowerCase("tr-TR"));
    columns[i] = n;
  }
  const si = /<(?:\w+:)?tableStyleInfo\b[^>]*>/.exec(xml)?.[0] ?? "";
  const af = /<(?:\w+:)?autoFilter\b[^>]*?(\/?)>([\s\S]*?)(?:<\/(?:\w+:)?autoFilter>|$)/.exec(xml);
  // Filtre düğmeleri: autoFilter yoksa gizli; tüm sütunlarda hiddenButton="1" ise gizli
  let filterButton = headerRow && !!af;
  if (af && !af[1]) {
    const cols = [...af[2].matchAll(/<(?:\w+:)?filterColumn\b[^>]*>/g)];
    if (cols.length >= width && cols.every((c) => isTrue(attr(c[0], "hiddenButton"), false))) filterButton = false;
  }
  const name = attr(head, "displayName") || attr(head, "name") || "Tablo";
  return {
    id: newId(),
    name,
    ref,
    style: attr(si, "name") || "TableStyleMedium2",
    headerRow,
    totalRow,
    // tableStyleInfo yoksa önceki davranış korunur (varsayılan stil, şeritli satırlar)
    bandedRows: isTrue(attr(si, "showRowStripes"), !si),
    bandedCols: isTrue(attr(si, "showColumnStripes"), false),
    firstColumn: isTrue(attr(si, "showFirstColumn"), false),
    lastColumn: isTrue(attr(si, "showLastColumn"), false),
    filterButton,
    columns,
    filter: { range: totalRow ? { ...ref, r2: Math.max(ref.r1, ref.r2 - 1) } : ref, columns: {} },
  };
}

/** Paketteki tüm tabloları sayfalarına ekler; okunamayanların sayısını döndürür. */
export async function readTablesFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<number> {
  const zip = await loadZip(data);
  if (!zip.file("xl/workbook.xml")) return 0;
  const paths = await sheetPaths(zip);
  let failed = 0;
  const names = new Set<string>();
  for (const sheet of wb.sheets) {
    const p = paths.get(sheet.name);
    const rels = p ? await zip.file(relsPathOf(p))?.async("string") : undefined;
    if (!p || !rels) continue;
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      if (!/Type="[^"]*\/table"/.test(m[0]) || /TargetMode="External"/.test(m[0])) continue;
      const target = /Target="([^"]+)"/.exec(m[0])?.[1];
      const xml = target ? await zip.file(resolveTarget(p, unescapeXml(target)))?.async("string") : undefined;
      if (!xml) {
        failed++;
        continue;
      }
      try {
        const t = parseTableXml(xml, () => wb.newId("tbl"));
        if (!t) {
          failed++;
          continue;
        }
        // Çakışan ya da aynı adlı tablo eklenmez (formül başvuruları karışmasın)
        const key = t.name.toLocaleLowerCase("tr-TR");
        const clash = sheet.tables.some((x) => !(t.ref.r2 < x.ref.r1 || t.ref.r1 > x.ref.r2 || t.ref.c2 < x.ref.c1 || t.ref.c1 > x.ref.c2));
        if (names.has(key) || clash) {
          failed++;
          continue;
        }
        names.add(key);
        sheet.tables.push(t);
      } catch {
        failed++;
      }
    }
  }
  return failed;
}
