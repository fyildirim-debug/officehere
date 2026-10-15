/**
 * OpenDocument Hesap Tablosu (.ods) yazıcısı: değerler, formüller (of:),
 * hücre stilleri ve sayı biçimleri, birleştirmeler, sütun/satır boyutları,
 * gizli satır/sütun/sayfalar, notlar, köprüler, adlar ve otomatik filtre.
 * Paket: "mimetype" ilk ve sıkıştırılmamış, ardından manifest ve XML parçaları.
 * Geçiş katmanı (@officehere/core/odf): .ods'den açılan kitapta (wb.kaynak) settings.xml birleştirilir,
 * manifest.rdf taşınır, sayfaların gömülü grafikleri (passthrough.odf) alt belgeleriyle yazılır.
 */
import JSZip from "jszip";
import { OdfCarry } from "@officehere/core/odf";
import { cellKey, colToName, isDateFormat, isError, isTimeOnlyFormat, keyCol, keyRow, MAX_COLS, serialToDate, serialToTime, type Scalar } from "@officehere/formula";
import { DEFAULT_FONT, DEFAULT_SIZE } from "../../model/styles";
import type { Rect } from "../../model/types";
import type { SheetModel, WorkbookModel } from "../../model/workbook";
import { displayText } from "../../model/display";
import { a1FormulaToOdf } from "./formula";
import { codeToDataStyles } from "./numfmt";
import { cellStyleXml, pxToIn } from "./styles";
import { esc } from "./xml";
import { ODS_MIME } from "./reader";

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ' +
  'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" ' +
  // Yabancı öznitelik (ODF §3.17): hücre içi onay kutusu; başka uygulamalar yok sayar, değer DOĞRU/YANLIŞ kalır
  'xmlns:ohtablo="urn:officehere:tablo:1.0" office:version="1.3"';

/** Sütun sayısı: LibreOffice'in eski sürümleri de açabilsin diye en az 1024. */
const COLS_TOTAL = 1024;

/** ODF metni: ardışık boşluklar text:s, sekmeler text:tab olur. */
export function odfText(s: string): string {
  let out = "";
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === "\t") {
      out += "<text:tab/>";
      i++;
    } else if (ch === " ") {
      let j = i;
      while (s[j] === " ") j++;
      const n = j - i;
      // Baştaki/sondaki ya da birden çok boşluk korunmalı (ODF boşlukları daraltır)
      out += n === 1 && i > 0 && j < s.length ? " " : n === 1 ? "<text:s/>" : `<text:s text:c="${n}"/>`;
      i = j;
    } else {
      let j = i;
      while (j < s.length && s[j] !== " " && s[j] !== "\t") j++;
      out += esc(s.slice(i, j));
      i = j;
    }
  }
  return out;
}

function paragraphs(s: string, link?: string): string {
  return s
    .split("\n")
    .map((line) => `<text:p>${link ? `<text:a xlink:href="${esc(link)}" xlink:type="simple">${odfText(line)}</text:a>` : odfText(line)}</text:p>`)
    .join("");
}

function isoDate(serial: number, withTime: boolean): string {
  const d = serialToDate(serial);
  const p = (n: number, w = 2) => String(n).padStart(w, "0");
  let s = `${p(d.y, 4)}-${p(d.m)}-${p(d.d)}`;
  if (withTime) {
    const t = serialToTime(serial);
    s += `T${p(t.h)}:${p(t.mi)}:${p(t.s)}`;
  }
  return s;
}

function isoDuration(days: number): string {
  const neg = days < 0;
  let secs = Math.round(Math.abs(days) * 86400);
  const h = Math.floor(secs / 3600);
  secs -= h * 3600;
  const m = Math.floor(secs / 60);
  const s = secs - m * 60;
  return `${neg ? "-" : ""}PT${h}H${String(m).padStart(2, "0")}M${String(s).padStart(2, "0")}S`;
}

/** Değer öznitelikleri (office:value-type …). */
function valueAttrs(v: Scalar, numFmt: string | undefined): string {
  if (v === null) return "";
  if (typeof v === "boolean") return ` office:value-type="boolean" office:boolean-value="${v}"`;
  if (typeof v === "string") return ` office:value-type="string"`;
  if (isError(v)) return ` office:value-type="string" office:string-value=""`;
  if (!Number.isFinite(v)) return "";
  if (numFmt && isDateFormat(numFmt)) {
    if (isTimeOnlyFormat(numFmt) && v >= 0 && v < 1) return ` office:value-type="time" office:time-value="${isoDuration(v)}"`;
    if (v >= 1) return ` office:value-type="date" office:date-value="${isoDate(v, v % 1 !== 0)}"`;
  }
  const pct = numFmt && /%/.test(numFmt.replace(/"[^"]*"/g, ""));
  return ` office:value-type="${pct ? "percentage" : "float"}" office:value="${v}"`;
}

class StyleRegistry {
  cellNames = new Map<number, string>();
  dataNames = new Map<string, string>();
  fonts = new Set<string>();
  xml: string[] = [];
  dataXml: string[] = [];
  colStyles = new Map<string, string>();
  rowStyles = new Map<string, string>();
  tableStyles = new Map<string, string>();
  sizeXml: string[] = [];

  constructor(private wb: WorkbookModel) {}

  fontName(f: string): string {
    this.fonts.add(f);
    return f;
  }

  dataStyle(code: string | undefined): string | undefined {
    if (!code || code === "General") return undefined;
    let n = this.dataNames.get(code);
    if (!n) {
      n = `N${this.dataNames.size + 1}`;
      this.dataNames.set(code, n);
      this.dataXml.push(codeToDataStyles(n, code));
    }
    return n;
  }

  cell(id: number): string | undefined {
    if (!id) return undefined;
    let n = this.cellNames.get(id);
    if (!n) {
      const st = this.wb.styles.get(id);
      if (!Object.keys(st).length) return undefined;
      n = `ce${this.cellNames.size + 1}`;
      this.cellNames.set(id, n);
      this.xml.push(cellStyleXml(n, st, this.dataStyle(st.numFmt), (f) => this.fontName(f)));
    }
    return n;
  }

  column(px: number): string {
    const key = String(Math.round(px));
    let n = this.colStyles.get(key);
    if (!n) {
      n = `co${this.colStyles.size + 1}`;
      this.colStyles.set(key, n);
      this.sizeXml.push(`<style:style style:name="${n}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${pxToIn(px)}"/></style:style>`);
    }
    return n;
  }

  row(px: number, custom: boolean): string {
    const key = `${Math.round(px)}:${custom}`;
    let n = this.rowStyles.get(key);
    if (!n) {
      n = `ro${this.rowStyles.size + 1}`;
      this.rowStyles.set(key, n);
      this.sizeXml.push(`<style:style style:name="${n}" style:family="table-row"><style:table-row-properties style:row-height="${pxToIn(px)}" fo:break-before="auto" style:use-optimal-row-height="${custom ? "false" : "true"}"/></style:style>`);
    }
    return n;
  }

  table(hidden: boolean, tab?: string): string {
    const key = `${hidden}:${tab ?? ""}`;
    let n = this.tableStyles.get(key);
    if (!n) {
      n = `ta${this.tableStyles.size + 1}`;
      this.tableStyles.set(key, n);
      this.sizeXml.push(`<style:style style:name="${n}" style:family="table"><style:table-properties table:display="${hidden ? "false" : "true"}" style:writing-mode="lr-tb"${tab ? ` tableooo:tab-color="${tab.toLowerCase()}"` : ""}/></style:style>`);
    }
    return n;
  }
}

function ofSheetName(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function rangeAddress(sheet: string, r: Rect): string {
  const s = "$" + ofSheetName(sheet);
  return `${s}.$${colToName(r.c1)}$${r.r1 + 1}:.$${colToName(r.c2)}$${r.r2 + 1}`;
}

/** Taşınan gömülü nesnelerin yazım bağlamı: taşıyıcı ve nesnelerin otomatik stilleri. */
interface ObjectCtx {
  carry: OdfCarry;
  styles: string[];
}

/**
 * Sayfanın .ods'den gelen gömülü nesneleri (özgün çerçeveler; alt belgeler kaynak paketten kopyalanır). Veri aralığının
 * sayfası silindiyse, adı ya da yapısı değiştiyse grafik yanlış veri gösterir: yazılmaz ve uyarılır.
 */
function objectsXml(wb: WorkbookModel, sheet: SheetModel, ctx: ObjectCtx, warn: Set<string>): string {
  let out = "";
  for (const o of sheet.passthrough?.odf ?? []) {
    const what = `gömülü ${o.label.toLocaleLowerCase("tr-TR")}`;
    if (o.deps.some(([name, v]) => wb.sheets.find((s) => s.name === name)?.structVersion !== v)) {
      warn.add(`${what} (veri aldığı sayfa değiştirildi)`);
      continue;
    }
    const r = ctx.carry.object(o);
    if (!r) {
      warn.add(`${what} (özgün dosya bulunamadı)`);
      continue;
    }
    ctx.styles.push(...r.styles);
    out += r.xml;
  }
  return out;
}

function sheetXml(wb: WorkbookModel, sheet: SheetModel, reg: StyleRegistry, warn: Set<string>, objects: ObjectCtx): string {
  // Satır başına hücre sütunları
  const rowsMap = new Map<number, Set<number>>();
  const add = (r: number, c: number) => {
    let s = rowsMap.get(r);
    if (!s) rowsMap.set(r, (s = new Set()));
    s.add(c);
  };
  wb.engine.forEachCell(sheet.id, (r, c) => add(r, c));
  for (const k of sheet.styles.keys()) add(keyRow(k), keyCol(k));
  for (const k of sheet.notes.keys()) add(keyRow(k), keyCol(k));
  for (const k of sheet.links.keys()) add(keyRow(k), keyCol(k));
  // Taşan dizilerin değerleri (matris formülü olarak yazılır)
  const spillAnchors = new Map<number, { rows: number; cols: number }>();
  const spilled = new Set<number>();
  wb.engine.forEachCell(sheet.id, (r, c, rec) => {
    if (!rec.formula) return;
    const sp = wb.engine.getSpill(sheet.id, r, c);
    if (sp && sp.anchorRow === r && sp.anchorCol === c && (sp.rows > 1 || sp.cols > 1)) {
      spillAnchors.set(cellKey(r, c), { rows: sp.rows, cols: sp.cols });
      for (let i = 0; i < sp.rows; i++)
        for (let j = 0; j < sp.cols; j++) {
          add(r + i, c + j);
          if (i || j) spilled.add(cellKey(r + i, c + j));
        }
    }
  });
  const mergeByRow = new Map<number, Rect[]>();
  for (const m of sheet.merges) {
    for (let r = m.r1; r <= m.r2; r++) {
      let l = mergeByRow.get(r);
      if (!l) mergeByRow.set(r, (l = []));
      l.push(m);
      for (let c = m.c1; c <= m.c2; c++) add(r, c);
    }
  }
  let maxRow = -1;
  let maxCol = -1;
  for (const [r, cols] of rowsMap) {
    if (r > maxRow) maxRow = r;
    for (const c of cols) if (c > maxCol) maxCol = c;
  }
  for (const [r] of sheet.rows.customItems()) if (r > maxRow) maxRow = r;
  for (const r of sheet.rows.hiddenItems()) if (r > maxRow) maxRow = r;
  for (const r of sheet.rowStyles.keys()) if (r > maxRow) maxRow = r;
  let colEnd = maxCol;
  for (const [c] of sheet.cols.customItems()) if (c > colEnd) colEnd = c;
  for (const c of sheet.cols.hiddenItems()) if (c > colEnd) colEnd = c;
  for (const c of sheet.colStyles.keys()) if (c > colEnd) colEnd = c;
  colEnd = Math.min(colEnd, MAX_COLS - 1);

  const out: string[] = [];
  out.push(`<table:table table:name="${esc(sheet.name)}" table:style-name="${reg.table(sheet.hidden, sheet.tabColor)}">`);
  // Gömülü nesneler (.ods'den gelen grafikler): sütunlardan önce table:shapes
  const shapes = objectsXml(wb, sheet, objects, warn);
  if (shapes) out.push(`<table:shapes>${shapes}</table:shapes>`);
  // Sütunlar
  const colKey = (c: number) => `${reg.column(sheet.cols.rawSize(c))}|${sheet.cols.isHidden(c)}|${reg.cell(sheet.colStyles.get(c) ?? 0) ?? ""}`;
  let c = 0;
  while (c <= colEnd) {
    const k = colKey(c);
    let n = 1;
    while (c + n <= colEnd && colKey(c + n) === k) n++;
    const [cs, hidden, def] = k.split("|");
    out.push(`<table:table-column table:style-name="${cs}"${n > 1 ? ` table:number-columns-repeated="${n}"` : ""}${hidden === "true" ? ' table:visibility="collapse"' : ""} table:default-cell-style-name="${def || "Default"}"/>`);
    c += n;
  }
  const restCols = Math.max(COLS_TOTAL, colEnd + 1) - (colEnd + 1);
  if (restCols > 0) out.push(`<table:table-column table:style-name="${reg.column(sheet.cols.defaultSize)}" table:number-columns-repeated="${restCols}" table:default-cell-style-name="Default"/>`);
  const totalCols = colEnd + 1 + restCols;
  // Satırlar
  const rowStyle = (r: number) => reg.row(sheet.rows.rawSize(r), sheet.rows.isCustom(r) && sheet.rows.isManual(r));
  const spillAt = (r: number, c2: number) => spillAnchors.get(cellKey(r, c2));
  const inSpill = (r: number, c2: number) => spilled.has(cellKey(r, c2));
  let r = 0;
  while (r <= maxRow) {
    const cols = rowsMap.get(r);
    const rs = rowStyle(r);
    const hidden = sheet.rows.isHidden(r) ? ' table:visibility="collapse"' : "";
    const rowDef = sheet.rowStyles.get(r) ? reg.cell(sheet.rowStyles.get(r)!) : undefined;
    if (!cols?.size) {
      // Boş satırlar tek öğede yinelenir
      let n = 1;
      while (r + n <= maxRow && !rowsMap.get(r + n)?.size && rowStyle(r + n) === rs && sheet.rows.isHidden(r + n) === !!hidden && !sheet.rowStyles.get(r + n) && !rowDef) n++;
      out.push(`<table:table-row table:style-name="${rs}"${n > 1 ? ` table:number-rows-repeated="${n}"` : ""}${hidden}${rowDef ? ` table:default-cell-style-name="${rowDef}"` : ""}><table:table-cell table:number-columns-repeated="${totalCols}"/></table:table-row>`);
      r += n;
      continue;
    }
    const cells: string[] = [];
    const sorted = [...cols].filter((x) => x < totalCols).sort((a, b) => a - b);
    const merges = mergeByRow.get(r) ?? [];
    let at = 0;
    for (const col of sorted) {
      if (col > at) cells.push(`<table:table-cell${col - at > 1 ? ` table:number-columns-repeated="${col - at}"` : ""}/>`);
      const m = merges.find((x) => col >= x.c1 && col <= x.c2);
      const covered = !!m && (m.r1 !== r || m.c1 !== col);
      cells.push(cellXml(wb, sheet, r, col, reg, warn, covered, m && !covered ? m : undefined, spillAt(r, col), inSpill(r, col)));
      at = col + 1;
    }
    if (at < totalCols) cells.push(`<table:table-cell table:number-columns-repeated="${totalCols - at}"/>`);
    out.push(`<table:table-row table:style-name="${rs}"${hidden}${rowDef ? ` table:default-cell-style-name="${rowDef}"` : ""}>${cells.join("")}</table:table-row>`);
    r++;
  }
  if (maxRow < 0) out.push(`<table:table-row table:style-name="${rowStyle(0)}"><table:table-cell table:number-columns-repeated="${totalCols}"/></table:table-row>`);
  // Sayfaya özgü adlar
  const local = wb.engine.getNames().filter((n) => n.sheet === sheet.id);
  if (local.length) out.push(`<table:named-expressions>${local.map((n) => namedXml(n.name, n.formula, sheet.name)).join("")}</table:named-expressions>`);
  out.push("</table:table>");
  return out.join("");
}

function cellXml(wb: WorkbookModel, sheet: SheetModel, r: number, c: number, reg: StyleRegistry, warn: Set<string>, covered: boolean, merge: Rect | undefined, spill: { rows: number; cols: number } | undefined, spilled: boolean): string {
  const k = cellKey(r, c);
  const styleId = wb.styleIdAt(sheet, r, c);
  const st = wb.styles.get(styleId);
  // Biçimsiz hücre stilli satır/sütunda: açıkça "Default" yazılmazsa satır/sütun stilini devralır
  const sn = reg.cell(styleId) ?? (sheet.rowStyles.has(r) || sheet.colStyles.has(c) ? "Default" : undefined);
  const tag = covered ? "table:covered-table-cell" : "table:table-cell";
  let attrs = sn ? ` table:style-name="${sn}"` : "";
  if (merge) attrs += ` table:number-columns-spanned="${merge.c2 - merge.c1 + 1}" table:number-rows-spanned="${merge.r2 - merge.r1 + 1}"`;
  const rec = wb.engine.getRecord(sheet.id, r, c);
  const value = rec ? rec.value : spilled ? wb.getValue(sheet, r, c) : null;
  if (rec?.formula) {
    const f = a1FormulaToOdf(rec.formula.text);
    if (f) attrs += ` table:formula="${esc(f)}"`;
    else warn.add("tablo başvurulu formüller değer olarak");
    if (spill) attrs += ` table:number-matrix-columns-spanned="${spill.cols}" table:number-matrix-rows-spanned="${spill.rows}"`;
  }
  attrs += valueAttrs(value, st.numFmt);
  if (typeof value === "string" && rec?.formula) attrs += ` office:string-value="${esc(value)}"`;
  const note = sheet.notes.get(k);
  const link = sheet.links.get(k);
  let body = "";
  if (note) body += `<office:annotation>${note.author ? `<dc:creator>${esc(note.author)}</dc:creator>` : ""}${paragraphs(note.text)}</office:annotation>`;
  if (value !== null) {
    const text = typeof value === "string" ? value : isError(value) ? value.code : displayText(wb, sheet, r, c);
    body += paragraphs(text, link);
  } else if (link) body += paragraphs(link, link);
  return body ? `<${tag}${attrs}>${body}</${tag}>` : `<${tag}${attrs}/>`;
}

function namedXml(name: string, formula: string, firstSheet: string): string {
  const f = a1FormulaToOdf(formula);
  const base = `$${ofSheetName(firstSheet)}.$A$1`;
  // Sayfalı yalın aralık adı named-range, diğerleri named-expression
  const m = f ? /^of:=\[(\$[^\]#]+)\]$/.exec(f) : null;
  if (m) return `<table:named-range table:name="${esc(name)}" table:base-cell-address="${esc(m[1].split(":")[0])}" table:cell-range-address="${esc(m[1])}"/>`;
  if (!f) return "";
  return `<table:named-expression table:name="${esc(name)}" table:base-cell-address="${esc(base)}" table:expression="${esc(f)}"/>`;
}

function stylesXml(fontDecls: string, defaultFont: string): string {
  const font = esc(defaultFont);
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}>${fontDecls}<office:styles>` +
    `<style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/><style:text-properties style:font-name="${font}" fo:font-size="${DEFAULT_SIZE}pt" fo:language="tr" fo:country="TR"/></style:default-style>` +
    `<number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
    `<style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="${font}" fo:font-size="${DEFAULT_SIZE}pt" fo:color="#000000"/></style:style>` +
    `</office:styles></office:document-styles>`
  );
}

function metaXml(wb: WorkbookModel): string {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "");
  const p = wb.props;
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS}><office:meta>` +
    `<meta:generator>OfficeHere Tablo</meta:generator>` +
    (p.title ? `<dc:title>${esc(p.title)}</dc:title>` : "") +
    (p.subject ? `<dc:subject>${esc(p.subject)}</dc:subject>` : "") +
    (p.author ? `<meta:initial-creator>${esc(p.author)}</meta:initial-creator><dc:creator>${esc(p.author)}</dc:creator>` : "") +
    `<meta:creation-date>${(p.created ?? new Date()).toISOString().replace(/\.\d+Z$/, "")}</meta:creation-date><dc:date>${now}</dc:date>` +
    `</office:meta></office:document-meta>`
  );
}

function settingsXml(wb: WorkbookModel): string {
  const active = wb.sheets[wb.activeIndex]?.name ?? "";
  const sheets = wb.sheets
    .map((s) => {
      const f = s.frozen;
      const split = f.rows || f.cols;
      const items = split
        ? `<config:config-item config:name="HorizontalSplitMode" config:type="short">${f.cols ? 2 : 0}</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">${f.rows ? 2 : 0}</config:config-item><config:config-item config:name="HorizontalSplitPosition" config:type="int">${f.cols}</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">${f.rows}</config:config-item><config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item><config:config-item config:name="PositionRight" config:type="int">${f.cols}</config:config-item><config:config-item config:name="PositionBottom" config:type="int">${f.rows}</config:config-item>`
        : "";
      return `<config:config-item-map-entry config:name="${esc(s.name)}"><config:config-item config:name="ShowGrid" config:type="boolean">${s.showGridlines}</config:config-item>${items}</config:config-item-map-entry>`;
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings>` +
    `<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    `<config:config-item config:name="ViewId" config:type="string">view1</config:config-item><config:config-item-map-named config:name="Tables">${sheets}</config:config-item-map-named>` +
    `<config:config-item config:name="ActiveTable" config:type="string">${esc(active)}</config:config-item>` +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
  );
}

/** Çalışma kitabını ODS (ya da şablon için OTS) paketine yazar. */
export async function writeOds(wb: WorkbookModel, opts: { template?: boolean; warnings?: string[] } = {}): Promise<Uint8Array<ArrayBuffer>> {
  wb.engine.recalc();
  const reg = new StyleRegistry(wb);
  const warn = new Set<string>();
  // Varsayılan yazı tipi kitabın temasından (modül düzeyindeki değer başka kitaptan kalmış olabilir)
  const defaultFont = wb.theme.minorFont || DEFAULT_FONT;
  reg.fontName(defaultFont);
  // ODF geçiş katmanı: .ods'den açılan kitabın ayarları, RDF üst verisi ve gömülü nesneleri (wb.kaynak)
  const objects: ObjectCtx = { carry: new OdfCarry({ family: "spreadsheet", source: wb.kaynak }), styles: [] };
  const tables = wb.sheets.map((s) => sheetXml(wb, s, reg, warn, objects)).join("");
  const globals = wb.engine.getNames().filter((n) => n.sheet === undefined);
  const named = globals.length ? `<table:named-expressions>${globals.map((n) => namedXml(n.name, n.formula, wb.sheets[0]?.name ?? "Sayfa1")).join("")}</table:named-expressions>` : "";
  const filters = wb.sheets
    .filter((s) => s.autoFilter)
    .map((s, i) => `<table:database-range table:name="__Anonymous_Sheet_DB__${i}" table:target-range-address="${esc(rangeAddress(s.name, s.autoFilter!.range))}" table:display-filter-buttons="true"/>`)
    .join("");
  const fontDecls = `<office:font-face-decls>${[...reg.fonts].map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(/\s/.test(f) ? `'${f}'` : f)}"/>`).join("")}</office:font-face-decls>`;
  const content =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}>${fontDecls}<office:automatic-styles>` +
    reg.sizeXml.join("") +
    reg.dataXml.join("") +
    reg.xml.join("") +
    objects.styles.join("") +
    `</office:automatic-styles><office:body><office:spreadsheet>${tables}${named}${filters ? `<table:database-ranges>${filters}</table:database-ranges>` : ""}</office:spreadsheet></office:body></office:document-content>`;
  const mime = opts.template ? `${ODS_MIME}-template` : ODS_MIME;
  const zip = new JSZip();
  zip.file("mimetype", mime, { compression: "STORE" });
  zip.file(
    "META-INF/manifest.xml",
    `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
      `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/>` +
      `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
      `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
      `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
      `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>` +
      `</manifest:manifest>`,
  );
  zip.file("content.xml", content);
  zip.file("styles.xml", stylesXml(fontDecls, defaultFont));
  zip.file("meta.xml", metaXml(wb));
  zip.file("settings.xml", settingsXml(wb));
  if (opts.warnings) opts.warnings.push(...warn);
  let bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: mime });
  // settings.xml birleştirilir (yazıcının görünüm ayarları + kaynağın diğer ayarları), manifest.rdf ve alt belgeler eklenir
  if (objects.carry.active) bytes = await objects.carry.finish(bytes);
  return new Uint8Array(bytes);
}
