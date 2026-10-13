/**
 * OpenDocument Hesap Tablosu (.ods, şablon .ots) okuyucusu. content.xml olay
 * akışıyla okunur; yinelenen satır/sütunlar son anlamlı hücreye kadar açılır,
 * boş kuyruklar atlanır. Değerler, formüller (OpenFormula → A1), stiller,
 * birleştirmeler, boyutlar, notlar, köprüler ve adlandırılmış aralıklar.
 */
import JSZip from "jszip";
import { cellKey, dateToSerial, errorFromCode, MAX_COLS, MAX_ROWS, type Scalar } from "@officehere/formula";
import { activateTheme } from "../../model/themeRuntime";
import { DEFAULT_FONT, DEFAULT_SIZE, type CellStyle } from "../../model/styles";
import { WorkbookModel, type CellInput, type SheetModel } from "../../model/workbook";
import { CELL_BUDGET, FormatError, safeSheetName } from "../bookUtil";
import { odfFormulaToA1, ofRefToA1 } from "./formula";
import { lengthToPx, OdsStyleSheet } from "./styles";
import { child, parseXml, scanXml, TreeBuilder } from "./xml";

export const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet";

/** "2024-03-15T10:30:00" → seri numarası (1900 sistemi). */
export function isoToSerial(v: string): number | null {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?/.exec(v.trim());
  if (!m) return null;
  let s = dateToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
  if (m[4]) s += (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6] ?? 0) + Number(m[7] ?? 0)) / 86400;
  return s;
}

/** "PT10H30M00S" / "P1DT2H" → gün kesri. */
export function durationToDays(v: string): number | null {
  const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(v.trim());
  if (!m) return null;
  const d = Number(m[2] ?? 0) + (Number(m[3] ?? 0) * 3600 + Number(m[4] ?? 0) * 60 + Number(m[5] ?? 0)) / 86400;
  return m[1] ? -d : d;
}

interface CellRec {
  attrs: Record<string, string>;
  covered: boolean;
  paras: string[];
  note?: { text: string; author?: string };
  link?: string;
}

interface RowRec {
  attrs: Record<string, string>;
  cells: CellRec[];
}

interface ColRun {
  start: number;
  count: number;
  styleName?: string;
  hidden: boolean;
  defaultCell?: string;
}

const STYLED_ROW_CAP = 1000;
const STYLED_RUN_CAP = 1024;

/** ODS/OTS paketini çalışma kitabına çevirir. */
export async function readOds(data: ArrayBuffer | Uint8Array): Promise<WorkbookModel> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new FormatError("Dosya geçerli bir OpenDocument paketi değil.");
  }
  const manifest = await zip.file("META-INF/manifest.xml")?.async("string");
  if (manifest && /encryption-data/.test(manifest)) throw new FormatError("Bu dosya parola ile şifrelenmiş. Şifreli OpenDocument dosyaları açılamaz.");
  const content = await zip.file("content.xml")?.async("string");
  if (!content) throw new FormatError("OpenDocument paketinde content.xml yok.");
  const sheetStyles = new OdsStyleSheet();
  const stylesXml = await zip.file("styles.xml")?.async("string");
  if (stylesXml) {
    try {
      const root = child(parseXml(stylesXml), "office:document-styles");
      sheetStyles.add(child(root, "office:font-face-decls"));
      sheetStyles.add(child(root, "office:styles"));
      sheetStyles.add(child(root, "office:automatic-styles"));
    } catch {
      /* stiller okunamazsa değerler yine gelir */
    }
  }
  const wb = new WorkbookModel(true);
  const warnings = new Set<string>();
  // Varsayılan hücre yazı tipi tema gövde yazı tipi olur (boş hücreler de aynı görünsün)
  const reader = new ContentReader(wb, sheetStyles, warnings);
  scanXml(content, reader);
  reader.finish();
  try {
    const meta = await zip.file("meta.xml")?.async("string");
    if (meta) readMeta(wb, meta);
  } catch {
    /* özellikler isteğe bağlı */
  }
  if (!wb.sheets.length) wb.addSheet("Sayfa1");
  const firstVisible = wb.sheets.findIndex((s) => !s.hidden);
  wb.activeIndex = Math.max(0, firstVisible);
  try {
    const settings = await zip.file("settings.xml")?.async("string");
    const act = settings ? /<config:config-item config:name="ActiveTable" config:type="string">([^<]*)</.exec(settings) : null;
    if (act) {
      const i = wb.sheets.findIndex((s) => s.name === reader.nameMap.get(act[1]) && !s.hidden);
      if (i >= 0) wb.activeIndex = i;
    }
  } catch {
    /* etkin sayfa isteğe bağlı */
  }
  wb.syncTables();
  wb.engine.recalc(true);
  wb.loadWarnings = [...warnings];
  return wb;
}

function readMeta(wb: WorkbookModel, xml: string): void {
  const pick = (tag: string) => {
    const m = new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`).exec(xml);
    return m ? m[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&") : undefined;
  };
  const date = (s?: string) => {
    const d = s ? new Date(s) : undefined;
    return d && !isNaN(d.getTime()) ? d : undefined;
  };
  wb.props = {
    title: pick("dc:title") || undefined,
    subject: pick("dc:subject") || undefined,
    author: pick("meta:initial-creator") || pick("dc:creator") || undefined,
    created: date(pick("meta:creation-date")),
    modified: date(pick("dc:date")),
  };
}

/** content.xml olay işleyicisi. */
class ContentReader {
  nameMap = new Map<string, string>();
  private capture: TreeBuilder | null = null;
  private sheet: SheetModel | null = null;
  private colRuns: ColRun[] = [];
  private colIndex = 0;
  private row: RowRec | null = null;
  private rowIndex = 0;
  private cell: CellRec | null = null;
  /** Hücre içindeki öğe yığını (metin bağlamı için). */
  private inCell: string[] = [];
  private noteParas: string[] | null = null;
  private noteAuthor = "";
  private subTable = 0;
  private names: { name: string; kind: "range" | "expr"; value: string; scope?: number }[] = [];
  private styleIds = new Map<string, number>();
  private cellsWritten = 0;
  private lastRow = -1;
  private lastCol = -1;
  private rowRuns: { start: number; count: number; styleName?: string; hidden: boolean }[] = [];
  private filterRanges: string[] = [];
  private baseReady = false;
  private matrices: { r1: number; c1: number; r2: number; c2: number }[] = [];

  constructor(
    private wb: WorkbookModel,
    private styles: OdsStyleSheet,
    private warnings: Set<string>,
  ) {}

  /** Varsayılan stil bilindiğinde (otomatik stiller okunduktan sonra) tema yazı tipi ayarlanır. */
  private ensureBase(): void {
    if (this.baseReady) return;
    this.baseReady = true;
    const base = this.styles.cellStyle(undefined);
    if (base.font && base.font !== DEFAULT_FONT) this.wb.theme = { ...this.wb.theme, minorFont: base.font };
    activateTheme(this.wb.theme);
  }

  private styleId(name: string | undefined): number {
    const key = name ?? "Default";
    let id = this.styleIds.get(key);
    if (id === undefined) {
      const st: CellStyle = { ...this.styles.cellStyle(key) };
      if (st.size === DEFAULT_SIZE) delete st.size;
      id = this.wb.styles.intern(st);
      this.styleIds.set(key, id);
    }
    return id;
  }

  open = (name: string, attrs: Record<string, string>, self: boolean): void => {
    if (this.capture) {
      this.capture.open(name, attrs);
      return;
    }
    if (this.subTable) {
      if (name === "table:table") this.subTable++;
      return;
    }
    if (this.cell) {
      // Kendini kapatan öğelerin kapanışını scanXml ayrıca bildirir
      this.cellOpen(name, attrs);
      return;
    }
    switch (name) {
      case "office:automatic-styles":
      case "office:font-face-decls":
        if (!self) this.capture = new TreeBuilder(name, attrs);
        return;
      case "table:table":
        this.startSheet(attrs);
        return;
      case "table:table-column":
        if (this.sheet) this.addColumn(attrs);
        return;
      case "table:table-row":
        if (this.sheet) this.row = { attrs, cells: [] };
        return;
      case "table:table-cell":
      case "table:covered-table-cell":
        if (!this.row) return;
        this.cell = { attrs, covered: name === "table:covered-table-cell", paras: [] };
        this.inCell = [];
        return;
      case "table:named-range":
        this.names.push({ name: attrs["table:name"], kind: "range", value: attrs["table:cell-range-address"] ?? "", scope: this.sheet?.id });
        return;
      case "table:named-expression":
        this.names.push({ name: attrs["table:name"], kind: "expr", value: attrs["table:expression"] ?? "", scope: this.sheet?.id });
        return;
      case "table:database-range":
        if (attrs["table:display-filter-buttons"] === "true" && attrs["table:target-range-address"]) this.filterRanges.push(attrs["table:target-range-address"]);
        return;
      case "table:content-validation":
        this.warnings.add("veri doğrulama kuralları");
        return;
      case "calcext:conditional-format":
        this.warnings.add("koşullu biçimlendirme");
        return;
      case "draw:frame":
        this.warnings.add("grafikler, resimler ve şekiller");
        return;
    }
  };

  close = (name: string): void => {
    if (this.capture) {
      if (this.capture.close()) {
        this.styles.add(this.capture.root);
        this.capture = null;
      }
      return;
    }
    if (this.subTable) {
      if (name === "table:table") this.subTable--;
      return;
    }
    if (this.cell) {
      if ((name === "table:table-cell" || name === "table:covered-table-cell") && !this.inCell.length) this.endCell();
      else this.cellClose(name);
      return;
    }
    if (name === "table:table-row" && this.row) {
      this.endRow(this.row);
      this.row = null;
    } else if (name === "table:table") this.endSheet();
  };

  text = (t: string): void => {
    if (this.capture) {
      this.capture.text(t);
      return;
    }
    if (!this.cell || this.subTable) return;
    const top = this.inCell[this.inCell.length - 1];
    if (top === "dc:creator") {
      this.noteAuthor += t;
      return;
    }
    if (!this.inCell.includes("text:p") && !this.inCell.includes("text:h")) return;
    if (this.inCell.includes("dc:date") || this.inCell.includes("text:note-citation")) return;
    const paras = this.noteParas ?? this.cell.paras;
    if (!paras.length) paras.push("");
    paras[paras.length - 1] += t;
  };

  private cellOpen(name: string, attrs: Record<string, string>): void {
    this.inCell.push(name);
    const paras = this.noteParas ?? this.cell!.paras;
    switch (name) {
      case "table:table":
        // Hücre içinde alt tablo: içeriği atlanır
        this.inCell.pop();
        this.subTable = 1;
        return;
      case "office:annotation":
        this.noteParas = [];
        this.noteAuthor = "";
        return;
      case "text:p":
      case "text:h":
        paras.push("");
        return;
      case "text:s": {
        const n = Math.min(1000, Math.max(1, Number(attrs["text:c"] ?? 1) || 1));
        if (paras.length) paras[paras.length - 1] += " ".repeat(n);
        return;
      }
      case "text:tab":
        if (paras.length) paras[paras.length - 1] += "\t";
        return;
      case "text:line-break":
        if (paras.length) paras[paras.length - 1] += "\n";
        return;
      case "text:a":
        if (!this.noteParas && attrs["xlink:href"] && !this.cell!.link) this.cell!.link = attrs["xlink:href"];
        return;
      case "draw:frame":
        this.warnings.add("grafikler, resimler ve şekiller");
        return;
    }
  }

  private cellClose(name: string): void {
    // Kendini kapatan öğelerde yığına hiç eklenmemiş olabilir
    if (this.inCell[this.inCell.length - 1] === name) this.inCell.pop();
    if (name === "office:annotation" && this.noteParas) {
      const text = this.noteParas.join("\n").trim();
      if (text) this.cell!.note = { text, author: this.noteAuthor.trim() || undefined };
      this.noteParas = null;
    }
  }

  private endCell(): void {
    this.row?.cells.push(this.cell!);
    this.cell = null;
    this.inCell = [];
    this.noteParas = null;
  }

  private startSheet(attrs: Record<string, string>): void {
    this.ensureBase();
    const raw = attrs["table:name"] ?? "";
    const name = safeSheetName(this.wb, raw);
    this.nameMap.set(raw, name);
    const sheet = this.wb.addSheet(name);
    const tp = this.styles.props("table", attrs["table:style-name"], "style:table-properties");
    if (tp["table:display"] === "false") sheet.hidden = true;
    const tab = tp["tableooo:tab-color"] ?? tp["table:tab-color"];
    if (tab && /^#[0-9a-fA-F]{6}$/.test(tab)) sheet.tabColor = tab.toUpperCase();
    this.sheet = sheet;
    this.colRuns = [];
    this.colIndex = 0;
    this.rowIndex = 0;
    this.lastRow = -1;
    this.lastCol = -1;
    this.rowRuns = [];
    this.matrices = [];
  }

  private addColumn(attrs: Record<string, string>): void {
    const count = Math.max(1, Math.min(MAX_COLS - this.colIndex, Number(attrs["table:number-columns-repeated"] ?? 1) || 1));
    if (this.colIndex >= MAX_COLS) return;
    this.colRuns.push({ start: this.colIndex, count, styleName: attrs["table:style-name"], hidden: attrs["table:visibility"] === "collapse" || attrs["table:visibility"] === "filter", defaultCell: attrs["table:default-cell-style-name"] });
    this.colIndex += count;
  }

  private colDefaultStyle(col: number): string | undefined {
    for (const r of this.colRuns) if (col >= r.start && col < r.start + r.count) return r.defaultCell;
    return undefined;
  }

  private endRow(row: RowRec): void {
    const sheet = this.sheet!;
    const repeat = Math.max(1, Number(row.attrs["table:number-rows-repeated"] ?? 1) || 1);
    const start = this.rowIndex;
    if (start >= MAX_ROWS) return;
    const count = Math.min(repeat, MAX_ROWS - start);
    this.rowIndex = start + count;
    const rowDefault = row.attrs["table:default-cell-style-name"];
    // Hücreleri sütunlara yerleştir: içerikli mi, yalnız stilli mi?
    type Placed = { col: number; count: number; rec: CellRec; content: boolean; styleId: number };
    const placed: Placed[] = [];
    let col = 0;
    let hasContent = false;
    let hasStyle = false;
    for (const rec of row.cells) {
      if (col >= MAX_COLS) break;
      const n = Math.min(Math.max(1, Number(rec.attrs["table:number-columns-repeated"] ?? 1) || 1), MAX_COLS - col);
      const a = rec.attrs;
      const content = !rec.covered && (a["office:value-type"] !== undefined || a["table:formula"] !== undefined || rec.paras.length > 0 || !!rec.note || a["table:number-columns-spanned"] !== undefined || a["table:number-rows-spanned"] !== undefined);
      const styleName = a["table:style-name"] ?? rowDefault ?? this.colDefaultStyle(col);
      const sid = styleName && styleName !== "Default" ? this.styleId(styleName) : content ? this.styleId(styleName) : 0;
      if (content) hasContent = true;
      if (sid) hasStyle = true;
      if (content || sid) placed.push({ col, count: n, rec, content, styleId: sid });
      col += n;
    }
    this.rowRuns.push({ start, count, styleName: row.attrs["table:style-name"], hidden: row.attrs["table:visibility"] === "collapse" || row.attrs["table:visibility"] === "filter" });
    if (!hasContent && !hasStyle) return;
    // Yalnız stilli satırlar sınırlı sayıda açılır (boş biçimli kuyruklar bellek tüketmesin)
    const times = hasContent ? count : Math.min(count, STYLED_ROW_CAP);
    for (let i = 0; i < times; i++) {
      const r = start + i;
      for (const p of placed) {
        if (!p.content && p.count > STYLED_RUN_CAP) continue;
        for (let j = 0; j < p.count; j++) {
          if (this.cellsWritten >= CELL_BUDGET) {
            this.warnings.add("çok büyük yinelenen alanların bir bölümü");
            return;
          }
          this.writeCell(sheet, r, p.col + j, p.rec, p.content, p.styleId);
        }
      }
      if (r > this.lastRow) this.lastRow = r;
    }
  }

  private writeCell(sheet: SheetModel, r: number, c: number, rec: CellRec, content: boolean, styleId: number): void {
    this.cellsWritten++;
    if (c > this.lastCol) this.lastCol = c;
    const k = cellKey(r, c);
    if (styleId) sheet.styles.set(k, styleId);
    if (!content) return;
    const a = rec.attrs;
    const cs = Number(a["table:number-columns-spanned"] ?? 1) || 1;
    const rs = Number(a["table:number-rows-spanned"] ?? 1) || 1;
    if (cs > 1 || rs > 1) sheet.merges.push({ r1: r, c1: c, r2: Math.min(MAX_ROWS - 1, r + rs - 1), c2: Math.min(MAX_COLS - 1, c + cs - 1) });
    if (rec.note) sheet.notes.set(k, rec.note);
    if (rec.link) sheet.links.set(k, rec.link);
    // Matris (dizi) formülü: bağlantı hücresi formülü taşır, diğer hücreler taşan değerdir
    const mc = Number(a["table:number-matrix-columns-spanned"] ?? 1) || 1;
    const mr = Number(a["table:number-matrix-rows-spanned"] ?? 1) || 1;
    if ((mc > 1 || mr > 1) && a["table:formula"]) this.matrices.push({ r1: r, c1: c, r2: r + mr - 1, c2: c + mc - 1 });
    else if (this.matrices.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2)) return;
    const { input, numFmt } = this.cellInput(rec);
    if (input !== null) this.wb.setInputRaw(sheet, r, c, input);
    // Tarih/saat/yüzde hücresinin biçimi yoksa varsayılan biçim
    if (numFmt && !this.wb.styles.get(styleId).numFmt) sheet.styles.set(k, this.wb.styles.patch(styleId, { numFmt }));
  }

  private cellInput(rec: CellRec): { input: CellInput; numFmt?: string } {
    const a = rec.attrs;
    const vt = a["office:value-type"];
    const cvt = a["calcext:value-type"];
    const text = rec.paras.join("\n");
    let value: Scalar = null;
    let numFmt: string | undefined;
    if (cvt === "error") value = errorFromCode(text.trim()) ?? errorFromCode("#VALUE!")!;
    else if (vt === "float" || vt === "percentage" || vt === "currency") {
      const n = Number(a["office:value"]);
      value = Number.isFinite(n) ? n : null;
      if (vt === "percentage") numFmt = "0%";
    } else if (vt === "date") {
      value = isoToSerial(a["office:date-value"] ?? "");
      numFmt = /T/.test(a["office:date-value"] ?? "") ? "dd.mm.yyyy hh:mm" : "dd.mm.yyyy";
    } else if (vt === "time") {
      value = durationToDays(a["office:time-value"] ?? "");
      numFmt = "hh:mm:ss";
    } else if (vt === "boolean") {
      const b = a["office:boolean-value"];
      value = b === "true" || b === "1";
    } else if (vt === "string") value = a["office:string-value"] ?? text;
    else if (text) value = text;
    const f = a["table:formula"];
    if (f) {
      const conv = odfFormulaToA1(f, (raw) => this.nameMap.get(raw) ?? raw);
      if (conv) return { input: { formula: conv }, numFmt };
      this.warnings.add("çevrilemeyen formüller (son hesaplanan değerleri kullanıldı)");
    }
    return { input: value, numFmt };
  }

  private endSheet(): void {
    const sheet = this.sheet;
    if (!sheet) return;
    // Sütunlar: en uzun kuyruk varsayılan genişlik olur, diğerleri yalnız kullanılan bölgede
    const width = (r: ColRun) => lengthToPx(this.styles.props("table-column", r.styleName, "style:table-column-properties")["style:column-width"]);
    const tail = [...this.colRuns].sort((x, y) => y.count - x.count)[0];
    const tailW = tail && tail.count >= 64 ? width(tail) : undefined;
    if (tailW && tailW > 2) sheet.cols.setDefault(Math.round(tailW));
    const colLimit = Math.max(this.lastCol + 1, 256);
    for (const r of this.colRuns) {
      if (r === tail && tail.count >= 64 && !r.hidden) continue;
      const w = width(r);
      const end = Math.min(r.start + r.count, colLimit, MAX_COLS);
      for (let c = r.start; c < end; c++) {
        if (w !== undefined && Math.abs(w - sheet.cols.defaultSize) > 0.5) sheet.cols.set(c, Math.round(w));
        if (r.hidden) sheet.cols.setHidden(c, true);
      }
      // Özel biçimli sütun (tüm sütun)
      if (r.defaultCell && r.defaultCell !== "Default") {
        const id = this.styleId(r.defaultCell);
        if (id) for (let c = r.start; c < end; c++) sheet.colStyles.set(c, id);
      }
    }
    // Satırlar: yalnız son anlamlı satıra kadar
    const rowLimit = this.lastRow + 1;
    for (const r of this.rowRuns) {
      if (r.start >= rowLimit) break;
      const rp = this.styles.props("table-row", r.styleName, "style:table-row-properties");
      const h = lengthToPx(rp["style:row-height"]);
      const custom = rp["style:use-optimal-row-height"] === "false";
      const end = Math.min(r.start + r.count, rowLimit);
      for (let i = r.start; i < end; i++) {
        if (custom && h !== undefined) sheet.rows.set(i, Math.max(0, Math.round(h)), true);
        if (r.hidden) sheet.rows.setHidden(i, true);
      }
    }
    this.sheet = null;
  }

  /** Belge sonu: adlar ve otomatik filtreler. */
  finish(): void {
    for (const n of this.names) {
      if (!n.name) continue;
      try {
        const map = (raw: string) => this.nameMap.get(raw) ?? raw;
        const f = n.kind === "range" ? ofRefToA1(n.value.replace(/^\[|\]$/g, ""), map) : odfFormulaToA1(n.value, map)?.slice(1);
        if (f) this.wb.engine.defineName(n.name, "=" + f, n.scope);
      } catch {
        /* geçersiz ad */
      }
    }
    for (const t of this.filterRanges) {
      const a1 = ofRefToA1(t.replace(/^\$/, ""), (raw) => this.nameMap.get(raw) ?? raw);
      const m = a1 ? /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(a1) : null;
      if (!m) continue;
      const sheet = this.wb.sheetByName((m[1] ?? m[2]).replace(/''/g, "'"));
      const parts = m[3].replace(/\$/g, "").split(":");
      const rng = parts.length === 2 ? parseA1Rect(parts[0], parts[1]) : null;
      if (sheet && rng && !sheet.autoFilter) sheet.autoFilter = { range: rng, columns: {} };
    }
  }
}

function parseA1Rect(a: string, b: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const p = (s: string) => {
    const m = /^([A-Z]{1,3})(\d+)$/i.exec(s);
    if (!m) return null;
    let c = 0;
    for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: Number(m[2]) - 1, c: c - 1 };
  };
  const x = p(a);
  const y = p(b);
  return x && y ? { r1: Math.min(x.r, y.r), c1: Math.min(x.c, y.c), r2: Math.max(x.r, y.r), c2: Math.max(x.c, y.c) } : null;
}
