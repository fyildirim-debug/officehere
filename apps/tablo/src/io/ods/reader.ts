/**
 * OpenDocument Hesap Tablosu (.ods, şablon .ots) okuyucusu. content.xml olay
 * akışıyla okunur; yinelenen satır/sütunlar son anlamlı hücreye kadar açılır,
 * boş kuyruklar atlanır. Değerler, formüller (OpenFormula → A1), stiller,
 * birleştirmeler, boyutlar, notlar, köprüler ve adlandırılmış aralıklar.
 */
import JSZip from "jszip";
import { cellKey, dateToSerial, errorFromCode, MAX_COLS, MAX_ROWS, type Scalar } from "@officehere/formula";
import { OpcPackage, registerSource, releaseSource } from "@officehere/core/opc";
import { objectLabel, odfInfo, styleRefs, withNamespaces, type OdfInfo } from "@officehere/core/odf";
import { activateTheme } from "../../model/themeRuntime";
import { DEFAULT_SIZE, type CellStyle } from "../../model/styles";
import { emptyPassthrough } from "../../model/passthrough";
import { WorkbookModel, type CellInput, type SheetModel } from "../../model/workbook";
import { BIG_MERGE_CELLS, CELL_BUDGET, clampMerge, FormatError, safeSheetName } from "../bookUtil";
import { checkPackageLimits, readCollapsed } from "../xlsxGuard";
import { odfFormulaToA1, ofRefToA1 } from "./formula";
import { lengthToPx, OdsStyleSheet } from "./styles";
import { child, esc, find, kids, parseXml, scanXml, TreeBuilder, type XmlEl } from "./xml";

export const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet";

/** Ağacın XML'i (ön ekler olduğu gibi; değerler kaçışlanır). */
function xmlOf(el: XmlEl): string {
  let s = `<${el.name}`;
  for (const [k, v] of Object.entries(el.attrs)) s += ` ${k}="${esc(v)}"`;
  if (!el.children.length) return `${s}/>`;
  return `${s}>${el.children.map((c) => (typeof c === "string" ? esc(c) : xmlOf(c))).join("")}</${el.name}>`;
}

/** Aralık listesindeki sayfa adları ("Sayfa1.A1:Sayfa1.B3 'Satış 2'.C1:C9" → Sayfa1, Satış 2). */
function rangeSheets(v: string): string[] {
  const out = new Set<string>();
  for (const m of v.matchAll(/(?:^|[\s:])\$?(?:'((?:[^']|'')+)'|([^\s.':$][^\s.':]*))\.\$?[A-Za-z]{1,3}\$?\d+/g)) out.add(m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]);
  return [...out];
}

/** Kitap çöpe gidince kaynak paket kaydı da bırakılır (bellek dosya boyutu kadar tutulmasın). */
const sourceFinalizer = typeof FinalizationRegistry === "function" ? new FinalizationRegistry<string>((id) => releaseSource(id)) : null;

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

/**
 * ODS/OTS paketini çalışma kitabına çevirir. Yapısal sorunda report.notice doldurulur. Özgün paket kaynak kaydına
 * eklenir (`wb.kaynak`; `kaynak: false` ile eklenmez): .ods'ye kaydederken ayarlar, RDF üst verisi ve gömülü
 * nesneler (sayfaların `passthrough.odf`'u) buradan taşınır.
 */
export async function readOds(data: ArrayBuffer | Uint8Array, report?: { notice?: string }, opts: { kaynak?: boolean } = {}): Promise<WorkbookModel> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch (e) {
    throw new FormatError("Dosya bozuk ya da eksik indirilmiş: OpenDocument paketi okunamadı.", `${(e as Error)?.name ?? "Error"}: ${(e as Error)?.message ?? e}`);
  }
  // Paket sınırları (sıkıştırma bombası, tutarsız boyut bilgisi); aşırı sıkıştırılmış parçalar akışla ayıklanır
  const suspicious = checkPackageLimits(zip);
  const text = (path: string) => (suspicious.has(path) && zip.file(path) ? readCollapsed(zip.file(path)!) : zip.file(path)?.async("string"));
  const manifest = await zip.file("META-INF/manifest.xml")?.async("string");
  if (manifest && /encryption-data/.test(manifest)) throw new FormatError("Bu dosya parola ile şifrelenmiş. Şifreli OpenDocument dosyaları açılamaz.");
  const content = await text("content.xml");
  if (!content) throw new FormatError("Dosya bozuk ya da eksik: OpenDocument paketinde içerik bölümü (content.xml) bulunamadı.");
  const sheetStyles = new OdsStyleSheet();
  const stylesXml = await text("styles.xml");
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
  // Özgün paket (geçiş katmanı): açılamazsa ya da ODF değilse taşıma olmadan sürer
  let src: OpcPackage | undefined;
  let info: OdfInfo | null = null;
  if (opts.kaynak !== false) {
    try {
      src = await OpcPackage.open(data);
      info = await odfInfo(src);
      if (!info) src = undefined;
    } catch {
      src = undefined;
    }
  }
  const wb = new WorkbookModel(true);
  const warnings = new Set<string>();
  // Varsayılan hücre yazı tipi tema gövde yazı tipi olur (boş hücreler de aynı görünsün)
  const reader = new ContentReader(wb, sheetStyles, warnings);
  reader.keepObjects = !!src;
  try {
    scanXml(content, reader);
  } catch (e) {
    // Bozuk ya da yarıda kesilmiş içerik: hata noktasına kadar okunanlar kurtarılır
    if (!reader.hasData()) throw e;
    reader.salvage();
    warnings.add("içeriğin bozuk bölümden sonraki kısmı");
    if (report) report.notice = "Dosya bozuk ya da yarıda kesilmiş: okunabilen bölüm açıldı. Kaydetmeden önce içeriği denetleyin; okunamayan kısım kaydedince kaybolur.";
  }
  reader.finish();
  const objects = reader.storeObjects(info);
  if (src) {
    wb.kaynak = registerSource(src);
    sourceFinalizer?.register(wb, wb.kaynak);
  }
  if (objects && report) {
    const note = `Bu dosyadaki ${objects} gömülü grafik ya da nesne gösterilemiyor; .ods olarak kaydedince korunur, başka biçimde kaydedince kaybolur.`;
    report.notice = report.notice ? `${report.notice} ${note}` : note;
  }
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
  /** Gömülü nesneler taşınsın mı (özgün paket kaynak kaydında). */
  keepObjects = false;
  /** Kökteki ad alanı bildirimleri ve otomatik stiller (taşınan çerçevelerin bildirimleri ve stilleri için). */
  private rootNs: Record<string, string> = {};
  private autoStyles: XmlEl | null = null;
  private inShapes = false;
  /** Okunmakta olan sayfa çerçevesi ve okunan gömülü nesne çerçeveleri. */
  private frame: { b: TreeBuilder; sheet: SheetModel } | null = null;
  private objectFrames: { sheet: SheetModel; el: XmlEl }[] = [];

  /** Çerçeve kapandı: gömülü nesne (draw:object/draw:object-ole) değilse kaydedince kaybolur. */
  private endFrame(): void {
    const { b, sheet } = this.frame!;
    this.frame = null;
    if (find(b.root, "draw:object") || find(b.root, "draw:object-ole")) this.objectFrames.push({ sheet, el: b.root });
    else this.warnings.add("grafikler, resimler ve şekiller");
  }

  /**
   * Gömülü nesneleri sayfaların geçiş verisine yazar (`passthrough.odf`): özgün çerçeve, andığı otomatik stiller ve
   * veri aralığının sayfaları (okunduğu andaki yapı sürümleriyle). Taşınan nesne sayısını döner.
   */
  storeObjects(info: OdfInfo | null): number {
    let n = 0;
    for (const { sheet, el } of this.objectFrames) {
      const xml = withNamespaces(xmlOf(el), this.rootNs);
      const names = styleRefs(xml);
      const styles = kids(this.autoStyles ?? undefined, "style:style")
        .filter((s) => names.has(s.attrs["style:name"]))
        .map((s) => withNamespaces(xmlOf(s), this.rootNs));
      const obj = find(el, "draw:object");
      const deps: [string, number][] = [];
      let ok = true;
      for (const raw of rangeSheets(obj?.attrs["draw:notify-on-update-of-ranges"] ?? "")) {
        const name = this.nameMap.get(raw);
        const dep = name === raw ? this.wb.sheets.find((s) => s.name === name) : undefined;
        if (!dep) ok = false;
        else deps.push([dep.name, dep.structVersion]);
      }
      // Sayfa adı değiştirilerek açılan (geçersiz ad) sayfaya bağlı nesne doğru yazılamaz: kaydedince kaybolur.
      if (!ok) {
        this.warnings.add("grafikler, resimler ve şekiller");
        continue;
      }
      sheet.passthrough ??= emptyPassthrough(sheet.structVersion);
      (sheet.passthrough.odf ??= []).push({ xml, styles, label: objectLabel(info, xml), deps });
      n++;
    }
    return n;
  }

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
    // Kitabın kendi teması esas alınır (modül düzeyindeki varsayılan önceki kitaptan kalmış olabilir)
    if (base.font && base.font !== this.wb.theme.minorFont) this.wb.theme = { ...this.wb.theme, minorFont: base.font };
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
    if (this.frame) {
      this.frame.b.open(name, attrs);
      return;
    }
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
      case "office:document-content":
        for (const [k, v] of Object.entries(attrs)) if (k.startsWith("xmlns:")) this.rootNs[k.slice(6)] = v;
        return;
      case "office:automatic-styles":
      case "office:font-face-decls":
        if (!self) this.capture = new TreeBuilder(name, attrs);
        return;
      case "table:table":
        this.startSheet(attrs);
        return;
      case "table:shapes":
        this.inShapes = !self;
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
        // Sayfaya bağlı çerçeve (table:shapes): gömülü nesneyse özgün XML'iyle taşınır (kapanınca karar verilir).
        if (this.inShapes && this.sheet && this.keepObjects) {
          this.frame = { b: new TreeBuilder(name, attrs), sheet: this.sheet };
          return;
        }
        this.warnings.add("grafikler, resimler ve şekiller");
        return;
    }
  };

  close = (name: string): void => {
    if (this.frame) {
      if (this.frame.b.close()) this.endFrame();
      return;
    }
    if (this.capture) {
      if (this.capture.close()) {
        this.styles.add(this.capture.root);
        if (this.capture.root.name === "office:automatic-styles") this.autoStyles = this.capture.root;
        this.capture = null;
      }
      return;
    }
    if (name === "table:shapes") {
      this.inShapes = false;
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
    if (this.frame) {
      this.frame.b.text(t);
      return;
    }
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

  /** Modelde sütuna atanacak stil (endSheet ile aynı kural: "Default" ve tanımsız → yok). */
  private colStyleId(name: string | undefined): number {
    return name && name !== "Default" ? this.styleId(name) : 0;
  }

  /** [col, col+n) aralığını sütun öbeklerine böler: her parçanın varsayılan hücre stili ve model sütun stili. */
  private colSegments(col: number, n: number): { start: number; count: number; style?: string; colSid: number }[] {
    const out: { start: number; count: number; style?: string; colSid: number }[] = [];
    const end = col + n;
    let at = col;
    for (const r of this.colRuns) {
      if (at >= end) break;
      const rEnd = r.start + r.count;
      if (rEnd <= at) continue;
      if (r.start > at) {
        // Tanımsız sütunlar (öbekler arasında boşluk)
        const stop = Math.min(r.start, end);
        out.push({ start: at, count: stop - at, colSid: 0 });
        at = stop;
        if (at >= end) break;
      }
      const stop = Math.min(rEnd, end);
      out.push({ start: at, count: stop - at, style: r.defaultCell, colSid: this.colStyleId(r.defaultCell) });
      at = stop;
    }
    if (at < end) out.push({ start: at, count: end - at, colSid: 0 });
    return out;
  }

  /** Şimdiye kadar hücre okundu mu (bozuk dosyada kurtarılacak veri var mı). */
  hasData(): boolean {
    return this.cellsWritten > 0;
  }

  /** Ayrıştırma yarıda kaldı: açık hücre/satır/sayfa kapatılır. */
  salvage(): void {
    this.capture = null;
    this.subTable = 0;
    if (this.cell) this.endCell();
    if (this.row && this.sheet) this.endRow(this.row);
    this.row = null;
    if (this.sheet) this.endSheet();
  }

  private endRow(row: RowRec): void {
    const sheet = this.sheet!;
    const repeat = Math.max(1, Number(row.attrs["table:number-rows-repeated"] ?? 1) || 1);
    const start = this.rowIndex;
    if (start >= MAX_ROWS) return;
    const count = Math.min(repeat, MAX_ROWS - start);
    this.rowIndex = start + count;
    const rowDefault = row.attrs["table:default-cell-style-name"];
    // Hücreleri sütunlara yerleştir: içerikli mi, yalnız stilli mi? Kendi stili olmayan (yinelenen)
    // hücreler her sütunun varsayılan hücre stilini alır; bu yüzden sütun öbekleri sınırında bölünür.
    // Model hücre stili yoksa sütun stiline düşer: yalnız etkin stil sütununkinden farklıysa hücreye yazılır
    // (açıkça "Default" stilli hücre, stilli sütunda biçimsiz kalmalıdır).
    type Placed = { col: number; count: number; rec: CellRec; content: boolean; styleId: number; own: boolean };
    const placed: Placed[] = [];
    let col = 0;
    let hasContent = false;
    let hasStyle = false;
    for (const rec of row.cells) {
      if (col >= MAX_COLS) break;
      const n = Math.min(Math.max(1, Number(rec.attrs["table:number-columns-repeated"] ?? 1) || 1), MAX_COLS - col);
      const a = rec.attrs;
      const content = !rec.covered && (a["office:value-type"] !== undefined || a["table:formula"] !== undefined || rec.paras.length > 0 || !!rec.note || a["table:number-columns-spanned"] !== undefined || a["table:number-rows-spanned"] !== undefined);
      const explicit = a["table:style-name"] ?? rowDefault;
      for (const seg of this.colSegments(col, n)) {
        const name = explicit ?? seg.style;
        // İçerikli hücre varsayılan stilin yazı tipi/boyutunu taşır; boş hücrede "Default" biçimsiz demektir
        const sid = content ? this.styleId(name) : name && name !== "Default" ? this.styleId(name) : 0;
        const own = sid !== seg.colSid;
        if (content) hasContent = true;
        if (own) hasStyle = true;
        if (content || own) placed.push({ col: seg.start, count: seg.count, rec, content, styleId: sid, own });
      }
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
          this.writeCell(sheet, r, p.col + j, p.rec, p.content, p.styleId, p.own);
        }
      }
      if (r > this.lastRow) this.lastRow = r;
    }
  }

  private writeCell(sheet: SheetModel, r: number, c: number, rec: CellRec, content: boolean, styleId: number, own: boolean): void {
    this.cellsWritten++;
    if (c > this.lastCol) this.lastCol = c;
    const k = cellKey(r, c);
    if (own) sheet.styles.set(k, styleId);
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
      const isTail = r === tail && tail.count >= 64 && !r.hidden;
      const w = isTail ? undefined : width(r);
      const end = Math.min(r.start + r.count, colLimit, MAX_COLS);
      for (let c = r.start; c < end && !isTail; c++) {
        if (w !== undefined && Math.abs(w - sheet.cols.defaultSize) > 0.5) sheet.cols.set(c, Math.round(w));
        if (r.hidden) sheet.cols.setHidden(c, true);
      }
      // Özel biçimli sütun (tüm sütun)
      if (r.defaultCell && r.defaultCell !== "Default") {
        const id = this.styleId(r.defaultCell);
        if (id) for (let c = r.start; c < end; c++) sheet.colStyles.set(c, id);
      }
    }
    // Dev birleştirmeler (bozuk yayılma sayıları): kullanılan alana kırpılır, yine dev ise atılır
    const used = { rows: this.lastRow + 1, cols: this.lastCol + 1 };
    sheet.merges = sheet.merges.flatMap((m) => {
      if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= BIG_MERGE_CELLS) return [m];
      const c = clampMerge(m, used);
      if (c !== m) this.warnings.add(c ? "çok büyük birleştirilmiş alanlar (kullanılan alana kırpıldı)" : "çok büyük birleştirilmiş alanlar");
      return c ? [c] : [];
    });
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
