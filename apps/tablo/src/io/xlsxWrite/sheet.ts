/**
 * Çalışma sayfası parçası, iki aşamada:
 *  1) plan(): hücre anahtarları satır sırasıyla dizilir; sayfa öğeleri (görünüm, sütunlar, koşullu
 *     biçim, köprüler, sayfa yapısı…) bir iskelet XmlPart'a yazılır (sheetData yer tutucudur);
 *     ilişkiler, notlar, resimler ve tablolar hazırlanır. Sonradan işlem (xlsxPost) bu iskelet
 *     üzerinde çalışır.
 *  2) stream(): iskeletin sheetData'ya kadarki bölümü, satırlar ve kalanı ~64 KB'lık dilimlerle
 *     sıkıştırıcıya akıtılır; hücre başına nesne üretilmez.
 *
 * Hücrelerin son durumu eski (exceljs) yazıcıyla aynıdır: kendi stili olmayan hücre satır, yoksa
 * sütun stilini taşır; tablo başlığı/toplamı, taşan dizi değerleri, değersiz köprü ve not hücreleri
 * "kendiliğinden oluşan" hücrelerdir (satır ve sütun stillerinin birleşimini alır); birleştirilmiş
 * alanın gizli hücreleri değer taşımaz ve ana hücrenin biçimini alır.
 */
import { addr, colToName, formatRange, isError, keyCol, keyRow, type CellRecord, type Scalar } from "@officehere/formula";
import type { CondRule, DataTableDef } from "../../model/types";
import type { SheetModel, WorkbookModel } from "../../model/workbook";
import { MAX_COLS } from "../../model/rect";
import { pxToPt, widthPxToChars } from "../../model/axis";
import { BIG_MERGE_CELLS, CELL_BUDGET, disjointMerges } from "../bookUtil";
import { usesDynamicFunctions } from "../xlsxDynamic";
import { addXlfnPrefixes } from "../xlsxFormula";
import { passthroughCfRules } from "../xlsxPost";
import { writeCondRule } from "../xlsxRules";
import { injectDataTables } from "../xlsxWhatIf";
import { WORKSHEET_ORDER, XmlPart } from "../xmlPart";
import { commentsXml, vmlXml, type NoteOut } from "./comments";
import { buildDrawing, type DrawingOut, type MediaItem } from "./drawings";
import { fitOf, pageXml } from "./pageSetup";
import { Rels, REL } from "./rels";
import type { SharedStrings } from "./sharedStrings";
import type { StyleSheet } from "./styles";
import { planTable, tableXml, totalFormula, type TablePlan } from "./tables";
import { attrs, esc, NS_MAIN, NS_MC, NS_R, NS_X14, NS_X14AC, NS_XM, XML_DECL } from "./xml";
import { ByteOut, asciiBytes, type DeflateSink } from "./zip";

/* eslint-disable @typescript-eslint/no-explicit-any */

/** sheetData yer tutucusu: sonradan işlem bitince satırlar buraya akıtılır. */
const ROWS_MARK = "<!--oh:satirlar-->";
const ROWS_PLACEHOLDER = `<sheetData>${ROWS_MARK}</sheetData>`;

/** Kendiliğinden oluşan hücrenin değeri: sabit, formül (önbelleksiz) ya da yalnız köprü. */
type Phantom = Scalar | { f: string } | { link: string };

interface Extra {
  v?: Phantom;
  /** Birleştirilmiş alanın gizli hücresi: değeri silinir, biçimi (onay kutusundan önce) budur. */
  clear?: number;
}

/** Ortak yazma bağlamı (kitap genelinde sayaçlar ve listeler). */
export interface WriteContext {
  wb: WorkbookModel;
  styles: StyleSheet;
  sst: SharedStrings;
  media: MediaItem[];
  /** Kitap genelinde tablo ve çizim sayaçları. */
  tables: number;
  drawings: number;
  /** Dinamik dizi formülü yazıldı mı (metadata.xml gerekir). */
  dynamic: boolean;
  /** Onay kutulu biçim taşıyan stil numaraları. */
  checkbox: Set<number>;
}

const COL_NAMES: string[] = [];
const colName = (c: number) => COL_NAMES[c] ?? (COL_NAMES[c] = colToName(c));
const COL_BYTES: Uint8Array[] = [];
const colBytes = (c: number) => COL_BYTES[c] ?? (COL_BYTES[c] = asciiBytes(colName(c)));

const B_CELL = asciiBytes('<c r="');
const B_S = asciiBytes(' s="');
const B_V = asciiBytes("><v>");
const B_TS = asciiBytes(' t="s"><v>');
const B_TB = asciiBytes(' t="b"><v>');
const B_TE = asciiBytes(' t="e"><v>');
const B_V_END = asciiBytes("</v></c>");
const B_SLASH_CLOSE = asciiBytes("/>");
const B_ROW_END = asciiBytes("</row>");

/** Formülün dosya biçimi ve dinamik dizi işlevi içerip içermediği (basit formüllerde ayrıştırmasız). */
function fileFormula(text: string): { f: string; dynamic: boolean } {
  if (!text.includes("(") && !text.includes("[")) return { f: text, dynamic: false };
  return { f: addXlfnPrefixes(text), dynamic: usesDynamicFunctions(text) };
}

const EXT_ICONS = new Set(["3Triangles", "3Stars", "5Boxes"]);

/** Koşullu biçim kuralının sol üst hücresi (metin/tarih kuralı formülleri için). */
function topLeft(rule: CondRule): string {
  const r = rule.ranges[0];
  return r ? addr(r.r1, r.c1) : "A1";
}

/**
 * Metin kuralları: dosyadaki tür, işleç, metin ve formül (masaüstü uygulamaların yazdığı biçimde).
 * exceljs işleci tür diye yazıyor ve metni yazmıyordu ("notContains" geçersiz tür, boşluk/hata
 * kurallarında geçersiz işleç); okuyucumuz bu kuralların metnini zaten okumaz, model değişmez.
 */
function textRule(op: string, text: string | undefined, tl: string): { type: string; operator?: string; text?: string; formula: string } | null {
  const lit = `"${(text ?? "").replace(/"/g, '""')}"`;
  switch (op) {
    case "containsText":
      return { type: "containsText", operator: "containsText", text: text ?? "", formula: `NOT(ISERROR(SEARCH(${lit},${tl})))` };
    case "notContains":
      return { type: "notContainsText", operator: "notContains", text: text ?? "", formula: `ISERROR(SEARCH(${lit},${tl}))` };
    case "beginsWith":
      return { type: "beginsWith", operator: "beginsWith", text: text ?? "", formula: `LEFT(${tl},LEN(${lit}))=${lit}` };
    case "endsWith":
      return { type: "endsWith", operator: "endsWith", text: text ?? "", formula: `RIGHT(${tl},LEN(${lit}))=${lit}` };
    case "containsBlanks":
      return { type: "containsBlanks", formula: `LEN(TRIM(${tl}))=0` };
    case "notContainsBlanks":
      return { type: "notContainsBlanks", formula: `LEN(TRIM(${tl}))>0` };
    case "containsErrors":
      return { type: "containsErrors", formula: `ISERROR(${tl})` };
    default:
      return null;
  }
}

function timeFormula(period: string, tl: string): string | undefined {
  switch (period) {
    case "thisWeek":
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case "lastWeek":
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case "nextWeek":
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case "yesterday":
      return `FLOOR(${tl},1)=TODAY()-1`;
    case "today":
      return `FLOOR(${tl},1)=TODAY()`;
    case "tomorrow":
      return `FLOOR(${tl},1)=TODAY()+1`;
    case "last7Days":
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case "lastMonth":
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case "thisMonth":
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case "nextMonth":
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
}

const cfvoXml = (c: { type: string; value?: unknown }) => `<cfvo${attrs([["type", c.type], ["val", c.value as number | string | undefined]])}/>`;
const colorRgb = (tag: string, c: { argb?: string }) => `<${tag} rgb="${esc(String(c.argb))}"/>`;
const formulaXml = (f: string) => `<formula>${esc(f)}</formula>`;

function guidUpper(): string {
  const h = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16)).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-4${h.slice(13, 16)}-${((parseInt(h[16], 16) & 3) | 8).toString(16)}${h.slice(17, 20)}-${h.slice(20)}}`.toUpperCase(); // yerel-ayar: ascii
}

export class SheetWriter {
  readonly path: string;
  readonly rels = new Rels();
  frame!: XmlPart;
  comments: { path: string; vmlPath: string; xml: string; vml: string } | null = null;
  drawing: DrawingOut | null = null;
  tables: { path: string; xml: string }[] = [];
  /** Paylaşılan metin kullanılacak mı (sharedStrings.xml gerekir). */
  hasStrings = false;

  private recKeys = new Float64Array(0);
  private recs: CellRecord[] | null = null;
  private styleKeys = new Float64Array(0);
  private extraKeys = new Float64Array(0);
  private extras = new Map<number, Extra>();
  private metaRows: number[] = [];
  private dataTables = new Map<number, DataTableDef[]>();
  private links: Map<number, string>;
  private rowMaxLevel = 0;

  constructor(
    private readonly ctx: WriteContext,
    readonly sheet: SheetModel,
    readonly index: number,
  ) {
    this.path = `xl/worksheets/sheet${index + 1}.xml`;
    this.links = sheet.links;
  }

  private extra(k: number): Extra {
    let e = this.extras.get(k);
    if (!e) this.extras.set(k, (e = {}));
    return e;
  }

  private exists(k: number): boolean {
    return this.extras.has(k) || this.sheet.styles.has(k) || this.ctx.wb.engine.hasInput(this.sheet.id, keyRow(k), keyCol(k));
  }

  /** Hücrenin biçimi (onay kutusundan önce): kendi stili, kayıtlı hücrede satır/sütun stili, kendiliğinden oluşanda birleşim. */
  private baseXf(k: number, rec: CellRecord | undefined, e: Extra | undefined): number {
    if (e?.clear !== undefined) return e.clear;
    const own = this.sheet.styles.get(k);
    if (own !== undefined) return this.ctx.styles.xf(own);
    const r = keyRow(k);
    const c = keyCol(k);
    if (rec) {
      const id = this.sheet.rowStyles.get(r) ?? this.sheet.colStyles.get(c) ?? 0;
      return id ? this.ctx.styles.xf(id) : 0;
    }
    return this.ctx.styles.mergedXf(this.sheet.rowStyles.get(r), this.sheet.colStyles.get(c));
  }

  /* ------------------------------------------------------------------ 1) plan */

  plan(): void {
    const { wb } = this.ctx;
    const sheet = this.sheet;
    const engine = wb.engine;
    // Kayıtlı hücreler: anahtarlar (satır sırasıyla); dosyadan okunan sayfada zaten sıralıdır
    const n = engine.cellCount(sheet.id);
    const keys = new Float64Array(n);
    const recs: CellRecord[] = new Array(n);
    const anchors: number[] = [];
    let sorted = true;
    let prev = -1;
    let i = 0;
    for (const [k, rec] of engine.cellEntries(sheet.id)) {
      keys[i] = k;
      recs[i++] = rec;
      if (k < prev) sorted = false;
      prev = k;
      if (rec.formula) {
        if (rec.formula.spill) anchors.push(k);
      } else if (typeof rec.value === "string") this.hasStrings = true;
    }
    if (!sorted) keys.sort();
    this.recKeys = keys;
    this.recs = sorted ? recs : null;

    // Tablolar (hücrelerden önce: başlık ve toplam satırı boşsa doldurur)
    const tablePlans = new Map<string, TablePlan>();
    for (const t of sheet.tables) {
      const p = planTable(t, 0);
      if (!p) continue;
      if (p.headerRow !== null) t.columns.forEach((name, j) => (this.extra(p.headerRow! * MAX_COLS + t.ref.c1 + j).v = name));
      for (let r = p.dataRows[0]; r <= p.dataRows[1]; r++) for (let j = 0; j < p.dataWidth; j++) this.extra(r * MAX_COLS + t.ref.c1 + j).v = null;
      if (p.totalRow !== null) {
        const f = totalFormula(p);
        t.columns.forEach((_, j) => (this.extra(p.totalRow! * MAX_COLS + t.ref.c1 + j).v = j === 0 ? "Toplam" : j === t.columns.length - 1 && f ? { f } : null));
      }
      if (t.headerRow || t.totalRow) this.hasStrings = true;
      // Aynı adlı tablo: sonraki geçerli, sırası ilkinin (exceljs ile aynı)
      tablePlans.set(p.name, p);
    }
    // Taşan dizilerin değerleri (sabit olarak yazılır)
    for (const k of anchors) {
      const r = keyRow(k);
      const c = keyCol(k);
      const sp = engine.getSpill(sheet.id, r, c);
      if (!sp || sp.anchorRow !== r || sp.anchorCol !== c || (sp.rows <= 1 && sp.cols <= 1)) continue;
      for (let a = 0; a < sp.rows; a++)
        for (let b = 0; b < sp.cols; b++) {
          if (!a && !b) continue;
          const v = wb.getValue(sheet, r + a, c + b);
          if (v === null) continue;
          this.extra((r + a) * MAX_COLS + c + b).v = v;
          if (typeof v === "string") this.hasStrings = true;
        }
    }
    // Değersiz köprüler: köprü adresi metin olarak
    for (const [k, link] of sheet.links) {
      if (engine.hasInput(sheet.id, keyRow(k), keyCol(k))) continue;
      this.extra(k).v = { link };
      this.hasStrings = true;
    }
    for (const k of sheet.notes.keys()) this.extra(k);
    // Birleştirilmiş alanlar: gizli hücreler değersiz, ana hücrenin biçimiyle (dev olanlar ve bütçe aşımı hariç)
    let budget = CELL_BUDGET;
    for (const m of disjointMerges(sheet.merges).merges) {
      const area = (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1);
      if (area > BIG_MERGE_CELLS || area > budget) continue;
      budget -= area;
      const mk = m.r1 * MAX_COLS + m.c1;
      let masterXf = 0;
      if (this.exists(mk)) masterXf = this.baseXf(mk, engine.getRecord(sheet.id, m.r1, m.c1), this.extras.get(mk));
      else if (wb.styleIdAt(sheet, m.r1, m.c1)) {
        this.extra(mk);
        masterXf = this.baseXf(mk, undefined, undefined);
      }
      for (let r = m.r1; r <= m.r2; r++)
        for (let c = m.c1; c <= m.c2; c++) {
          if (r === m.r1 && c === m.c1) continue;
          const k = r * MAX_COLS + c;
          if (!masterXf && !this.exists(k)) continue;
          this.extra(k).clear = masterXf;
        }
    }
    this.extraKeys = Float64Array.from(this.extras.keys()).sort();
    this.styleKeys = Float64Array.from(sheet.styles.keys()).sort();
    for (const dt of sheet.dataTables) {
      const k = dt.ref.r1 * MAX_COLS + dt.ref.c1;
      const list = this.dataTables.get(k);
      if (list) list.push(dt);
      else this.dataTables.set(k, [dt]);
    }
    // İçeriği olmasa da yazılan satırlar: özel yükseklik, gizli, anahat (yalnız stili olan boş satır
    // yazılmaz: okuyucumuz satır stilini okumaz, eski yazıcı da yazmıyordu)
    const meta = new Set<number>([...sheet.rows.customItems().map(([r]) => r), ...sheet.rows.hiddenItems(), ...sheet.rows.levelItems().map(([r]) => r)]);
    this.metaRows = [...meta].filter((r) => r >= 0).sort((a, b) => a - b);
    this.rowMaxLevel = sheet.rows.maxLevel();

    this.buildFrame([...tablePlans.values()]);
  }

  /** Kaydı olan bir köprü hücresi metin değerli mi (köprü yalnız metin değerli hücrede yazılır). */
  private linkCells(): [number, string][] {
    const out: [number, string][] = [];
    const engine = this.ctx.wb.engine;
    for (const [k, link] of this.links) {
      if (this.extras.get(k)?.clear !== undefined) continue;
      const rec = engine.getRecord(this.sheet.id, keyRow(k), keyCol(k));
      if (rec && (rec.formula || typeof rec.value !== "string")) continue;
      out.push([k, link]);
    }
    return out.sort((a, b) => a[0] - b[0]);
  }

  private dimension(): string {
    let minR = Infinity;
    let maxR = -1;
    let minC = Infinity;
    let maxC = -1;
    const scan = (keys: Float64Array) => {
      for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        const c = k % MAX_COLS;
        if (c < minC) minC = c;
        if (c > maxC) maxC = c;
      }
      if (keys.length) {
        minR = Math.min(minR, Math.floor(keys[0] / MAX_COLS));
        maxR = Math.max(maxR, Math.floor(keys[keys.length - 1] / MAX_COLS));
      }
    };
    scan(this.recKeys);
    scan(this.styleKeys);
    scan(this.extraKeys);
    if (this.metaRows.length) {
      minR = Math.min(minR, this.metaRows[0]);
      maxR = Math.max(maxR, this.metaRows[this.metaRows.length - 1]);
    }
    if (maxR < 0) return "A1:A1";
    if (maxC < 0) return `A${minR + 1}:A${maxR + 1}`;
    return `${colName(minC)}${minR + 1}:${colName(maxC)}${maxR + 1}`;
  }

  private colsXml(): string {
    const { styles } = this.ctx;
    const s = this.sheet;
    const set = new Set<number>([...s.cols.customItems().map(([c]) => c), ...s.cols.hiddenItems(), ...s.colStyles.keys(), ...s.cols.levelItems().map(([c]) => c)]);
    const list = [...set].filter((c) => c >= 0 && c < MAX_COLS).sort((a, b) => a - b);
    type Col = { min: number; max: number; width: number; style: number; hidden: boolean; level: number };
    const out: Col[] = [];
    let run: Col | null = null;
    for (const c of list) {
      const width = widthPxToChars(s.cols.isCustom(c) ? s.cols.rawSize(c) : s.cols.defaultSize);
      const cs = s.colStyles.get(c);
      const style = cs ? styles.xf(cs) : 0;
      const hidden = s.cols.isHidden(c);
      const level = s.cols.level(c);
      // Varsayılan sütun (genişliği 9, gizli/anahatlı/stilli değil) yazılmaz ve dizi kesilir
      if (width === 9 && !hidden && !level && !cs) {
        run = null;
        continue;
      }
      if (run && run.max === c && run.width === width && run.hidden === hidden && run.level === level && run.style === style) run.max = c + 1;
      else out.push((run = { min: c + 1, max: c + 1, width, style, hidden, level }));
    }
    const boxed = (c: number) => this.ctx.checkbox.size && this.ctx.checkbox.has(s.colStyles.get(c - 1) ?? -1);
    // "Daraltılmış": en derin anahat düzeyindeki sütunlar (eski yazıcıyla aynı kural)
    const maxLevel = s.cols.maxLevel();
    return out
      .map(
        (c) =>
          `<col min="${c.min}" max="${c.max}"${c.width ? ` width="${c.width}"` : ""}${c.style ? ` style="${boxed(c.min) ? styles.boxed(c.style) : c.style}"` : ""}${c.hidden ? ' hidden="1"' : ""}${c.level ? ` outlineLevel="${c.level}"` : ""}${c.level && c.level >= maxLevel ? ' collapsed="1"' : ""} customWidth="1"/>`,
      )
      .join("");
  }

  private viewXml(): string {
    const s = this.sheet;
    const active = addr(s.selection.active.row, s.selection.active.col);
    let a = ' workbookViewId="0"';
    if (s.showHeadings === false) a += ' showRowColHeaders="0"';
    if (s.showGridlines === false) a += ' showGridLines="0"';
    if (s.zoom) a += ` zoomScale="${s.zoom}"`;
    if (s.viewMode === "pageLayout") a += ' view="pageLayout"';
    else if (s.viewMode === "pageBreak") a += ' view="pageBreakPreview"';
    let inner: string;
    if (s.frozen.rows || s.frozen.cols) {
      const pane = s.frozen.cols && s.frozen.rows ? "bottomRight" : s.frozen.cols ? "topRight" : "bottomLeft";
      inner = `<pane${attrs([["xSplit", s.frozen.cols || undefined], ["ySplit", s.frozen.rows || undefined], ["topLeftCell", addr(s.frozen.rows, s.frozen.cols)], ["activePane", pane], ["state", "frozen"]])}/><selection pane="${pane}" activeCell="${active}" sqref="${active}"/>`;
    } else inner = `<selection activeCell="${active}" sqref="${active}"/>`;
    return `<sheetViews><sheetView${a}>${inner}</sheetView></sheetViews>`;
  }

  private formatPrXml(): string {
    const s = this.sheet;
    const h = pxToPt(s.rows.defaultSize);
    let a = ` defaultRowHeight="${h}" outlineLevelRow="${s.rows.maxLevel() || 0}" outlineLevelCol="${s.cols.maxLevel() || 0}" x14ac:dyDescent="55"`;
    if (Math.abs(s.cols.defaultSize - 64) > 0.5) {
      const w = widthPxToChars(s.cols.defaultSize);
      if (w) a += ` defaultColWidth="${w}"`;
    }
    if (h !== 15) a += ' customHeight="1"';
    return `<sheetFormatPr${a}/>`;
  }

  /** Koşullu biçimler: her kural ayrı conditionalFormatting öğesinde; uzantı (x14) gerektirenler ayrıca. */
  private condFormatsXml(): { main: string; ext: string } {
    const s = this.sheet;
    const keep = passthroughCfRules(s);
    let priority = 1;
    let main = "";
    let ext = "";
    for (const rule of s.condFormats) {
      if (keep.has(rule.id)) continue;
      const xr = writeCondRule(rule, priority++);
      if (!xr) continue;
      const sqref = rule.ranges.map((r) => formatRange(r)).join(" ");
      const dxf = xr.style ? this.ctx.styles.dxf(xr.style) : undefined;
      const isExtIcon = xr.type === "iconSet" && EXT_ICONS.has(xr.iconSet);
      if (isExtIcon) {
        const icon = `<x14:iconSet${attrs([["iconSet", xr.iconSet], ["reverse", xr.reverse ? "1" : undefined], ["showValue", xr.showValue === false ? "0" : undefined]])}>${xr.cfvo.map((c: any) => `<x14:cfvo type="${esc(String(c.type))}">${c.value !== undefined ? `<xm:f>${esc(String(c.value))}</xm:f>` : ""}</x14:cfvo>`).join("")}</x14:iconSet>`;
        ext += `<x14:conditionalFormatting xmlns:xm="${NS_XM}"><x14:cfRule type="iconSet" priority="${xr.priority}" id="${guidUpper()}">${icon}</x14:cfRule><xm:sqref>${esc(sqref)}</xm:sqref></x14:conditionalFormatting>`;
        continue;
      }
      const cfRule = this.cfRuleXml(rule, xr, dxf);
      if (cfRule) main += `<conditionalFormatting sqref="${esc(sqref)}">${cfRule}</conditionalFormatting>`;
    }
    return { main, ext: ext ? `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${NS_X14}"><x14:conditionalFormattings>${ext}</x14:conditionalFormattings></ext></extLst>` : "" };
  }

  private cfRuleXml(rule: CondRule, xr: any, dxfId: number | undefined): string {
    const base = (type: string, extra: [string, string | number | undefined][] = []) => attrs([["type", type], ["dxfId", dxfId], ["priority", xr.priority], ...extra]);
    const tl = topLeft(rule);
    switch (xr.type) {
      case "expression":
        return `<cfRule${base("expression")}>${formulaXml(xr.formulae[0] ?? "")}</cfRule>`;
      case "cellIs":
        return `<cfRule${base("cellIs", [["operator", xr.operator]])}>${xr.formulae.map(formulaXml).join("")}</cfRule>`;
      case "top10":
        return `<cfRule${base("top10", [["percent", xr.percent ? "1" : undefined], ["bottom", xr.bottom ? "1" : undefined], ["rank", parseInt(String(xr.rank), 10)]])}/>`;
      case "aboveAverage":
        return `<cfRule${base("aboveAverage", [["aboveAverage", xr.aboveAverage === false ? "0" : undefined]])}/>`;
      case "dataBar":
        return `<cfRule type="dataBar" priority="${xr.priority}"><dataBar>${xr.cfvo.map(cfvoXml).join("")}${colorRgb("color", xr.color)}</dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${NS_X14}"><x14:id/></ext></extLst></cfRule>`;
      case "colorScale":
        return `<cfRule type="colorScale" priority="${xr.priority}"><colorScale>${xr.cfvo.map(cfvoXml).join("")}${xr.color.map((c: any) => colorRgb("color", c)).join("")}</colorScale></cfRule>`;
      case "iconSet":
        return `<cfRule type="iconSet" priority="${xr.priority}"><iconSet${attrs([["iconSet", xr.iconSet === "3TrafficLights" ? undefined : xr.iconSet], ["reverse", xr.reverse ? "1" : undefined], ["showValue", xr.showValue === false ? "0" : undefined]])}>${xr.cfvo.map(cfvoXml).join("")}</iconSet></cfRule>`;
      case "containsText": {
        const t = textRule(xr.operator, xr.text, tl);
        if (!t) return "";
        return `<cfRule${base(t.type, [["operator", t.operator], ["text", t.text]])}>${formulaXml(t.formula)}</cfRule>`;
      }
      case "timePeriod": {
        const f = timeFormula(xr.timePeriod, tl);
        return `<cfRule${base("timePeriod", [["timePeriod", xr.timePeriod]])}>${f ? formulaXml(f) : ""}</cfRule>`;
      }
      case "duplicateValues":
      case "uniqueValues":
        return `<cfRule${base(xr.type)}/>`;
    }
    return "";
  }

  private buildFrame(tablePlans: TablePlan[]): void {
    const { ctx, sheet: s } = this;
    // İlişkiler eski yazıcının sırasıyla: köprüler, notlar (+ VML), çizim, tablolar
    const links = this.linkCells();
    const hyperlinks = links.map(([k, target]) => {
      const rid = this.rels.add(REL.hyperlink, target, true);
      const internal = /^[^!]+![a-zA-Z]+[\d]+$/.test(target);
      return `<hyperlink ref="${addr(keyRow(k), keyCol(k))}" r:id="${rid}"${internal ? ` location="${esc(target)}"` : ""}/>`;
    });
    let legacy = "";
    if (s.notes.size) {
      const notes: NoteOut[] = [...s.notes.entries()].sort((a, b) => a[0] - b[0]).map(([k, n]) => ({ row: keyRow(k), col: keyCol(k), text: n.text }));
      const id = this.index + 1;
      this.rels.add(REL.comments, `../comments${id}.xml`);
      const vrid = this.rels.add(REL.vmlDrawing, `../drawings/vmlDrawing${id}.vml`);
      this.comments = { path: `xl/comments${id}.xml`, vmlPath: `xl/drawings/vmlDrawing${id}.vml`, xml: commentsXml(notes), vml: vmlXml(notes) };
      legacy = `<legacyDrawing r:id="${vrid}"/>`;
    }
    let drawing = "";
    if (s.images.length) {
      const d = buildDrawing(s, ctx.drawings + 1, ctx.media);
      if (d) {
        ctx.drawings++;
        this.drawing = d;
        drawing = `<drawing r:id="${this.rels.add(REL.drawing, `../drawings/drawing${ctx.drawings}.xml`)}"/>`;
      }
    }
    const tableParts: string[] = [];
    for (const p of tablePlans) {
      p.id = ++ctx.tables;
      p.path = `xl/tables/table${p.id}.xml`;
      this.tables.push({ path: p.path, xml: tableXml(p) });
      tableParts.push(`<tablePart r:id="${this.rels.add(REL.table, `../tables/table${p.id}.xml`)}"/>`);
    }

    const fit = fitOf(s.page);
    const sheetPr = (s.tabColor ? `<tabColor rgb="FF${esc(s.tabColor.slice(1))}"/>` : "") + (fit ? '<pageSetUpPr fitToPage="1"/>' : "");
    const cols = this.colsXml();
    const cf = this.condFormatsXml();
    const xml =
      XML_DECL +
      `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" mc:Ignorable="x14ac" xmlns:x14ac="${NS_X14AC}">` +
      (sheetPr ? `<sheetPr>${sheetPr}</sheetPr>` : "") +
      `<dimension ref="${this.dimension()}"/>` +
      this.viewXml() +
      this.formatPrXml() +
      (cols ? `<cols>${cols}</cols>` : "") +
      ROWS_PLACEHOLDER +
      (s.protected ? '<sheetProtection sheet="1"/>' : "") +
      (s.autoFilter ? `<autoFilter ref="${esc(formatRange(s.autoFilter.range))}"/>` : "") +
      cf.main +
      (hyperlinks.length ? `<hyperlinks>${hyperlinks.join("")}</hyperlinks>` : "") +
      pageXml(s.page) +
      drawing +
      (tableParts.length ? `<tableParts count="${tableParts.length}">${tableParts.join("")}</tableParts>` : "") +
      cf.ext +
      legacy +
      "</worksheet>";
    this.frame = new XmlPart(xml, WORKSHEET_ORDER);
  }

  /* ------------------------------------------------------------------ 2) akış */

  /** Hücrenin XML'i (yazılmayacaksa ""). Satırın hücre aralığı (spans) için var olduğu bilgisi ayrıca tutulur. */
  private cellXml(k: number, rec: CellRecord | undefined, e: Extra | undefined, ref: string): string {
    const { ctx } = this;
    let xf = this.baseXf(k, rec, e);
    let body = "";
    let t = "";
    let array = false;
    if (e?.clear === undefined) {
      if (rec) {
        if (rec.formula) {
          const { f, dynamic } = fileFormula(rec.formula.text);
          const sp = rec.formula.spill ? ctx.wb.engine.getSpill(this.sheet.id, keyRow(k), keyCol(k)) : undefined;
          const r = keyRow(k);
          const c = keyCol(k);
          let fa = "";
          if (sp && sp.anchorRow === r && sp.anchorCol === c && (sp.rows > 1 || sp.cols > 1)) {
            fa = ` t="array" ref="${formatRange({ r1: r, c1: c, r2: r + sp.rows - 1, c2: c + sp.cols - 1 })}"`;
            array = true;
          } else if (dynamic) {
            fa = ` t="array" ref="${ref}"`;
            array = true;
          }
          const v = rec.value;
          body = `<f${fa}>${esc(f)}</f>`;
          if (v === null) {
            /* önbelleksiz formül */
          } else if (typeof v === "number") body += `<v>${v}</v>`;
          else if (typeof v === "string") {
            t = ' t="str"';
            body += `<v>${esc(v)}</v>`;
          } else if (typeof v === "boolean") {
            t = ' t="b"';
            body += `<v>${v ? 1 : 0}</v>`;
          } else if (isError(v)) {
            t = ' t="e"';
            body += `<v>${esc(v.code)}</v>`;
          }
        } else ({ t, body } = this.scalarXml(rec.value));
      } else if (e?.v !== undefined && e.v !== null) {
        const v = e.v;
        if (typeof v === "object" && !isError(v)) {
          if ("f" in v) body = `<f>${esc(v.f)}</f>`;
          else {
            t = ' t="s"';
            body = `<v>${ctx.sst.add(v.link)}</v>`;
          }
        } else ({ t, body } = this.scalarXml(v));
      }
    }
    if (!body && !xf) return "";
    // Onay kutusu: hücrenin etkin model stili onay kutulu ise biçimin onay kutulu kopyası
    if (ctx.checkbox.size) {
      const id = this.sheet.styles.get(k) ?? this.sheet.rowStyles.get(keyRow(k)) ?? this.sheet.colStyles.get(keyCol(k)) ?? 0;
      if (ctx.checkbox.has(id)) xf = ctx.styles.boxed(xf);
    }
    const sAttr = xf ? ` s="${xf}"` : "";
    const dts = this.dataTables.size ? this.dataTables.get(k) : undefined;
    const cm = array && !dts ? ' cm="1"' : "";
    if (array && !dts) ctx.dynamic = true;
    const xml = body ? `<c r="${ref}"${sAttr}${t}${cm}>${body}</c>` : `<c r="${ref}"${sAttr}/>`;
    return dts ? injectDataTables(xml, dts) : xml;
  }

  private scalarXml(v: Scalar): { t: string; body: string } {
    if (v === null) return { t: "", body: "" };
    if (typeof v === "number") return { t: "", body: `<v>${v}</v>` };
    if (typeof v === "string") return { t: ' t="s"', body: `<v>${this.ctx.sst.add(v)}</v>` };
    if (typeof v === "boolean") return { t: ' t="b"', body: `<v>${v ? 1 : 0}</v>` };
    return { t: ' t="e"', body: `<v>${esc(v.code)}</v>` };
  }

  /**
   * Satır açılış etiketi. Hücresi ya da (sıfırdan büyük) özel yüksekliği olan satır biçimini (s) de
   * taşır; yalnız gizli/anahatlı/daraltılmış ya da sıfır yükseklikli boş satır yalın yazılır.
   */
  private rowOpen(r: number, spans: string | null): string {
    const s = this.sheet;
    let a = `<row r="${r + 1}"`;
    const ht = s.rows.isCustom(r) ? pxToPt(s.rows.rawSize(r)) : undefined;
    if (ht !== undefined) a += ` ht="${ht}"${s.rows.isManual(r) ? ' customHeight="1"' : ""}`;
    if (s.rows.isHidden(r)) a += ' hidden="1"';
    const lv = s.rows.level(r);
    if (spans === null && !ht) return lv ? `${a} outlineLevel="${lv}"` : a;
    if (spans) a += spans;
    const rs = s.rowStyles.get(r);
    if (rs) {
      let xf = this.ctx.styles.xf(rs);
      if (this.ctx.checkbox.has(rs)) xf = this.ctx.styles.boxed(xf);
      if (xf) a += ` s="${xf}" customFormat="1"`;
    }
    a += ' x14ac:dyDescent="0.25"';
    if (lv) a += ` outlineLevel="${lv}"`;
    // "Daraltılmış": en derin anahat düzeyindeki satırlar (eski yazıcıyla aynı kural)
    if (lv && lv >= this.rowMaxLevel) a += ' collapsed="1"';
    return a;
  }

  /** Sayfanın tamamını (iskelet + satırlar) alıcıya akıtır. */
  async stream(sink: DeflateSink): Promise<void> {
    const whole = this.frame.toString();
    const at = whole.indexOf(ROWS_PLACEHOLDER);
    if (at < 0) throw new Error("Sayfa iskeletinde satır yer tutucusu bulunamadı");
    const o = new ByteOut(sink);
    o.text(whole.slice(0, at));
    const { ctx, sheet } = this;
    const R = this.recKeys;
    const S = this.styleKeys;
    const X = this.extraKeys;
    const recs = this.recs;
    const engine = ctx.wb.engine;
    const sid = sheet.id;
    const meta = this.metaRows;
    const any = R.length + S.length + X.length + meta.length > 0;
    o.ascii(any ? "<sheetData>" : "<sheetData/>");
    // Hızlı yol: yalnız kaydı olan sabit değerli hücre (onay kutusu ve veri tablosu yoksa)
    const fast = !ctx.checkbox.size && !this.dataTables.size;
    const rowSt = sheet.rowStyles.size ? sheet.rowStyles : null;
    const colSt = sheet.colStyles.size ? sheet.colStyles : null;
    const { styles, sst } = ctx;
    let mi = 0;
    let ri = 0;
    let si = 0;
    let xi = 0;
    for (;;) {
      const k0 = Math.min(ri < R.length ? R[ri] : Infinity, si < S.length ? S[si] : Infinity, xi < X.length ? X[xi] : Infinity);
      if (k0 === Infinity) break;
      const r = Math.floor(k0 / MAX_COLS);
      while (mi < meta.length && meta[mi] < r) {
        o.ascii(this.rowOpen(meta[mi++], null));
        o.bytes(B_SLASH_CLOSE);
      }
      if (mi < meta.length && meta[mi] === r) mi++;
      // Satırın son hücresi (spans): her dizide satır sonuna kadar ilerlenir
      const end = (r + 1) * MAX_COLS;
      let re = ri;
      while (re < R.length && R[re] < end) re++;
      let se = si;
      while (se < S.length && S[se] < end) se++;
      let xe = xi;
      while (xe < X.length && X[xe] < end) xe++;
      const last = Math.max(re > ri ? R[re - 1] : -1, se > si ? S[se - 1] : -1, xe > xi ? X[xe - 1] : -1);
      const base = r * MAX_COLS;
      o.ascii(this.rowOpen(r, ` spans="${k0 - base + 1}:${last - base + 1}"`));
      o.byte(62);
      const open = o.pos;
      const rowStr = String(r + 1);
      const rd = asciiBytes(rowStr);
      const rowId = rowSt?.get(r);
      for (;;) {
        const kr = ri < re ? R[ri] : Infinity;
        const ks = si < se ? S[si] : Infinity;
        const kx = xi < xe ? X[xi] : Infinity;
        const k = kr < ks ? (kr < kx ? kr : kx) : ks < kx ? ks : kx;
        if (k === Infinity) break;
        const c = k - base;
        if (kr === k) {
          const rec = recs ? recs[ri] : engine.getRecord(sid, r, c);
          ri++;
          if (ks !== k && kx !== k && fast && rec && !rec.formula) {
            // Kayıtlı sabit hücre: biçim satır, yoksa sütun stilinden; ara metin üretilmez
            const v = rec.value;
            const id = rowId ?? colSt?.get(c) ?? 0;
            const xf = id ? styles.xf(id) : 0;
            if (v === null && !xf) continue;
            o.bytes(B_CELL);
            o.bytes(colBytes(c));
            o.bytes(rd);
            o.byte(34);
            if (xf) {
              o.bytes(B_S);
              o.int(xf);
              o.byte(34);
            }
            if (typeof v === "number") {
              o.bytes(B_V);
              o.num(v);
              o.bytes(B_V_END);
            } else if (typeof v === "string") {
              o.bytes(B_TS);
              o.int(sst.add(v));
              o.bytes(B_V_END);
            } else if (typeof v === "boolean") {
              o.bytes(B_TB);
              o.byte(v ? 49 : 48);
              o.bytes(B_V_END);
            } else if (v !== null) {
              o.bytes(B_TE);
              o.ascii(v.code);
              o.bytes(B_V_END);
            } else o.bytes(B_SLASH_CLOSE);
            continue;
          }
          if (ks === k) si++;
          const e = kx === k ? this.extras.get(k) : undefined;
          if (kx === k) xi++;
          const xml = this.cellXml(k, rec, e, colName(c) + rowStr);
          if (xml) o.text(xml);
          continue;
        }
        if (ks === k) si++;
        const e = kx === k ? this.extras.get(k) : undefined;
        if (kx === k) xi++;
        const xml = this.cellXml(k, undefined, e, colName(c) + rowStr);
        if (xml) o.text(xml);
      }
      if (o.pos === open) {
        o.pos--;
        o.bytes(B_SLASH_CLOSE);
      } else o.bytes(B_ROW_END);
      await o.flush();
    }
    while (mi < meta.length) {
      o.ascii(this.rowOpen(meta[mi++], null));
      o.bytes(B_SLASH_CLOSE);
    }
    if (any) o.ascii("</sheetData>");
    o.text(whole.slice(at + ROWS_PLACEHOLDER.length));
    await o.flush(true);
  }
}
