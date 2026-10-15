/**
 * Uygulama denetleyicisi: çalışma kitabı modeli, seçim, düzenleme oturumu,
 * pano ve görünüm durumu. React bileşenleri `version` üzerinden abone olur.
 */
import { addr, colToName, findRefSpans, parseLocalFormula, TR_SYNTAX, quoteSheetName, cycleAbsolute, renderRefToken, tokenize, shiftFormula, MAX_COLS, MAX_ROWS } from "@officehere/formula";
import type { FileRef } from "@officehere/core";
import { CondFormatter } from "./model/condfmt";
import { Ops, type FillMode } from "./model/ops";
import { intersects as intersectsRect, rect, rectToA1 } from "./model/rect";
import { isPivotCell, PIVOT_EDIT_MESSAGE } from "./pivot/pivotOps";
import { expandForMerges, extendTo, jumpCell, nextInSelection, singleSelection, stepCell } from "./model/selection";
import type { CellPos, Rect, Selection } from "./model/types";
import { WorkbookModel, type SheetModel } from "./model/workbook";
import { autoCloseParens, parseRangeLoose, validateInputRule, validationList } from "./model/inputRules";
import { adjustRowHeights, widenForNumber } from "./grid/autofit";
import type { TraceArrow, Watch } from "./actions/audit";
import { activateTheme } from "./model/themeRuntime";
import { afterTableEntry, tableTab } from "./actions/tableAuto";
import { filtersOf, recomputeFilters } from "./actions/data";
import { structuredPointRef } from "./model/structPoint";
import { dataTableAt, refreshDataTables } from "./model/whatIf";

export const DATA_TABLE_MESSAGE = "Veri tablosunun bir bölümü değiştirilemez. Tablonun tamamını seçip silebilirsiniz.";

export type EditMode = "enter" | "edit";

export interface EditState {
  sheetId: number;
  row: number;
  col: number;
  text: string;
  mode: EditMode;
  /** Düzenleme formül çubuğunda mı başladı. */
  inBar: boolean;
  caret: number;
  /** İşaretle (fare/ok) eklenen son başvurunun konumu. */
  pointSpan?: { start: number; end: number };
  /** İşaretlenen başvurunun bağlantı ve köşe noktası. */
  pointAnchor?: CellPos;
  pointCorner?: CellPos;
  pointSheetId?: number;
}

export interface ClipData {
  sheetId: number;
  rect: Rect;
  cut: boolean;
}

export const REF_COLORS = ["#3F6FD8", "#E2482F", "#8B3FC1", "#2E9E44", "#C55A11", "#0E9ED5", "#B8327F", "#7A6A00"];

export class Controller {
  wb: WorkbookModel;
  ops: Ops;
  cf: CondFormatter;
  version = 0;
  private listeners = new Set<() => void>();
  fileName = "Kitap1";
  fileRef: FileRef | null = null;
  dirty = false;
  /** Belge açıldığından beri dokunulmamış mı (masaüstünde aynı pencerede açmak için). */
  pristine = true;
  edit: EditState | null = null;
  clip: ClipData | null = null;
  /** Shift ile genişletmede hareket eden köşe. */
  corner: CellPos = { row: 0, col: 0 };
  /** Sekme (Tab) dizisinin başladığı sütun: ardından Enter bu sütunun altına döner. */
  tabStartCol: number | null = null;
  /** Kitap içeriği her değiştiğinde artar (seçim/düzenleme değişiminde artmaz); önbellekler için. */
  dataVersion = 0;
  showFormulaBar = true;
  /** Sayfa Düzeni görünümünde cetveller (Görünüm > Göster > Cetvel). */
  showRuler = true;
  formulaBarExpanded = false;
  /** Durum çubuğu iletisi (ör. "Hedefi seçin ve ENTER tuşuna basın"). */
  status = "";
  /** Gözden Geçir > Tüm Notları Göster. */
  showAllNotes = false;
  /** PivotTable Alanları bölmesi (PivotTable içindeyken). */
  showFieldList = true;
  /** Formül Denetimi izleme okları (düzenleme yapılınca silinir). */
  arrows: TraceArrow[] = [];
  /** Gözden Geçir > Yazım Denetimi: düzenleyicide tarayıcının yazım denetimi (Türkçe). */
  spellcheck = false;
  /** Arka planda hata denetimi (Formüller > Hata Denetimi seçenekleri). */
  errorChecking = true;
  /** İzleme Penceresi. */
  watches: Watch[] = [];
  showWatch = false;
  /** Şeridin geçmesi istenen sekme (ör. PivotTable oluşturulunca "Özet Tablo Analizi"). */
  requestTab: string | null = null;
  /**
   * Son doldurma tutamacı işlemi (Otomatik Doldurma Seçenekleri düğmesi için).
   * undo: işlemden sonraki geri al sayısı; başka bir değişiklik olunca düğme kaybolur.
   */
  fillOptions: { sheetId: number; src: Rect; target: Rect; mode: FillMode; undo: number } | null = null;
  /** Veri doğrulama açılan listesi açık mı. */
  listDropdown = false;
  /** Uygulamanın Kaydet eylemi (çökme uyarısından çağrılır). */
  saveHook: (() => Promise<boolean>) | null = null;
  /** Biçimlendirme bölmesinde seçili grafik öğesi ("area", "plot", "title", "legend", "axis:val", "axis:cat", "series:0", "labels"). */
  chartElement = "area";
  /** Sağdaki biçimlendirme bölmesi (Grafik Alanını Biçimlendir). */
  formatPane: "chart" | "selection" | null = null;
  /** Son yinelenebilir işlem (F4 / Ctrl+Y: Yinele). */
  lastAction: { label: string; run: () => void } | null = null;
  /** Şeridin altındaki bilgi çubuğu (ör. içe aktarmada atlanan öğeler). */
  infoBar: { kind: "info" | "warning" | "error"; text: string } | null = null;
  /** Seçili yüzen nesne (grafik, resim, şekil). */
  selObj: { kind: "chart" | "image" | "shape"; id: string } | null = null;
  /** Yeni Pencere ile bağlı pencereler (aynı kitap). */
  windowLink: { doc: string; index: number; peers: Map<string, number> } | null = null;
  /** Seçili dilimleyici. */
  selSlicer: string | null = null;
  /** Çizim kipi: sayfada sürükleyerek çizilecek şekil (Metin Kutusu). */
  drawShape: "textbox" | null = null;
  /** Şekil içinde yerinde metin düzenleme (şekil seçiliyken yazınca başlar). */
  shapeEdit: { id: string; text: string } | null = null;
  /** Biçim Boyacısı: kopyalanan stiller. */
  painter: { sheetId: number; rect: Rect; sticky: boolean } | null = null;
  /** Odak ızgaraya geri verilsin diye çağrılır. */
  focusGrid: () => void = () => {};
  /** Izgarada bir hücreyi görünür yap. */
  scrollIntoView: (row: number, col: number) => void = () => {};
  /** Hücrenin ekran (istemci) dikdörtgeni; ızgara çizilmemişse null. */
  cellClientRect: (row: number, col: number) => { x: number; y: number; w: number; h: number } | null = () => null;
  /** İletişim kutusu açma isteği (App tarafından ayarlanır). */
  openDialog: (id: string, arg?: unknown) => void = () => {};
  /** Uyarı/hata mesajı göster. */
  /** detail: teknik ayrıntı (iletişim kutusunda "Ayrıntılar" altında, kapalı). */
  alert: (msg: string, title?: string, detail?: string) => void = () => {};

  constructor(wb = new WorkbookModel()) {
    this.wb = wb;
    this.ops = this.makeOps(wb);
    this.cf = new CondFormatter(wb);
    activateTheme(wb.theme);
  }

  private makeOps(wb: WorkbookModel): Ops {
    const ops = new Ops(wb);
    ops.onChange = () => {
      this.dataVersion++;
      this.dirty = true;
      this.pristine = false;
      // Veri tabloları (TABLE) bağımlı oldukları değerler değişince yeniden hesaplanır
      if (wb.sheets.some((s) => s.dataTables.length)) this.scheduleDataTables();
      // Tema geri al/yinele ile değişmiş olabilir
      activateTheme(this.wb.theme);
      // Excel gibi: çalışma sayfası değişince izleme okları kalkar
      if (this.arrows.length) this.arrows = [];
      this.emit();
    };
    ops.onBlocked = (m) => this.alert(m);
    // Tarayıcıda ölçüm yapılabildiğinde otomatik satır yüksekliği
    if (typeof document !== "undefined") ops.afterCells = (sheet, rects) => adjustRowHeights(wb, sheet, rects);
    return ops;
  }

  private dtTimer: ReturnType<typeof setTimeout> | null = null;

  /** Veri tablolarının yeniden hesaplanması (değişiklikten hemen sonra, tek seferde). */
  scheduleDataTables(): void {
    if (this.dtTimer) return;
    this.dtTimer = setTimeout(() => {
      this.dtTimer = null;
      if (refreshDataTables(this.wb)) {
        this.dataVersion++;
        this.emit();
      }
    }, 0);
  }

  /** Yeni çalışma kitabı yükler. */
  load(wb: WorkbookModel, fileName: string, ref: FileRef | null): void {
    this.wb = wb;
    this.dataVersion++;
    this.ops = this.makeOps(wb);
    this.cf = new CondFormatter(wb);
    activateTheme(wb.theme);
    wb.syncTables();
    wb.engine.recalc();
    wb.sheets.forEach((s) => this.ops.syncHiddenRows(s));
    // Dosyadan gelen filtre ölçütleri (ve Gelişmiş Filtre) yeniden uygulanır: gizli satırlar filtre gizliliği olur
    wb.sheets.forEach((s) => {
      if (s.advFilter || filtersOf(s).some((f) => Object.keys(f.columns).length)) recomputeFilters(this, s);
    });
    this.fileName = fileName;
    this.fileRef = ref;
    this.dirty = false;
    this.infoBar = wb.loadWarnings.length
      ? { kind: "warning", text: `Dosyadaki bazı öğeler okunamadı ve atlandı: ${wb.loadWarnings.join(", ")}. Çalışma kitabının geri kalanı açıldı.` }
      : null;
    this.edit = null;
    this.clip = null;
    this.corner = { ...this.sel.active };
    this.emit();
  }

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getVersion = (): number => this.version;

  emit(): void {
    this.version++;
    for (const l of this.listeners) l();
  }

  get sheet(): SheetModel {
    return this.wb.active;
  }

  get sel(): Selection {
    return this.sheet.selection;
  }

  /** Etkin aralık (son seçilen). */
  get range(): Rect {
    const s = this.sel;
    return s.ranges[s.ranges.length - 1];
  }

  get mode(): "ready" | "enter" | "edit" | "point" {
    if (!this.edit) return "ready";
    if (this.edit.pointSpan) return "point";
    return this.edit.mode;
  }

  /* ------------------------------ Seçim ------------------------------ */

  setSelection(sel: Selection, scroll = true): void {
    this.selObj = null;
    this.tabStartCol = null;
    this.sheet.selection = sel;
    if (scroll) this.scrollIntoView(this.corner.row, this.corner.col);
    this.emit();
  }

  selectCell(row: number, col: number, scroll = true): void {
    const s = singleSelection(this.sheet, row, col);
    this.selObj = null;
    this.tabStartCol = null;
    // Filtre sonucu iletisi bir sonraki gezinmede silinir
    if (this.status.endsWith("tanesi bulundu")) this.status = "";
    this.corner = { row, col };
    this.sheet.selection = s;
    if (scroll) this.scrollIntoView(row, col);
    this.emit();
  }

  selectRange(r: Rect, active?: CellPos): void {
    const rr = expandForMerges(this.sheet, r);
    const a = active ?? { row: rr.r1, col: rr.c1 };
    this.corner = { row: rr.r2, col: rr.c2 };
    this.setSelection({ ranges: [rr], active: a, anchor: a }, false);
  }

  /** Ok tuşları: hareket ya da Shift ile genişletme. */
  move(dr: number, dc: number, extend: boolean, jump: boolean): void {
    const sh = this.sheet;
    // Sayfa Aşağı/Yukarı gibi çok adımlı hareketler gizli satırları atlayarak adım adım
    const step = (from: CellPos): CellPos => {
      let p = from;
      const n = Math.max(Math.abs(dr), Math.abs(dc));
      if (n <= 1) return stepCell(sh, from, dr, dc);
      for (let i = 0; i < n; i++) {
        const q = stepCell(sh, p, Math.sign(dr), Math.sign(dc));
        if (q.row === p.row && q.col === p.col) break;
        p = q;
      }
      return p;
    };
    if (extend) {
      const from = this.corner;
      const next = jump ? jumpCell(this.wb, sh, from, dr, dc) : step(from);
      this.corner = next;
      this.sheet.selection = extendTo(sh, this.sel, next);
      this.scrollIntoView(next.row, next.col);
      this.emit();
      return;
    }
    const next = jump ? jumpCell(this.wb, sh, this.sel.active, dr, dc) : step(this.sel.active);
    this.selectCell(next.row, next.col);
  }

  /** Enter/Tab sonrası seçim içinde ilerleme. */
  advance(dir: "down" | "up" | "right" | "left"): void {
    // Tablo içinde Tab: satır sonunda sonraki satıra, son hücrede yeni tablo satırına
    if ((dir === "right" || dir === "left") && this.sheet.tables.length && tableTab(this, dir === "left")) {
      this.tabStartCol = null;
      return;
    }
    const single = this.sel.ranges.length === 1 && rectIsCell(this.sheet, this.range);
    if (single && (dir === "right" || dir === "left")) {
      if (this.tabStartCol === null && dir === "right") this.tabStartCol = this.sel.active.col;
    } else if (single && dir === "down" && this.tabStartCol !== null) {
      // Tab ile satırda ilerleyip Enter'a basınca başlangıç sütununun altına dön
      const col = this.tabStartCol;
      const next = stepCell(this.sheet, { row: this.sel.active.row, col }, 1, 0);
      this.selectCell(next.row, col);
      return;
    } else this.tabStartCol = null;
    const s = nextInSelection(this.sheet, this.sel, dir);
    this.corner = { ...s.active };
    this.sheet.selection = s;
    this.scrollIntoView(s.active.row, s.active.col);
    this.emit();
  }

  selectAll(): void {
    this.selectRange(rect(0, 0, MAX_ROWS - 1, MAX_COLS - 1), this.sel.active);
  }

  selectColumns(c1: number, c2: number, add = false): void {
    const r = rect(0, Math.min(c1, c2), MAX_ROWS - 1, Math.max(c1, c2));
    const active = { row: this.sheet.scroll.row, col: Math.min(c1, c2) };
    this.applyHeaderSelection(r, active, add);
  }

  selectRows(r1: number, r2: number, add = false): void {
    const r = rect(Math.min(r1, r2), 0, Math.max(r1, r2), MAX_COLS - 1);
    const active = { row: Math.min(r1, r2), col: this.sheet.scroll.col };
    this.applyHeaderSelection(r, active, add);
  }

  private applyHeaderSelection(r: Rect, active: CellPos, add: boolean): void {
    const sel = add ? { ranges: [...this.sel.ranges, r], active, anchor: active } : { ranges: [r], active, anchor: active };
    this.corner = { row: r.r2, col: r.c2 };
    this.setSelection(sel, false);
  }

  /** Ad Kutusu / Git: başvuru ya da ad metnine git. */
  goTo(text: string): boolean {
    const t = text.trim();
    if (!t) return false;
    const wb = this.wb;
    let sheet: SheetModel | undefined = this.sheet;
    let body = t;
    const bang = t.lastIndexOf("!");
    if (bang > 0) {
      let name = t.slice(0, bang);
      if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
      sheet = wb.sheetByName(name);
      body = t.slice(bang + 1);
    }
    if (!sheet) return false;
    let r = parseRangeLoose(body);
    if (!r) {
      // Tanımlı ad
      const def = wb.engine.getNames().find((n) => n.name.toLocaleUpperCase("tr-TR") === t.toLocaleUpperCase("tr-TR"));
      if (def) return this.goTo(def.formula.slice(1));
      const table = wb.sheets.flatMap((s) => s.tables.map((tb) => ({ s, tb }))).find((x) => x.tb.name.toLocaleUpperCase("tr-TR") === t.toLocaleUpperCase("tr-TR"));
      if (table) {
        sheet = table.s;
        r = table.tb.ref;
      }
    }
    if (!r) return false;
    if (sheet !== this.sheet) this.activateSheet(wb.sheets.indexOf(sheet));
    this.selectRange(r);
    this.scrollIntoView(r.r1, r.c1);
    return true;
  }

  activateSheet(index: number): void {
    if (index < 0 || index >= this.wb.sheets.length) return;
    if (this.edit && !this.edit.text.startsWith("=")) this.commitEdit(null);
    this.wb.activeIndex = index;
    this.corner = { ...this.sel.active };
    this.emit();
  }

  /** Ad Kutusu metni. */
  nameBoxText(): string {
    const r = this.range;
    const def = this.wb.engine.getNames().find((n) => {
      const body = n.formula.slice(1);
      const bang = body.lastIndexOf("!");
      const sheetPart = bang > 0 ? body.slice(0, bang).replace(/^'|'$/g, "") : this.sheet.name;
      return sheetPart.toLocaleUpperCase("tr-TR") === this.sheet.name.toLocaleUpperCase("tr-TR") && rectToA1(parseRangeLoose(body.slice(bang + 1)) ?? rect(-1, -1)) === rectToA1(r);
    });
    if (def) return def.name;
    return addr(this.sel.active.row, this.sel.active.col);
  }

  /* ------------------------------ Düzenleme ------------------------------ */

  /** Korumalı sayfada kilitli hücre mi? */
  isLocked(row: number, col: number): boolean {
    return this.sheet.protected && this.wb.styleAt(this.sheet, row, col).locked !== false;
  }

  /** Seçimde kilitli hücre varsa uyarır (korumalı sayfa). */
  guardSelection(): boolean {
    // PivotTable hücreleri yalnız PivotTable işlemleriyle değişir
    if (this.sheet.pivots.length && this.sel.ranges.some((r) => this.sheet.pivots.some((p) => p.output && p.values.length + p.rows.length + p.cols.length > 0 && intersectsRect(p.output, r)))) {
      this.alert(PIVOT_EDIT_MESSAGE);
      return false;
    }
    if (this.sheet.dataTables.length && this.sel.ranges.some((r) => this.sheet.dataTables.some((d) => intersectsRect(d.ref, r) && !(r.r1 <= d.ref.r1 && r.c1 <= d.ref.c1 && r.r2 >= d.ref.r2 && r.c2 >= d.ref.c2)))) {
      this.alert(DATA_TABLE_MESSAGE);
      return false;
    }
    if (!this.sheet.protected) return true;
    for (const r of this.sel.ranges)
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 200); row++)
        for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 100); col++)
          if (this.isLocked(row, col)) {
            this.alert("Değiştirmeye çalıştığınız hücre korumalı bir sayfada. Değişiklik yapmak için Gözden Geçir sekmesinde Sayfa Korumasını Kaldır'ı tıklatın.");
            return false;
          }
    return true;
  }

  /** İşlemi F4 ile yinelenebilir olarak kaydeder (çalıştırıldığında geçerli seçime uygulanır). */
  recordRepeat(label: string, run: () => void): void {
    this.lastAction = { label, run };
  }

  /** F4: son işlemi geçerli seçimde yineler. */
  repeatLast(): boolean {
    const a = this.lastAction;
    if (!a || this.edit) return false;
    a.run();
    return true;
  }

  /** Hücre düzenlemeyi başlatır. text verilirse "Gir" kipinde yazılanla başlar. */
  startEdit(text?: string, inBar = false): void {
    // Salt okunur açılan dosya (görüntüleyici kipi): sarı şerit "Düzenlemeyi etkinleştir"i söyler.
    if (this.edit || this.ops.readOnly) return;
    const { row, col } = this.sel.active;
    if (isPivotCell(this.sheet, row, col)) {
      this.alert(PIVOT_EDIT_MESSAGE);
      return;
    }
    if (this.isLocked(row, col)) {
      this.alert("Değiştirmeye çalıştığınız hücre veya grafik korumalı bir sayfada. Değişiklik yapmak için Gözden Geçir sekmesinde Sayfa Korumasını Kaldır'ı tıklatın.");
      return;
    }
    if (this.sheet.dataTables.length && dataTableAt(this.sheet, row, col)) {
      this.alert(DATA_TABLE_MESSAGE);
      return;
    }
    const initial = text ?? this.wb.editText(this.sheet, row, col);
    this.edit = { sheetId: this.sheet.id, row, col, text: initial, mode: text === undefined ? "edit" : "enter", inBar, caret: initial.length };
    this.status = "";
    this.emit();
  }

  updateEdit(text: string, caret: number): void {
    if (!this.edit) return;
    const keepPoint = this.edit.pointSpan && text === this.edit.text;
    this.edit = { ...this.edit, text, caret, pointSpan: keepPoint ? this.edit.pointSpan : undefined };
    if (!keepPoint) {
      this.edit.pointAnchor = undefined;
      this.edit.pointCorner = undefined;
    }
    this.emit();
  }

  setEditMode(mode: EditMode): void {
    if (!this.edit) return;
    this.edit = { ...this.edit, mode, pointSpan: undefined };
    this.emit();
  }

  /** İmleç başvuru eklenebilecek konumda mı (formül yazarken işaret kipi)? */
  canPoint(): boolean {
    const e = this.edit;
    if (!e || !e.text.startsWith("=")) return false;
    if (e.pointSpan && e.caret === e.pointSpan.end) return true;
    const before = e.text.slice(0, e.caret).replace(/\s+$/, "");
    if (before.length === 0) return false;
    const ch = before[before.length - 1];
    return "=+-*/^&(;<>:,{".includes(ch);
  }

  /** İşaret kipinde başvuru metnini ekler/değiştirir. */
  pointTo(r: Rect, sheet: SheetModel, anchor?: CellPos, corner?: CellPos): void {
    const e = this.edit;
    if (!e) return;
    // Tablo hücreleri yapısal başvuruyla ([@Tutar], Tablo1[Tutar]); diğerleri A1
    let refText = sheet.tables.length ? structuredPointRef(sheet, r, e) : null;
    if (!refText) {
      refText = rectToA1(r);
      if (sheet.id !== e.sheetId) refText = `${quoteSheetName(sheet.name)}!${refText}`;
    }
    let start = e.caret;
    let end = e.caret;
    if (e.pointSpan && e.caret === e.pointSpan.end) {
      start = e.pointSpan.start;
      end = e.pointSpan.end;
    }
    const text = e.text.slice(0, start) + refText + e.text.slice(end);
    const caret = start + refText.length;
    this.edit = { ...e, text, caret, pointSpan: { start, end: caret }, pointAnchor: anchor ?? { row: r.r1, col: r.c1 }, pointCorner: corner ?? { row: r.r2, col: r.c2 }, pointSheetId: sheet.id };
    this.emit();
  }

  /** F4: imleçteki başvurunun $ durumunu döndürür. */
  toggleAbsolute(): void {
    const e = this.edit;
    if (!e || !e.text.startsWith("=")) return;
    const body = e.text.slice(1);
    const toks = tokenize(body, TR_SYNTAX);
    for (const t of toks) {
      if (t.type !== "ref" || !t.ref || t.ref.refError) continue;
      const s = t.pos + 1;
      const en = s + t.text.length;
      if (e.caret >= s && e.caret <= en) {
        const next = renderRefToken(cycleAbsolute(t.ref));
        const text = e.text.slice(0, s) + next + e.text.slice(en);
        this.edit = { ...e, text, caret: s + next.length, pointSpan: e.pointSpan ? { start: s, end: s + next.length } : undefined };
        this.emit();
        return;
      }
    }
  }

  /** Düzenlemeyi uygular. move: sonra seçimin ilerleyeceği yön. fillAll: Ctrl+Enter. */
  commitEdit(move: "down" | "up" | "right" | "left" | null, fillAll = false): boolean {
    const e = this.edit;
    if (!e) return true;
    let text = e.text;
    if (text.startsWith("=") && text.length > 1) {
      text = autoCloseParens(text);
      try {
        parseLocalFormula(text, "tr");
      } catch {
        this.alert("Bu formülle ilgili bir sorun var.\n\nFormül yazmaya çalışmıyor musunuz? İlk karakter eşittir (=) ya da eksi (-) işareti olduğunda, bunun bir formül olduğu düşünülür.", "OfficeHere Tablo");
        return false;
      }
    }
    const sheet = this.wb.sheetById(e.sheetId)!;
    if (sheet !== this.sheet) this.wb.activeIndex = this.wb.sheets.indexOf(sheet);
    // Taşan dizinin parçasına yazma: dizi formülünün bir kısmı değiştirilemez (#SPILL! olur)
    const cells: { row: number; col: number }[] = [];
    if (fillAll) {
      for (const r of this.sel.ranges) for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 100000); i++) for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 1000); j++) cells.push({ row: i, col: j });
    } else cells.push({ row: e.row, col: e.col });
    this.edit = null;
    const validation = this.skipValidation ? null : this.validateInput(sheet, e.row, e.col, text);
    this.skipValidation = false;
    if (validation) {
      // Durdur: Yeniden Dene/İptal; Uyarı: Evet/Hayır/İptal; Bilgi: Tamam/İptal
      this.edit = e;
      this.openDialog("validationError", {
        ...validation,
        onRetry: () => {
          if (this.edit) this.edit = { ...this.edit, mode: "edit" };
          this.emit();
          this.focusGrid();
        },
        onCancel: () => {
          this.cancelEdit();
          this.focusGrid();
        },
        onAccept: () => {
          this.skipValidation = true;
          this.commitEdit(move, fillAll);
          this.focusGrid();
        },
      });
      this.emit();
      return false;
    }
    if (fillAll && cells.length > 1 && text.startsWith("=")) {
      // Ctrl+Enter formülleri göreli kaydırarak doldurur
      this.ops.enterText(sheet, [{ row: e.row, col: e.col }], text, "Yazma");
      const base = this.wb.getFormula(sheet, e.row, e.col);
      if (base) {
        const entries = cells.filter((c) => c.row !== e.row || c.col !== e.col).map((c) => ({ row: c.row, col: c.col, input: { formula: shiftFormula(base, c.row - e.row, c.col - e.col) } }));
        this.ops.setInputs(sheet, entries, "Yazma");
      }
    } else {
      const mark = this.ops.history.mark();
      this.ops.enterText(sheet, cells, text);
      // Biçimli sayı (tarih, para…) varsayılan genişliğe sığmıyorsa sütun genişler
      const w = widenForNumber(this.wb, sheet, e.row, e.col);
      if (w !== null) this.ops.sheetOp(sheet, "Sütun Genişliği", () => sheet.cols.set(e.col, w, false));
      // Tablo: alta/sağa yazınca genişleme, hesaplanmış sütun, başlık adı
      if (text !== "") afterTableEntry(this, sheet, e.row, e.col);
      this.ops.history.squash(mark, "Yazma");
    }
    if (move) this.advance(move);
    else this.emit();
    return true;
  }

  cancelEdit(): void {
    if (!this.edit) return;
    const sheet = this.wb.sheetById(this.edit.sheetId);
    if (sheet && sheet !== this.sheet) this.wb.activeIndex = this.wb.sheets.indexOf(sheet);
    this.edit = null;
    this.emit();
  }

  /** Düzenlenen formüldeki başvurular (renkli vurgular). */
  editRefs(): { rect: Rect; color: string; sheetName?: string }[] {
    const e = this.edit;
    const text = e ? e.text : "";
    if (!text.startsWith("=")) return [];
    const spans = findRefSpans(text, TR_SYNTAX);
    const out: { rect: Rect; color: string; sheetName?: string }[] = [];
    spans.forEach((s, i) => {
      const t = s.ref;
      let r: Rect;
      if (t.kind === "cell") r = rect(t.a.row, t.a.col);
      else if (t.kind === "range") r = rect(t.a.row, t.a.col, t.b!.row, t.b!.col);
      else if (t.kind === "cols") r = rect(0, t.a.col, MAX_ROWS - 1, t.b!.col);
      else r = rect(t.a.row, 0, t.b!.row, MAX_COLS - 1);
      out.push({ rect: r, color: REF_COLORS[i % REF_COLORS.length], sheetName: t.sheet });
    });
    return out;
  }

  /** Bir sonraki onayda veri doğrulama atlanır (uyarı kutusunda "Evet"). */
  private skipValidation = false;

  /** Veri doğrulama kuralı ihlali. */
  validateInput(sheet: SheetModel, row: number, col: number, text: string): { title: string; message: string; style: "stop" | "warning" | "information" } | null {
    return validateInputRule(this.wb, sheet, row, col, text);
  }

  /* ------------------------------ Biçim Boyacısı ------------------------------ */

  startPainter(sticky: boolean): void {
    this.painter = { sheetId: this.sheet.id, rect: { ...this.range }, sticky };
    this.status = "Biçimi yapıştırmak için hücreleri seçin";
    this.emit();
  }

  applyPainter(): void {
    const p = this.painter;
    if (!p) return;
    const src = this.wb.sheetById(p.sheetId);
    if (!src) return;
    const dst = this.range;
    const rows = p.rect.r2 - p.rect.r1 + 1;
    const cols = p.rect.c2 - p.rect.c1 + 1;
    const target = rect(dst.r1, dst.c1, Math.max(dst.r2, dst.r1 + rows - 1), Math.max(dst.c2, dst.c1 + cols - 1));
    const wb = this.wb;
    const sheet = this.sheet;
    this.ops.cellsOp(sheet, [target], "Biçim Boyacısı", () => {
      for (let i = target.r1; i <= target.r2; i++)
        for (let j = target.c1; j <= target.c2; j++) {
          const sr = p.rect.r1 + ((i - target.r1) % rows);
          const sc = p.rect.c1 + ((j - target.c1) % cols);
          wb.setStyleRaw(sheet, i, j, wb.styleIdAt(src, sr, sc) || undefined);
        }
    });
    if (!p.sticky) {
      this.painter = null;
      this.status = "";
    }
  }

  cancelModes(): void {
    this.painter = null;
    this.clip = null;
    this.drawShape = null;
    this.status = "";
    this.emit();
  }

  /** Etkin hücre adresi. */
  activeAddress(): string {
    return `${colToName(this.sel.active.col)}${this.sel.active.row + 1}`;
  }
}

/** Aralık tek hücre (ya da tek birleştirilmiş alan) mı? */
function rectIsCell(sheet: SheetModel, r: Rect): boolean {
  if (r.r1 === r.r2 && r.c1 === r.c2) return true;
  const m = sheet.mergeAt(r.r1, r.c1);
  return !!m && m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2;
}

export { autoCloseParens, parseRangeLoose, validationList };
