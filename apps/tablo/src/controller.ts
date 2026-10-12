/**
 * Uygulama denetleyicisi: çalışma kitabı modeli, seçim, düzenleme oturumu,
 * pano ve görünüm durumu. React bileşenleri `version` üzerinden abone olur.
 */
import { addr, colToName, findRefSpans, parseLocalFormula, TR_SYNTAX, quoteSheetName, cycleAbsolute, renderRefToken, tokenize, shiftFormula, MAX_COLS, MAX_ROWS } from "@officehere/formula";
import type { FileRef } from "@officehere/core";
import { CondFormatter } from "./model/condfmt";
import { Ops } from "./model/ops";
import { rect, rectToA1 } from "./model/rect";
import { expandForMerges, extendTo, jumpCell, nextInSelection, singleSelection, stepCell } from "./model/selection";
import type { CellPos, Rect, Selection } from "./model/types";
import { WorkbookModel, type SheetModel } from "./model/workbook";

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
  showFormulaBar = true;
  formulaBarExpanded = false;
  /** Durum çubuğu iletisi (ör. "Hedefi seçin ve ENTER tuşuna basın"). */
  status = "";
  /** Seçili yüzen nesne (grafik, resim, şekil). */
  selObj: { kind: "chart" | "image" | "shape"; id: string } | null = null;
  /** Biçim Boyacısı: kopyalanan stiller. */
  painter: { sheetId: number; rect: Rect; sticky: boolean } | null = null;
  /** Odak ızgaraya geri verilsin diye çağrılır. */
  focusGrid: () => void = () => {};
  /** Izgarada bir hücreyi görünür yap. */
  scrollIntoView: (row: number, col: number) => void = () => {};
  /** İletişim kutusu açma isteği (App tarafından ayarlanır). */
  openDialog: (id: string, arg?: unknown) => void = () => {};
  /** Uyarı/hata mesajı göster. */
  alert: (msg: string, title?: string) => void = () => {};

  constructor(wb = new WorkbookModel()) {
    this.wb = wb;
    this.ops = new Ops(wb);
    this.cf = new CondFormatter(wb);
    this.ops.onChange = () => {
      this.dirty = true;
      this.pristine = false;
      this.emit();
    };
  }

  /** Yeni çalışma kitabı yükler. */
  load(wb: WorkbookModel, fileName: string, ref: FileRef | null): void {
    this.wb = wb;
    this.ops = new Ops(wb);
    this.cf = new CondFormatter(wb);
    this.ops.onChange = () => {
      this.dirty = true;
      this.pristine = false;
      this.emit();
    };
    wb.engine.recalc();
    wb.sheets.forEach((s) => this.ops.syncHiddenRows(s));
    this.fileName = fileName;
    this.fileRef = ref;
    this.dirty = false;
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
    this.sheet.selection = sel;
    if (scroll) this.scrollIntoView(this.corner.row, this.corner.col);
    if (this.painter) this.applyPainter();
    this.emit();
  }

  selectCell(row: number, col: number, scroll = true): void {
    const s = singleSelection(this.sheet, row, col);
    this.selObj = null;
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
    if (extend) {
      const from = this.corner;
      const next = jump ? jumpCell(this.wb, sh, from, dr, dc) : stepCell(sh, from, dr, dc);
      this.corner = next;
      this.sheet.selection = extendTo(sh, this.sel, next);
      this.scrollIntoView(next.row, next.col);
      this.emit();
      return;
    }
    const next = jump ? jumpCell(this.wb, sh, this.sel.active, dr, dc) : stepCell(sh, this.sel.active, dr, dc);
    this.selectCell(next.row, next.col);
  }

  /** Enter/Tab sonrası seçim içinde ilerleme. */
  advance(dir: "down" | "up" | "right" | "left"): void {
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

  /** Hücre düzenlemeyi başlatır. text verilirse "Gir" kipinde yazılanla başlar. */
  startEdit(text?: string, inBar = false): void {
    if (this.edit) return;
    const { row, col } = this.sel.active;
    if (this.sheet.protected) {
      this.alert("Değiştirmeye çalıştığınız hücre korumalı bir sayfada.");
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
    let refText = rectToA1(r.r1 === r.r2 && r.c1 === r.c2 ? r : r);
    if (sheet.id !== e.sheetId) refText = `${quoteSheetName(sheet.name)}!${refText}`;
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
    const validation = this.validateInput(sheet, e.row, e.col, text);
    if (validation) {
      this.edit = e;
      this.alert(validation.message, validation.title);
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
      this.ops.enterText(sheet, cells, text);
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

  /** Veri doğrulama kuralı ihlali. */
  validateInput(sheet: SheetModel, row: number, col: number, text: string): { title: string; message: string } | null {
    const v = sheet.validations.find((x) => x.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2));
    if (!v || text === "" || text.startsWith("=")) return null;
    const rule = v.rule;
    if (rule.type === "any" || (rule.errorStyle && rule.errorStyle !== "stop")) return null;
    const fail = { title: rule.errorTitle || "OfficeHere Tablo", message: rule.error || "Bu değer, bu hücre için tanımlanan veri doğrulama kısıtlamalarıyla eşleşmiyor." };
    if (rule.type === "list") {
      const items = validationList(this.wb, sheet, v.rule.formula1 ?? "", row, col);
      return items.some((i) => i.toLocaleUpperCase("tr-TR") === text.trim().toLocaleUpperCase("tr-TR")) ? null : fail;
    }
    const n = Number(text.replace(/\./g, "").replace(",", "."));
    const evalF = (f?: string) => {
      if (!f) return NaN;
      if (f.startsWith("=")) {
        const r = this.wb.engine.evaluate(sheet.id, row, col, f);
        return typeof r === "number" ? r : NaN;
      }
      return Number(f.replace(",", "."));
    };
    let x = n;
    if (rule.type === "textLength") x = text.length;
    if (rule.type === "date") {
      const p = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(text.trim());
      x = p ? Math.round((Date.UTC(+p[3], +p[2] - 1, +p[1]) - Date.UTC(1899, 11, 30)) / 86400000) : NaN;
    }
    if (Number.isNaN(x)) return fail;
    if (rule.type === "whole" && !Number.isInteger(x)) return fail;
    if (rule.type === "custom") {
      const r = this.wb.engine.evaluate(sheet.id, row, col, rule.formula1 ?? "=TRUE");
      return r === true ? null : fail;
    }
    const a = evalF(rule.formula1);
    const b = evalF(rule.formula2);
    let ok = true;
    switch (rule.operator ?? "between") {
      case "between":
        ok = x >= a && x <= b;
        break;
      case "notBetween":
        ok = x < a || x > b;
        break;
      case "equal":
        ok = x === a;
        break;
      case "notEqual":
        ok = x !== a;
        break;
      case "greaterThan":
        ok = x > a;
        break;
      case "lessThan":
        ok = x < a;
        break;
      case "greaterThanOrEqual":
        ok = x >= a;
        break;
      case "lessThanOrEqual":
        ok = x <= a;
        break;
    }
    return ok ? null : fail;
  }

  /* ------------------------------ Biçim Boyacısı ------------------------------ */

  startPainter(sticky: boolean): void {
    this.painter = { sheetId: this.sheet.id, rect: { ...this.range }, sticky };
    this.status = "Biçimi yapıştırmak için hücreleri seçin";
    this.emit();
  }

  private applyPainter(): void {
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
    this.status = "";
    this.emit();
  }

  /** Etkin hücre adresi. */
  activeAddress(): string {
    return `${colToName(this.sel.active.col)}${this.sel.active.row + 1}`;
  }
}

/** Kapanmamış parantezleri tamamlar (Excel davranışı). */
export function autoCloseParens(text: string): string {
  let depth = 0;
  let inStr = false;
  for (const ch of text) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === "(") depth++;
    else if (!inStr && ch === ")") depth--;
  }
  return depth > 0 && !inStr ? text + ")".repeat(depth) : text;
}

/** "A1", "a1:b5", "A:A", "3:5" gibi başvuruları gevşek ayrıştırır. */
export function parseRangeLoose(text: string): Rect | null {
  const t = text.trim().replace(/\$/g, "").toUpperCase();
  const m1 = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(t);
  const col = (s: string) => s.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  if (m1) {
    const r1 = +m1[2] - 1;
    const c1 = col(m1[1]);
    const r2 = m1[4] ? +m1[4] - 1 : r1;
    const c2 = m1[3] ? col(m1[3]) : c1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS || c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return rect(r1, c1, r2, c2);
  }
  const m2 = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(t);
  if (m2) return rect(0, col(m2[1]), MAX_ROWS - 1, col(m2[2]));
  const m3 = /^(\d{1,7}):(\d{1,7})$/.exec(t);
  if (m3) return rect(+m3[1] - 1, 0, +m3[2] - 1, MAX_COLS - 1);
  return null;
}

/** Veri doğrulama listesi öğeleri ("a;b;c" ya da "=$A$1:$A$5"). */
export function validationList(wb: WorkbookModel, sheet: SheetModel, formula: string, row: number, col: number): string[] {
  if (!formula) return [];
  if (formula.startsWith("=")) {
    const m = wb.engine.evaluateArray(sheet.id, row, col, formula);
    return m.flat().filter((v) => v !== null && v !== "").map((v) => (typeof v === "number" ? String(v).replace(".", ",") : String(v)));
  }
  return formula
    .split(/[;,]/)
    .map((s) => s.trim().replace(/^"|"$/g, ""))
    .filter(Boolean);
}
