/**
 * Formül Denetimi eylemleri: Etkileyenleri/Etkilenenleri İzle (izleme
 * okları, düzey düzey), Okları Kaldır, Hatayı İzle, arka plan hata
 * denetimi (yeşil üçgen kuralları), Hatayı Yoksay, Sayıya Dönüştür,
 * İzleme Penceresi.
 */
import { addr, cellKey, dependentsOf, isError, parseCellInput, precedentsOf, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { rect } from "../model/rect";
import type { CellPos, Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

export interface TraceArrow {
  /** Okun çizildiği sayfa. */
  sheetId: number;
  kind: "prec" | "dep";
  /** Kaynak (etkileyen) hücre/aralık; dış sayfadaysa null. */
  from: Rect | null;
  /** Hedef (formül) hücresi; dış sayfadaysa null. */
  to: CellPos | null;
  /** Kaynak değeri hata mı (kırmızı ok). */
  error?: boolean;
}

const MAX_ARROWS = 400;

function hasFormula(wb: WorkbookModel, sheet: SheetModel, row: number, col: number): boolean {
  return !!wb.getFormula(sheet, row, col);
}

function rangeHasError(wb: WorkbookModel, sheet: SheetModel, r: Rect): boolean {
  for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 200); i++)
    for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 50); j++) if (isError(wb.getValue(sheet, i, j))) return true;
  return false;
}

const key = (p: CellPos) => cellKey(p.row, p.col);

/** Etkileyenleri İzle: her tıklama bir düzey daha geriye gider. */
export function tracePrecedents(ctl: Controller): void {
  const { wb, sheet } = ctl;
  const start = { ...ctl.sel.active };
  if (!hasFormula(wb, sheet, start.row, start.col)) {
    ctl.alert("Etkin hücre, geçerli bir başvuru içeren bir formül ya da ad içermelidir.");
    return;
  }
  const traced = new Set(ctl.arrows.filter((a) => a.kind === "prec" && a.sheetId === sheet.id && a.to).map((a) => key(a.to!)));
  let level: CellPos[] = [start];
  for (let depth = 0; depth < 60; depth++) {
    const fresh = level.filter((c) => !traced.has(key(c)) && hasFormula(wb, sheet, c.row, c.col));
    if (fresh.length) {
      for (const c of fresh) addPrecedentArrows(ctl, sheet, c);
      ctl.emit();
      return;
    }
    const next = new Map<number, CellPos>();
    for (const c of level)
      for (const p of precedentsOf(wb.engine, sheet.id, c.row, c.col)) {
        if (p.sheet !== sheet.id) continue;
        for (let i = p.r1; i <= Math.min(p.r2, p.r1 + 300); i++)
          for (let j = p.c1; j <= Math.min(p.c2, p.c1 + 60); j++) if (hasFormula(wb, sheet, i, j)) next.set(cellKey(i, j), { row: i, col: j });
      }
    if (!next.size) return;
    level = [...next.values()];
  }
}

function addPrecedentArrows(ctl: Controller, sheet: SheetModel, c: CellPos): void {
  for (const p of precedentsOf(ctl.wb.engine, sheet.id, c.row, c.col)) {
    if (ctl.arrows.length >= MAX_ARROWS) return;
    if (p.sheet !== sheet.id) {
      ctl.arrows.push({ sheetId: sheet.id, kind: "prec", from: null, to: c });
      continue;
    }
    const r = rect(p.r1, p.c1, p.r2, p.c2);
    ctl.arrows.push({ sheetId: sheet.id, kind: "prec", from: r, to: c, error: rangeHasError(ctl.wb, sheet, r) });
  }
}

/** Etkilenenleri İzle: her tıklama bir düzey daha ileri gider. */
export function traceDependents(ctl: Controller): void {
  const { wb, sheet } = ctl;
  const start = { ...ctl.sel.active };
  const traced = new Set(ctl.arrows.filter((a) => a.kind === "dep" && a.sheetId === sheet.id && a.from).map((a) => cellKey(a.from!.r1, a.from!.c1)));
  let level: CellPos[] = [start];
  for (let depth = 0; depth < 60; depth++) {
    const fresh = level.filter((c) => !traced.has(key(c)));
    let added = false;
    for (const c of fresh) {
      const deps = dependentsOf(wb.engine, sheet.id, c.row, c.col);
      for (const d of deps) {
        if (ctl.arrows.length >= MAX_ARROWS) break;
        added = true;
        if (d.sheet !== sheet.id) ctl.arrows.push({ sheetId: sheet.id, kind: "dep", from: rect(c.row, c.col), to: null });
        else ctl.arrows.push({ sheetId: sheet.id, kind: "dep", from: rect(c.row, c.col), to: { row: d.row, col: d.col }, error: isError(wb.getValue(sheet, c.row, c.col)) });
      }
    }
    if (added) {
      ctl.emit();
      return;
    }
    if (depth === 0 && !fresh.length && !level.length) break;
    const next = new Map<number, CellPos>();
    for (const c of level) for (const d of dependentsOf(wb.engine, sheet.id, c.row, c.col)) if (d.sheet === sheet.id) next.set(cellKey(d.row, d.col), { row: d.row, col: d.col });
    if (!next.size) break;
    level = [...next.values()];
  }
  if (!ctl.arrows.some((a) => a.kind === "dep" && a.sheetId === sheet.id)) ctl.alert("Etkilenenleri İzle komutu, etkin hücreye başvuran formül bulamadı.");
}

export function removeArrows(ctl: Controller, kind?: "prec" | "dep"): void {
  ctl.arrows = kind ? ctl.arrows.filter((a) => a.kind !== kind) : [];
  ctl.emit();
}

/** Hatayı İzle: hata değerli formülün hata kaynağına kadar kırmızı/mavi oklar. */
export function traceError(ctl: Controller): void {
  const { wb, sheet } = ctl;
  const start = ctl.sel.active;
  if (!isError(wb.getValue(sheet, start.row, start.col)) || !hasFormula(wb, sheet, start.row, start.col)) {
    ctl.alert("Etkin hücre bir hata değeri içeren formül içermelidir.");
    return;
  }
  const seen = new Set<number>();
  const walk = (c: CellPos, depth: number) => {
    if (depth > 30 || seen.has(key(c)) || ctl.arrows.length >= MAX_ARROWS) return;
    seen.add(key(c));
    for (const p of precedentsOf(wb.engine, sheet.id, c.row, c.col)) {
      if (p.sheet !== sheet.id) continue;
      const r = rect(p.r1, p.c1, p.r2, p.c2);
      const err = rangeHasError(wb, sheet, r);
      ctl.arrows.push({ sheetId: sheet.id, kind: "prec", from: r, to: c, error: err });
      if (!err) continue;
      for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 200); i++)
        for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 50); j++) if (isError(wb.getValue(sheet, i, j)) && hasFormula(wb, sheet, i, j)) walk({ row: i, col: j }, depth + 1);
    }
  };
  walk(start, 0);
  ctl.emit();
}

/* ------------------------- Arka plan hata denetimi ------------------------- */

export type ErrorKind = "formulaError" | "numberAsText" | "inconsistent";

export const ERROR_TITLES: Record<string, string> = {
  "#DIV/0!": "Sıfıra Bölme Hatası",
  "#NAME?": "Geçersiz Ad Hatası",
  "#N/A": "Kullanılamayan Değer Hatası",
  "#REF!": "Geçersiz Hücre Başvurusu Hatası",
  "#VALUE!": "Değer Hatası",
  "#NUM!": "Geçersiz Sayısal Değer Hatası",
  "#NULL!": "Boş Kesişim Hatası",
  "#SPILL!": "Taşma Hatası",
  "#CALC!": "Hesaplama Hatası",
};

export const ERROR_HELP: Record<string, string> = {
  "#DIV/0!": "Formül ya da işlev sıfıra ya da boş hücrelere bölme yapıyor.",
  "#NAME?": "Formül tanınmayan bir metin içeriyor.",
  "#N/A": "Formül ya da işlev için bir değer kullanılamıyor.",
  "#REF!": "Formül geçerli olmayan bir hücreye başvuruyor.",
  "#VALUE!": "Formülde kullanılan bir değer yanlış veri türünde.",
  "#NUM!": "Formülde ya da işlevde geçersiz sayısal değerler kullanılmış.",
  "#NULL!": "Kesişmeyen iki alanın kesişimi belirtilmiş.",
  "#SPILL!": "Taşma aralığı boş değil.",
  "#CALC!": "Hesaplama altyapısı desteklenmeyen bir durumla karşılaştı.",
};

/** Hücrede arka plan hata göstergesi (yeşil üçgen) gerektiren durum. */
export function cellError(wb: WorkbookModel, sheet: SheetModel, row: number, col: number): { kind: ErrorKind; title: string; help: string; code?: string } | null {
  if (sheet.ignoredErrors.has(cellKey(row, col))) return null;
  const v: Scalar = wb.getValue(sheet, row, col);
  const f = wb.getFormula(sheet, row, col);
  if (f && isError(v)) return { kind: "formulaError", title: ERROR_TITLES[v.code] ?? "Formül Hatası", help: ERROR_HELP[v.code] ?? "", code: v.code };
  if (!f && typeof v === "string" && wb.engine.hasInput(sheet.id, row, col)) {
    const t = v.trim();
    if (t && /^[-+]?[\d.,]+$/.test(t) && /\d/.test(t)) {
      const p = parseCellInput(t);
      if (p.kind === "value" && typeof p.value === "number" && !/^0\d/.test(t)) return { kind: "numberAsText", title: "Metin Olarak Depolanmış Sayı", help: "Bu hücredeki sayı metin olarak biçimlendirilmiş ya da başına kesme işareti konmuş." };
    }
  }
  if (f && inconsistentWithNeighbors(wb, sheet, row, col, f)) return { kind: "inconsistent", title: "Tutarsız Formül", help: "Bu hücredeki formül, alanınızdaki diğer formüllerin düzeniyle eşleşmiyor." };
  return null;
}

/** R1C1 benzeri göreli imza: başvurular hücreye göre kaydırılmış metin. */
function relSignature(formula: string, row: number, col: number): string {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(])/g, (_m, ca: string, c: string, ra: string, r: string) => {
    const cc = c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const rr = Number(r) - 1;
    return `${ca ? "C" + cc : "C[" + (cc - col) + "]"}${ra ? "R" + rr : "R[" + (rr - row) + "]"}`;
  });
}

/** Excel kuralı: iki komşu (sol-sağ ya da üst-alt) aynı göreli formülü kullanırken bu hücre farklıysa. */
function inconsistentWithNeighbors(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, f: string): boolean {
  const sig = relSignature(f, row, col);
  const pair = (a: CellPos, b: CellPos) => {
    if (a.row < 0 || a.col < 0 || b.row < 0 || b.col < 0) return false;
    const fa = wb.getFormula(sheet, a.row, a.col);
    const fb = wb.getFormula(sheet, b.row, b.col);
    if (!fa || !fb) return false;
    const sa = relSignature(fa, a.row, a.col);
    return sa === relSignature(fb, b.row, b.col) && sa !== sig;
  };
  return pair({ row: row - 1, col }, { row: row + 1, col }) || pair({ row, col: col - 1 }, { row, col: col + 1 });
}

export function ignoreError(ctl: Controller, row: number, col: number): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Hatayı Yoksay", () => {
    sheet.ignoredErrors.add(cellKey(row, col));
  });
}

/** Metin olarak depolanmış sayıları sayıya dönüştürür (seçimdeki tüm hücreler). */
export function convertToNumber(ctl: Controller): void {
  const { wb, sheet } = ctl;
  const entries: { row: number; col: number; input: number }[] = [];
  for (const r of ctl.sel.ranges)
    for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 100000); i++)
      for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 1000); j++) {
        const e = cellError(wb, sheet, i, j);
        if (e?.kind !== "numberAsText") continue;
        const p = parseCellInput(String(wb.getValue(sheet, i, j)).trim());
        if (typeof p.value === "number") entries.push({ row: i, col: j, input: p.value });
      }
  if (entries.length) ctl.ops.setInputs(sheet, entries, "Sayıya Dönüştür");
}

/** Sayfadaki hata göstergeli hücreler (Hata Denetimi iletişim kutusu için, satır sırası). */
export function sheetErrors(wb: WorkbookModel, sheet: SheetModel): CellPos[] {
  const out: CellPos[] = [];
  wb.engine.forEachCell(sheet.id, (row, col) => {
    if (cellError(wb, sheet, row, col)) out.push({ row, col });
  });
  return out.sort((a, b) => a.row - b.row || a.col - b.col);
}

/* ------------------------------ İzleme Penceresi ------------------------------ */

export interface Watch {
  sheetId: number;
  row: number;
  col: number;
}

export function addWatches(ctl: Controller): void {
  const sheet = ctl.sheet;
  const have = new Set(ctl.watches.map((w) => `${w.sheetId}:${w.row}:${w.col}`));
  for (const r of ctl.sel.ranges)
    for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 50); i++)
      for (let j = r.c1; j <= Math.min(r.c2, r.c1 + 20); j++) {
        const k = `${sheet.id}:${i}:${j}`;
        if (have.has(k) || ctl.watches.length >= 500) continue;
        have.add(k);
        ctl.watches.push({ sheetId: sheet.id, row: i, col: j });
      }
  ctl.showWatch = true;
  ctl.emit();
}

export function watchLabel(w: Watch): string {
  return addr(w.row, w.col);
}
