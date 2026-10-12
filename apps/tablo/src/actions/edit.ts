/** Düzenleme eylemleri: pano, temizleme, doldurma, ekleme/silme, Otomatik Toplam. */
import { addr, formatRange, jsDateToSerial, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { buildClipboard, clipIdFromHtml, getInternalClipboard, pastePayload, pasteText, type PasteOptions } from "../model/clipboard";
import { MAX_COLS, MAX_ROWS, isFullCols, isFullRows, rect, rectCols, rectRows } from "../model/rect";
import { currentRegion, guessSumRange } from "../model/selection";

/* ------------------------------ Pano ------------------------------ */

export function copySelection(ctl: Controller, cut: boolean): { tsv: string; html: string } | null {
  if (ctl.sel.ranges.length > 1) {
    ctl.alert("Bu komut çoklu seçimlerde kullanılamaz.");
    return null;
  }
  const r = ctl.range;
  const { tsv, html, payload } = buildClipboard(ctl.wb, ctl.sheet, r);
  ctl.clip = { sheetId: ctl.sheet.id, rect: payload.src, cut };
  ctl.status = cut ? "Hedefi seçin ve ENTER tuşuna basın ya da Yapıştır'ı seçin" : "Hedefi seçin ve ENTER tuşuna basın ya da Yapıştır'ı seçin";
  ctl.emit();
  return { tsv, html };
}

/** Sistem panosuna yaz (düğmelerden çağrıldığında). */
export async function writeSystemClipboard(data: { tsv: string; html: string }): Promise<void> {
  try {
    const nav = navigator as Navigator & { clipboard?: Clipboard };
    if (nav.clipboard && "write" in nav.clipboard && typeof ClipboardItem !== "undefined") {
      await nav.clipboard.write([new ClipboardItem({ "text/plain": new Blob([data.tsv], { type: "text/plain" }), "text/html": new Blob([data.html], { type: "text/html" }) })]);
      return;
    }
    await nav.clipboard?.writeText(data.tsv);
  } catch {
    // İzin yoksa iç pano kullanılır
  }
}

/** Pano olayından ya da düğmeden gelen veriyi yapıştırır. */
export function pasteData(ctl: Controller, text: string, html: string | undefined, opts?: PasteOptions): void {
  if (ctl.edit || !ctl.guardSelection()) return;
  const internal = getInternalClipboard();
  const id = html ? clipIdFromHtml(html) : null;
  const useInternal = internal && ((id && id === internal.id) || (!text && !html) || (!id && !html && text.replace(/\r\n$/, "") === internalTsv(internal)));
  const target = ctl.range;
  let dst = null;
  if (useInternal) {
    const cut = !!ctl.clip?.cut;
    dst = pastePayload(ctl.ops, ctl.sheet, target, internal, opts ?? { what: "all" }, cut);
    if (cut) ctl.clip = null;
  } else if (text || html) {
    dst = pasteText(ctl.ops, ctl.sheet, target, text, html);
  }
  if (dst) {
    ctl.selectRange(dst, ctl.sel.active.row >= dst.r1 && ctl.sel.active.row <= dst.r2 ? ctl.sel.active : undefined);
  }
  if (!ctl.clip) ctl.status = "";
  ctl.emit();
}

function internalTsv(p: NonNullable<ReturnType<typeof getInternalClipboard>>): string {
  return p.cells.map((r) => r.map((c) => c?.text ?? "").join("\t")).join("\r\n");
}

/** Şerit düğmesinden yapıştırma: sistem panosunu okumayı dener. */
export async function pasteFromButton(ctl: Controller, opts?: PasteOptions): Promise<void> {
  let text = "";
  let html: string | undefined;
  try {
    const cb = (navigator as Navigator & { clipboard?: Clipboard }).clipboard;
    if (cb && typeof cb.read === "function") {
      const items = await cb.read();
      for (const it of items) {
        if (it.types.includes("text/html")) html = await (await it.getType("text/html")).text();
        if (it.types.includes("text/plain")) text = await (await it.getType("text/plain")).text();
      }
    } else if (cb) text = await cb.readText();
  } catch {
    text = "";
    html = undefined;
  }
  const internal = getInternalClipboard();
  if (opts && internal) {
    pasteData(ctl, "", undefined, opts);
    return;
  }
  pasteData(ctl, text, html, opts);
}

/* ------------------------------ Temizleme ------------------------------ */

export function clearContents(ctl: Controller): void {
  if (!ctl.guardSelection()) return;
  ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "contents");
}

/* ------------------------------ Doldurma ------------------------------ */

export function fillDirection(ctl: Controller, dir: "down" | "right" | "up" | "left"): void {
  const r = ctl.range;
  const sheet = ctl.sheet;
  if (dir === "down") {
    const src = rectRows(r) === 1 ? (r.r1 > 0 ? rect(r.r1 - 1, r.c1, r.r1 - 1, r.c2) : null) : rect(r.r1, r.c1, r.r1, r.c2);
    if (!src) return;
    ctl.ops.fill(sheet, src, rect(src.r1, r.c1, r.r2, r.c2), "copy");
  } else if (dir === "right") {
    const src = rectCols(r) === 1 ? (r.c1 > 0 ? rect(r.r1, r.c1 - 1, r.r2, r.c1 - 1) : null) : rect(r.r1, r.c1, r.r2, r.c1);
    if (!src) return;
    ctl.ops.fill(sheet, src, rect(r.r1, src.c1, r.r2, r.c2), "copy");
  } else if (dir === "up") {
    if (rectRows(r) < 2) return;
    ctl.ops.fill(sheet, rect(r.r2, r.c1, r.r2, r.c2), r, "copy");
  } else {
    if (rectCols(r) < 2) return;
    ctl.ops.fill(sheet, rect(r.r1, r.c2, r.r2, r.c2), r, "copy");
  }
}

/** Doldurma tutamacına çift tıklama: komşu sütunun verisi boyunca aşağı doldurur. */
export function autoFillDown(ctl: Controller): void {
  const r = ctl.range;
  const sheet = ctl.sheet;
  const wb = ctl.wb;
  const probe = (c: number) => {
    if (c < 0 || c >= MAX_COLS) return r.r2;
    let row = r.r2 + 1;
    if (wb.isEmpty(sheet, row, c)) return r.r2;
    while (row < MAX_ROWS && !wb.isEmpty(sheet, row, c)) row++;
    return row - 1;
  };
  const end = Math.max(probe(r.c1 - 1), probe(r.c2 + 1));
  if (end <= r.r2) return;
  const dst = rect(r.r1, r.c1, end, r.c2);
  ctl.ops.fill(sheet, r, dst, "series");
  ctl.selectRange(dst, ctl.sel.active);
}

/* ------------------------------ Otomatik Toplam ------------------------------ */

export function autoSum(ctl: Controller, fn = "SUM"): void {
  const sheet = ctl.sheet;
  const r = ctl.range;
  const wb = ctl.wb;
  const trName: Record<string, string> = { SUM: "TOPLA", AVERAGE: "ORTALAMA", COUNT: "BAĞ_DEĞ_SAY", MAX: "MAK", MIN: "MİN" };
  if (rectRows(r) > 1 || rectCols(r) > 1) {
    // Seçimin altına (ve gerekirse sağına) toplam yaz
    const entries: { row: number; col: number; input: { formula: string } }[] = [];
    const lastRowEmpty = Array.from({ length: rectCols(r) }, (_, j) => wb.isEmpty(sheet, r.r2, r.c1 + j)).every(Boolean);
    const dataR2 = lastRowEmpty && rectRows(r) > 1 ? r.r2 - 1 : r.r2;
    const outRow = lastRowEmpty && rectRows(r) > 1 ? r.r2 : r.r2 + 1;
    if (rectRows(r) > 1)
      for (let c = r.c1; c <= r.c2; c++) {
        const has = Array.from({ length: dataR2 - r.r1 + 1 }, (_, i) => typeof wb.getValue(sheet, r.r1 + i, c) === "number").some(Boolean);
        if (has) entries.push({ row: outRow, col: c, input: { formula: `=${fn}(${formatRange(rect(r.r1, c, dataR2, c))})` } });
      }
    else
      entries.push({ row: r.r1, col: r.c2 + 1, input: { formula: `=${fn}(${formatRange(r)})` } });
    ctl.ops.setInputs(sheet, entries, "Otomatik Toplam");
    return;
  }
  const pos = ctl.sel.active;
  const guess = guessSumRange(wb, sheet, pos);
  const text = `=${trName[fn] ?? fn}(${guess ? formatRange(guess) : ""})`;
  ctl.startEdit(text);
  if (ctl.edit && guess) {
    const start = text.indexOf("(") + 1;
    ctl.edit.pointSpan = { start, end: text.length - 1 };
    ctl.edit.caret = text.length - 1;
    ctl.edit.pointAnchor = { row: guess.r1, col: guess.c1 };
    ctl.edit.pointCorner = { row: guess.r2, col: guess.c2 };
    ctl.edit.pointSheetId = sheet.id;
  } else if (ctl.edit) ctl.edit.caret = text.length - 1;
  ctl.emit();
}

/* ------------------------------ Tarih / saat ------------------------------ */

export function insertNow(ctl: Controller, kind: "date" | "time"): void {
  const now = jsDateToSerial(new Date());
  const v: Scalar = kind === "date" ? Math.floor(now) : now - Math.floor(now);
  const { row, col } = ctl.sel.active;
  const sheet = ctl.sheet;
  ctl.ops.cellsOp(sheet, [rect(row, col)], "Tarih Ekle", () => {
    ctl.wb.setInputRaw(sheet, row, col, v);
    const id = ctl.wb.styleIdAt(sheet, row, col);
    ctl.wb.setStyleRaw(sheet, row, col, ctl.wb.styles.patch(id, { numFmt: kind === "date" ? "dd.mm.yyyy" : "hh:mm" }));
  });
}

/* ------------------------------ Ekle / Sil ------------------------------ */

export function insertRows(ctl: Controller): void {
  const r = ctl.range;
  ctl.ops.insertRows(ctl.sheet, r.r1, isFullRows(r) ? rectRows(r) : rectRows(r));
}

export function insertCols(ctl: Controller): void {
  const r = ctl.range;
  ctl.ops.insertCols(ctl.sheet, r.c1, rectCols(r));
}

export function deleteRows(ctl: Controller): void {
  const r = ctl.range;
  ctl.ops.deleteRows(ctl.sheet, r.r1, rectRows(r));
  ctl.selectCell(Math.min(r.r1, MAX_ROWS - 1), ctl.sel.active.col);
}

export function deleteCols(ctl: Controller): void {
  const r = ctl.range;
  ctl.ops.deleteCols(ctl.sheet, r.c1, rectCols(r));
  ctl.selectCell(ctl.sel.active.row, Math.min(r.c1, MAX_COLS - 1));
}

/** Ctrl++ / Hücre Ekle: tüm satır/sütun seçiliyse doğrudan, değilse iletişim kutusu. */
export function insertCellsSmart(ctl: Controller): void {
  const r = ctl.range;
  if (isFullRows(r)) insertRows(ctl);
  else if (isFullCols(r)) insertCols(ctl);
  else ctl.openDialog("insertCells");
}

export function deleteCellsSmart(ctl: Controller): void {
  const r = ctl.range;
  if (isFullRows(r)) deleteRows(ctl);
  else if (isFullCols(r)) deleteCols(ctl);
  else ctl.openDialog("deleteCells");
}

/** Ctrl+A: önce geçerli bölge, ikinci basışta tüm sayfa. */
export function selectAllSmart(ctl: Controller): void {
  const region = currentRegion(ctl.wb, ctl.sheet, ctl.sel.active);
  const cur = ctl.range;
  const single = region.r1 === region.r2 && region.c1 === region.c2;
  if (!single && !(cur.r1 === region.r1 && cur.r2 === region.r2 && cur.c1 === region.c1 && cur.c2 === region.c2)) ctl.selectRange(region, ctl.sel.active);
  else ctl.selectAll();
}

export function cellLabel(row: number, col: number): string {
  return addr(row, col);
}

export function undo(ctl: Controller): void {
  if (ctl.edit) {
    ctl.cancelEdit();
    return;
  }
  ctl.ops.undo();
}

export function redo(ctl: Controller): void {
  if (ctl.edit) return;
  ctl.ops.redo();
}
