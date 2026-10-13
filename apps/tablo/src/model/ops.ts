/**
 * Geri alınabilir çalışma kitabı işlemleri. Her işlem değişen hücrelerin
 * önceki/sonraki durumunu (ya da yapısal işlemlerde anlık görüntüyü) kaydeder.
 */
import { cellKey, keyCol, keyRow, parseCellInput, qualifyStructRefs, type Scalar } from "@officehere/formula";
import { OpsSheets, uniqueTableName } from "./opsSheets";
import { MAX_COLS, MAX_ROWS, contains, intersects, isFullCols, isFullRows, rect, rectCols, rectRows } from "./rect";
import type { CellStyle } from "./styles";
import type { Rect } from "./types";
import { generateFill, type FillSource } from "./fill";
import type { CellInput, SheetModel } from "./workbook";
import { copySparkline } from "../sparkline/sparkOps";
import { inheritedFormat } from "./formatInherit";

export type StylePatch = Partial<CellStyle> | ((s: CellStyle) => Partial<CellStyle>);

export class Ops extends OpsSheets {
  /* ------------------------------ İçerik ------------------------------ */

  /** Kullanıcının yazdığı metni hücrelere uygular (Türkçe ayrıştırma, sayı biçimi algılama). */
  enterText(sheet: SheetModel, cells: { row: number; col: number }[], text: string, label = "Yazma"): void {
    const wb = this.wb;
    const rects = cells.map((c) => rect(c.row, c.col));
    this.cellsOp(sheet, rects, label, () => {
      for (const { row, col } of cells) {
        const st = wb.styleAt(sheet, row, col);
        const p = parseCellInput(text, { numFmt: st.numFmt });
        if (p.kind === "empty") wb.setInputRaw(sheet, row, col, null);
        else if (p.kind === "formula") {
          // Tablosuz yapısal başvurular ([@Adet]) bulunduğu tablonun adıyla nitelenir
          const t = p.formula!.includes("[") ? sheet.tableAt(row, col) : undefined;
          wb.setInputRaw(sheet, row, col, { formula: t ? qualifyStructRefs(p.formula!, t.name) : p.formula! });
        }
        else wb.setInputRaw(sheet, row, col, p.value ?? null);
        if (p.numFmt && (!st.numFmt || st.numFmt === "General")) {
          wb.setStyleRaw(sheet, row, col, wb.styles.patch(wb.styleIdAt(sheet, row, col), { numFmt: p.numFmt }));
        }
        // Formül sonucu, Genel biçimli hücrede başvurduğu hücrenin (ya da fonksiyonun) biçimini alır
        if (p.kind === "formula" && (!st.numFmt || st.numFmt === "General")) {
          const fmt = inheritedFormat(wb, sheet, p.formula!);
          if (fmt) wb.setStyleRaw(sheet, row, col, wb.styles.patch(wb.styleIdAt(sheet, row, col), { numFmt: fmt }));
        }
        // Alt+Enter ile girilen satır sonu metni kaydırmayı açar
        if (typeof p.value === "string" && p.value.includes("\n") && !st.wrap) {
          wb.setStyleRaw(sheet, row, col, wb.styles.patch(wb.styleIdAt(sheet, row, col), { wrap: true }));
        }
      }
    });
  }

  /** Hazır girdileri yazar (yapıştırma, şablon, içe aktarma). */
  setInputs(sheet: SheetModel, entries: { row: number; col: number; input: CellInput; style?: number | null }[], label: string): void {
    if (!entries.length) return;
    let r1 = Infinity;
    let c1 = Infinity;
    let r2 = -1;
    let c2 = -1;
    for (const e of entries) {
      r1 = Math.min(r1, e.row);
      c1 = Math.min(c1, e.col);
      r2 = Math.max(r2, e.row);
      c2 = Math.max(c2, e.col);
    }
    const bounds: Rect = { r1, c1, r2, c2 };
    const rects = (r2 - r1 + 1) * (c2 - c1 + 1) <= 200000 ? [bounds] : entries.map((e) => rect(e.row, e.col));
    this.cellsOp(sheet, rects, label, () => {
      for (const e of entries) {
        this.wb.setInputRaw(sheet, e.row, e.col, e.input);
        if (e.style !== undefined) this.wb.setStyleRaw(sheet, e.row, e.col, e.style === null ? undefined : e.style);
      }
    });
  }

  /** Stil yaması uygular; tüm satır/sütun seçimleri satır/sütun stiline yazılır. */
  applyStyle(sheet: SheetModel, rects: Rect[], patch: StylePatch, label = "Biçimlendirme"): void {
    const wb = this.wb;
    const pf = typeof patch === "function" ? patch : () => patch;
    this.cellsOp(sheet, rects, label, () => {
      for (const r of rects) {
        if (isFullCols(r)) {
          for (let c = r.c1; c <= r.c2; c++) {
            const base = sheet.colStyles.get(c);
            sheet.colStyles.set(c, wb.styles.patch(base, pf(wb.styles.get(base))));
          }
          // Satır stili olan satırların kesişen hücreleri açık stil alır
          const nCols = r.c2 - r.c1 + 1;
          if (sheet.rowStyles.size * nCols > 200000 || isFullRows(r)) {
            for (const [row, base] of [...sheet.rowStyles]) sheet.rowStyles.set(row, wb.styles.patch(base, pf(wb.styles.get(base))));
          } else {
            for (const [row, base] of sheet.rowStyles)
              for (let c = r.c1; c <= r.c2; c++) {
                const k = cellKey(row, c);
                if (sheet.styles.has(k)) continue;
                sheet.styles.set(k, wb.styles.patch(base, pf(wb.styles.get(base))));
              }
          }
        } else if (isFullRows(r)) {
          for (let row = r.r1; row <= r.r2; row++) {
            const base = sheet.rowStyles.get(row);
            sheet.rowStyles.set(row, wb.styles.patch(base, pf(wb.styles.get(base))));
          }
        }
        const keys = wb.cellKeysIn(sheet, r);
        if (!isFullCols(r) && !isFullRows(r)) {
          for (let row = r.r1; row <= r.r2; row++)
            for (let col = r.c1; col <= r.c2; col++) {
              const id = wb.styleIdAt(sheet, row, col);
              wb.setStyleRaw(sheet, row, col, wb.styles.patch(id, pf(wb.styles.get(id))));
            }
        } else {
          for (const k of keys) {
            const row = keyRow(k);
            const col = keyCol(k);
            const own = sheet.styles.get(k);
            if (own === undefined) continue;
            wb.setStyleRaw(sheet, row, col, wb.styles.patch(own, pf(wb.styles.get(own))));
          }
        }
      }
    });
  }

  /** Kenarlık uygulama (dış/iç kenarlıklar ayrımıyla). */
  applyBorders(sheet: SheetModel, r: Rect, kind: string, side?: CellStyle["bTop"]): void {
    const wb = this.wb;
    const thin = side ?? { style: "thin" as const, color: "#000000" };
    this.cellsOp(sheet, [rect(Math.max(0, r.r1 - 1), Math.max(0, r.c1 - 1), Math.min(MAX_ROWS - 1, r.r2 + 1), Math.min(MAX_COLS - 1, r.c2 + 1))], "Kenarlıklar", () => {
      const set = (row: number, col: number, p: Partial<CellStyle>) => {
        const id = wb.styleIdAt(sheet, row, col);
        wb.setStyleRaw(sheet, row, col, wb.styles.patch(id, p));
      };
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++)
        for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 500); col++) {
          const top = row === r.r1;
          const bottom = row === r.r2;
          const left = col === r.c1;
          const right = col === r.c2;
          const p: Partial<CellStyle> = {};
          switch (kind) {
            case "bottom":
              if (bottom) p.bBottom = thin;
              break;
            case "top":
              if (top) p.bTop = thin;
              break;
            case "left":
              if (left) p.bLeft = thin;
              break;
            case "right":
              if (right) p.bRight = thin;
              break;
            case "none":
              Object.assign(p, { bTop: undefined, bBottom: undefined, bLeft: undefined, bRight: undefined });
              break;
            case "all":
              Object.assign(p, { bTop: thin, bBottom: thin, bLeft: thin, bRight: thin });
              break;
            case "outside":
              if (top) p.bTop = thin;
              if (bottom) p.bBottom = thin;
              if (left) p.bLeft = thin;
              if (right) p.bRight = thin;
              break;
            case "thickOutside": {
              const t = { style: "medium" as const, color: thin.color };
              if (top) p.bTop = t;
              if (bottom) p.bBottom = t;
              if (left) p.bLeft = t;
              if (right) p.bRight = t;
              break;
            }
            case "doubleBottom":
              if (bottom) p.bBottom = { style: "double", color: thin.color };
              break;
            case "thickBottom":
              if (bottom) p.bBottom = { style: "medium", color: thin.color };
              break;
            case "topBottom":
              if (top) p.bTop = thin;
              if (bottom) p.bBottom = thin;
              break;
            case "topThickBottom":
              if (top) p.bTop = thin;
              if (bottom) p.bBottom = { style: "medium", color: thin.color };
              break;
            case "topDoubleBottom":
              if (top) p.bTop = thin;
              if (bottom) p.bBottom = { style: "double", color: thin.color };
              break;
            case "inside":
              if (!bottom) p.bBottom = thin;
              if (!right) p.bRight = thin;
              break;
          }
          if (Object.keys(p).length) set(row, col, p);
        }
      // Komşu hücrelerin bitişik kenarlıklarını temizle ("Kenarlık Yok")
      if (kind === "none") {
        for (let col = r.c1; col <= r.c2 && col <= r.c1 + 500; col++) {
          if (r.r1 > 0) set(r.r1 - 1, col, { bBottom: undefined });
          if (r.r2 < MAX_ROWS - 1) set(r.r2 + 1, col, { bTop: undefined });
        }
        for (let row = r.r1; row <= r.r2 && row <= r.r1 + 5000; row++) {
          if (r.c1 > 0) set(row, r.c1 - 1, { bRight: undefined });
          if (r.c2 < MAX_COLS - 1) set(row, r.c2 + 1, { bLeft: undefined });
        }
      }
    });
  }

  /** Temizle: tümü / biçimler / içerik / notlar / köprüler. */
  clear(sheet: SheetModel, rects: Rect[], what: "all" | "formats" | "contents" | "notes" | "links"): void {
    const wb = this.wb;
    const labels = { all: "Tümünü Temizle", formats: "Biçimleri Temizle", contents: "İçeriği Temizle", notes: "Notları Temizle", links: "Köprüleri Temizle" };
    const doCells = () => {
      for (const r of rects) {
        for (const k of wb.cellKeysIn(sheet, r)) {
          const row = keyRow(k);
          const col = keyCol(k);
          if (what === "all" || what === "contents") {
            // Taşan dizinin parçasıysa yalnız çapada silinir
            wb.setInputRaw(sheet, row, col, null);
          }
          if (what === "all" || what === "formats") wb.setStyleRaw(sheet, row, col, undefined);
          if (what === "all" || what === "notes") sheet.notes.delete(k);
          if (what === "all" || what === "links") sheet.links.delete(k);
        }
        if ((what === "all" || what === "formats") && (isFullCols(r) || isFullRows(r))) {
          if (isFullCols(r)) for (let c = r.c1; c <= r.c2; c++) sheet.colStyles.delete(c);
          if (isFullRows(r)) for (let row = r.r1; row <= r.r2; row++) sheet.rowStyles.delete(row);
        }
      }
    };
    if (what === "all" || what === "formats") {
      // Birleştirmeler ve koşullu biçimler de temizlenir
      this.sheetOp(
        sheet,
        labels[what],
        () => {
          doCells();
          sheet.merges = sheet.merges.filter((m) => !rects.some((r) => intersects(r, m)));
        },
        rects,
      );
      return;
    }
    this.cellsOp(sheet, rects, labels[what], doCells);
  }

  /* ---------------------------- Yapısal ---------------------------- */

  insertRows(sheet: SheetModel, at: number, count: number): void {
    this.snapshotOp("Satır Ekle", () => {
      this.wb.engine.insertRows(sheet.id, at, count);
      sheet.shiftStructure("row", at, count);
      this.wb.shiftDataRefs(sheet, "row", at, count);
      // Yeni satırlar üstteki satırın biçimini alır
      if (at > 0) {
        const h = sheet.rows.isCustom(at - 1) ? sheet.rows.rawSize(at - 1) : undefined;
        const rs = sheet.rowStyles.get(at - 1);
        for (let r = at; r < at + count; r++) {
          if (h !== undefined) sheet.rows.set(r, h);
          if (rs !== undefined) sheet.rowStyles.set(r, rs);
        }
        for (const [k, id] of [...sheet.styles]) {
          if (keyRow(k) !== at - 1) continue;
          for (let r = at; r < at + count; r++) sheet.styles.set(cellKey(r, keyCol(k)), id);
        }
      }
    });
  }

  deleteRows(sheet: SheetModel, at: number, count: number): void {
    this.snapshotOp("Satır Sil", () => {
      this.wb.engine.deleteRows(sheet.id, at, count);
      sheet.shiftStructure("row", at, -count);
      this.wb.shiftDataRefs(sheet, "row", at, -count);
    });
  }

  insertCols(sheet: SheetModel, at: number, count: number): void {
    this.snapshotOp("Sütun Ekle", () => {
      this.wb.engine.insertCols(sheet.id, at, count);
      sheet.shiftStructure("col", at, count);
      this.wb.shiftDataRefs(sheet, "col", at, count);
      if (at > 0) {
        const w = sheet.cols.isCustom(at - 1) ? sheet.cols.rawSize(at - 1) : undefined;
        const cs = sheet.colStyles.get(at - 1);
        for (let c = at; c < at + count; c++) {
          if (w !== undefined) sheet.cols.set(c, w);
          if (cs !== undefined) sheet.colStyles.set(c, cs);
        }
      }
    });
  }

  deleteCols(sheet: SheetModel, at: number, count: number): void {
    this.snapshotOp("Sütun Sil", () => {
      this.wb.engine.deleteCols(sheet.id, at, count);
      sheet.shiftStructure("col", at, -count);
      this.wb.shiftDataRefs(sheet, "col", at, -count);
    });
  }

  /** Hücre ekle/sil ve kaydır. */
  shiftCells(sheet: SheetModel, r: Rect, dir: "down" | "right" | "up" | "left"): void {
    const label = dir === "down" || dir === "right" ? "Hücre Ekle" : "Hücre Sil";
    this.snapshotOp(label, () => this.shiftCellsRaw(sheet, r, dir));
  }

  /** Hücre kaydırma (geri alma kaydı olmadan; çağıran snapshotOp içinde kullanır). */
  shiftCellsRaw(sheet: SheetModel, r: Rect, dir: "down" | "right" | "up" | "left"): void {
    {
      const wb = this.wb;
      wb.engine.shiftCells(sheet.id, r.r1, r.c1, r.r2, r.c2, dir);
      const h = rectRows(r);
      const w = rectCols(r);
      const moveMap = <T>(m: Map<number, T>) => {
        const out = new Map<number, T>();
        for (const [k, v] of m) {
          let row = keyRow(k);
          let col = keyCol(k);
          const inBandV = col >= r.c1 && col <= r.c2;
          const inBandH = row >= r.r1 && row <= r.r2;
          if (dir === "down" && inBandV && row >= r.r1) row += h;
          else if (dir === "right" && inBandH && col >= r.c1) col += w;
          else if (dir === "up" && inBandV && row >= r.r1) {
            if (row <= r.r2) continue;
            row -= h;
          } else if (dir === "left" && inBandH && col >= r.c1) {
            if (col <= r.c2) continue;
            col -= w;
          }
          if (row < MAX_ROWS && col < MAX_COLS) out.set(cellKey(row, col), v);
        }
        return out;
      };
      sheet.styles = moveMap(sheet.styles);
      sheet.notes = moveMap(sheet.notes);
      sheet.links = moveMap(sheet.links);
    }
  }

  /* ------------------------- Birleştirme ------------------------- */

  merge(sheet: SheetModel, rects: Rect[], mode: "center" | "across" | "merge"): void {
    const wb = this.wb;
    this.sheetOp(
      sheet,
      mode === "center" ? "Birleştir ve Ortala" : "Hücreleri Birleştir",
      () => {
        for (const r of rects) {
          const parts = mode === "across" ? Array.from({ length: rectRows(r) }, (_, i) => rect(r.r1 + i, r.c1, r.r1 + i, r.c2)) : [r];
          for (const p of parts) {
            if (rectRows(p) === 1 && rectCols(p) === 1) continue;
            sheet.merges = sheet.merges.filter((m) => !intersects(m, p));
            // Sol üst dışındaki değerler silinir (Excel uyarısından sonra)
            for (const k of wb.cellKeysIn(sheet, p)) {
              const row = keyRow(k);
              const col = keyCol(k);
              if (row === p.r1 && col === p.c1) continue;
              wb.setInputRaw(sheet, row, col, null);
            }
            sheet.merges.push(p);
            if (mode === "center") {
              const id = wb.styleIdAt(sheet, p.r1, p.c1);
              wb.setStyleRaw(sheet, p.r1, p.c1, wb.styles.patch(id, { hAlign: "center" }));
            }
          }
        }
      },
      rects,
    );
  }

  unmerge(sheet: SheetModel, rects: Rect[]): void {
    this.sheetOp(sheet, "Hücrelerin Birleştirmesini Kaldır", () => {
      sheet.merges = sheet.merges.filter((m) => !rects.some((r) => intersects(r, m)));
    });
  }

  /** Birleştirmeyle çakışan aralıklarda sol üst dışında veri var mı? */
  mergeLosesData(sheet: SheetModel, rects: Rect[]): boolean {
    for (const r of rects) {
      let n = 0;
      for (const k of this.wb.cellKeysIn(sheet, r)) if (this.wb.engine.hasInput(sheet.id, keyRow(k), keyCol(k))) n++;
      if (n > 1) return true;
    }
    return false;
  }

  /* ------------------------- Boyutlar ------------------------- */

  setSizes(sheet: SheetModel, axis: "row" | "col", items: number[], px: number | null, label: string): void {
    this.sheetOp(sheet, label, () => {
      const ax = axis === "row" ? sheet.rows : sheet.cols;
      for (const i of items) {
        if (px === null) ax.reset(i);
        else if (px <= 0) ax.setHidden(i, true);
        else {
          ax.set(i, px);
          ax.setHidden(i, false);
        }
      }
    });
  }

  setHidden(sheet: SheetModel, axis: "row" | "col", items: number[], hidden: boolean): void {
    this.sheetOp(sheet, hidden ? "Gizle" : "Göster", () => {
      const ax = axis === "row" ? sheet.rows : sheet.cols;
      ax.setHiddenMany(items, hidden);
      for (const i of items) {
        if (!hidden && ax.rawSize(i) === 0) ax.reset(i);
        if (axis === "row") sheet.filterHidden.delete(i);
      }
      if (axis === "row") this.syncHiddenRows(sheet);
    });
  }

  /* --------------------------- Doldurma --------------------------- */

  /** Doldurmanın bir satırı/sütunu: kaynak konumları, hedef konumu ve üretilen değerler. */
  private fillLine(sheet: SheetModel, src: Rect, dst: Rect, mode: "series" | "copy" | "formats" | "values", l: number) {
    const wb = this.wb;
    const vertical = dst.c1 === src.c1 && dst.c2 === src.c2;
    const reverse = vertical ? dst.r1 < src.r1 : dst.c1 < src.c1;
    const srcLen = vertical ? rectRows(src) : rectCols(src);
    const count = (vertical ? rectRows(dst) : rectCols(dst)) - srcLen;
    const sources: FillSource[] = [];
    const srcPos: [number, number][] = [];
    for (let i = 0; i < srcLen; i++) {
      const row = vertical ? src.r1 + i : src.r1 + l;
      const col = vertical ? src.c1 + l : src.c1 + i;
      srcPos.push([row, col]);
      sources.push({ input: wb.getInput(sheet, row, col), numFmt: wb.styleAt(sheet, row, col).numFmt });
    }
    const targetPos = (i: number): [number, number] => {
      if (vertical) return [reverse ? src.r1 - 1 - i : src.r2 + 1 + i, src.c1 + l];
      return [src.r1 + l, reverse ? src.c1 - 1 - i : src.c2 + 1 + i];
    };
    const sourceIndex = (i: number) => (reverse ? srcLen - 1 - (i % srcLen) : i % srcLen);
    const values =
      mode === "formats" || count <= 0
        ? []
        : generateFill(sources, count, reverse, mode === "copy", (i, si) => {
            const [tr, tc] = targetPos(i);
            const [sr, sc] = srcPos[si];
            return [tr - sr, tc - sc];
          });
    return { count, values, srcPos, targetPos, sourceIndex };
  }

  /** Kaynak aralıktan hedef aralığa doldurur (hedef kaynağı içerir). */
  fill(sheet: SheetModel, src: Rect, dst: Rect, mode: "series" | "copy" | "formats" | "values" = "series"): void {
    const wb = this.wb;
    const vertical = dst.c1 === src.c1 && dst.c2 === src.c2;
    this.cellsOp(sheet, [dst], "Otomatik Doldurma", () => {
      const lines = vertical ? rectCols(src) : rectRows(src);
      for (let l = 0; l < lines; l++) {
        const { count, values, srcPos, targetPos, sourceIndex } = this.fillLine(sheet, src, dst, mode, l);
        for (let i = 0; i < count; i++) {
          const [row, col] = targetPos(i);
          const [sr, sc] = srcPos[sourceIndex(i)];
          if (mode !== "formats") {
            let v = values[i];
            if (mode === "values" && v !== null && typeof v === "object" && "formula" in v) v = wb.getValue(sheet, sr, sc);
            wb.setInputRaw(sheet, row, col, v ?? null);
          }
          if (mode !== "values") wb.setStyleRaw(sheet, row, col, sheet.styles.get(cellKey(sr, sc)));
          // Mini grafik aynı gruba göreli veri aralığıyla kopyalanır
          if ((mode === "series" || mode === "copy") && sheet.sparklines.length) copySparkline(sheet, { row: sr, col: sc }, { row, col }, wb);
        }
      }
    });
  }

  /**
   * Doldurma tutamacı ipucu: bırakınca hedefin en uç hücresine yazılacak girdi ve kaynağının
   * sayı biçimi. fill() ile aynı üretimi kullanır; hedef kaynaktan büyük değilse null.
   */
  fillPreview(sheet: SheetModel, src: Rect, dst: Rect, mode: "series" | "copy" = "series"): { input: CellInput; numFmt?: string } | null {
    const vertical = dst.c1 === src.c1 && dst.c2 === src.c2;
    const lines = vertical ? rectCols(src) : rectRows(src);
    const { count, values, srcPos, sourceIndex } = this.fillLine(sheet, src, dst, mode, lines - 1);
    if (count <= 0) return null;
    const [sr, sc] = srcPos[sourceIndex(count - 1)];
    return { input: values[count - 1] ?? null, numFmt: this.wb.styleAt(sheet, sr, sc).numFmt };
  }

  /* ----------------------- Not ve köprü ----------------------- */

  setNote(sheet: SheetModel, row: number, col: number, text: string | null, author?: string): void {
    this.cellsOp(sheet, [rect(row, col)], text === null ? "Notu Sil" : "Not", () => {
      const k = cellKey(row, col);
      if (text === null) sheet.notes.delete(k);
      else sheet.notes.set(k, { text, author });
    });
  }

  setLink(sheet: SheetModel, row: number, col: number, url: string | null, display?: string): void {
    const wb = this.wb;
    this.cellsOp(sheet, [rect(row, col)], url === null ? "Köprüyü Kaldır" : "Köprü Ekle", () => {
      const k = cellKey(row, col);
      if (url === null) {
        sheet.links.delete(k);
        return;
      }
      sheet.links.set(k, url);
      if (display !== undefined && display !== "") wb.setInputRaw(sheet, row, col, display);
      else if (!wb.engine.hasInput(sheet.id, row, col)) wb.setInputRaw(sheet, row, col, url);
      const id = wb.styleIdAt(sheet, row, col);
      wb.setStyleRaw(sheet, row, col, wb.styles.patch(id, { color: "#0563C1", underline: "single" }));
    });
  }

  /* ---------------------------- Diğer ---------------------------- */

  /** Bul ve Değiştir: eşleşen hücrelerin metnini değiştirir. */
  replaceAll(sheet: SheetModel, cells: { row: number; col: number }[], find: RegExp, replacement: string): number {
    const wb = this.wb;
    let n = 0;
    this.cellsOp(sheet, cells.map((c) => rect(c.row, c.col)), "Değiştir", () => {
      for (const { row, col } of cells) {
        const text = wb.editText(sheet, row, col);
        const next = text.replace(find, replacement);
        if (next === text) continue;
        n++;
        const p = parseCellInput(next, { numFmt: wb.styleAt(sheet, row, col).numFmt });
        if (p.kind === "formula") wb.setInputRaw(sheet, row, col, { formula: p.formula! });
        else wb.setInputRaw(sheet, row, col, p.kind === "empty" ? null : (p.value as Scalar));
      }
    });
    return n;
  }

  isInside(r: Rect, row: number, col: number): boolean {
    return contains(r, row, col);
  }
}

export { uniqueTableName };
