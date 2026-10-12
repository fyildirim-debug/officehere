/**
 * Izgara geometrisi: başlıklar, dondurulmuş bölmeler, kaydırma ve
 * yakınlaştırmaya göre hücre ↔ ekran dönüşümleri. Tüm değerler CSS pikseli.
 */
import { MAX_COLS, MAX_ROWS } from "../model/rect";
import type { SheetModel } from "../model/workbook";

export interface Quadrant {
  /** Ekrandaki alan. */
  x: number;
  y: number;
  w: number;
  h: number;
  r1: number;
  r2: number;
  c1: number;
  c2: number;
  frozenRows: boolean;
  frozenCols: boolean;
}

export class GridLayout {
  z: number;
  headerH: number;
  headerW: number;
  frozenH: number;
  frozenW: number;
  fr: number;
  fc: number;

  constructor(
    public sheet: SheetModel,
    public width: number,
    public height: number,
    public scrollX: number,
    public scrollY: number,
    measureDigits: (n: number) => number,
  ) {
    this.z = sheet.zoom / 100;
    const z = this.z;
    const show = sheet.showHeadings;
    this.headerH = show ? Math.round(20 * Math.max(0.6, z)) : 0;
    // Satır başlığı genişliği görünen en büyük satır numarasına göre
    this.fr = Math.min(sheet.frozen.rows, MAX_ROWS - 1);
    this.fc = Math.min(sheet.frozen.cols, MAX_COLS - 1);
    this.frozenH = sheet.rows.offset(this.fr) * z;
    this.frozenW = sheet.cols.offset(this.fc) * z;
    this.headerW = 0;
    if (show) {
      const lastRow = this.rowAtY(height + 50);
      const digits = String(Math.max(lastRow + 1, 100)).length;
      this.headerW = Math.round(Math.max(26 * Math.max(0.6, z), measureDigits(digits) + 10 * Math.max(0.6, z)));
    }
  }

  get cellsX(): number {
    return this.headerW + this.frozenW;
  }

  get cellsY(): number {
    return this.headerH + this.frozenH;
  }

  /** Kaydırılabilir alanın genişliği/yüksekliği. */
  get paneW(): number {
    return Math.max(0, this.width - this.cellsX);
  }

  get paneH(): number {
    return Math.max(0, this.height - this.cellsY);
  }

  colX(c: number): number {
    const s = this.sheet.cols;
    if (c < this.fc) return this.headerW + s.offset(c) * this.z;
    return this.cellsX + (s.offset(c) - s.offset(this.fc)) * this.z - this.scrollX;
  }

  rowY(r: number): number {
    const s = this.sheet.rows;
    if (r < this.fr) return this.headerH + s.offset(r) * this.z;
    return this.cellsY + (s.offset(r) - s.offset(this.fr)) * this.z - this.scrollY;
  }

  colW(c: number): number {
    return this.sheet.cols.size(c) * this.z;
  }

  rowH(r: number): number {
    return this.sheet.rows.size(r) * this.z;
  }

  /** Ekran x → sütun (scrollX verilirse onunla). */
  colAtX(x: number, scrollX = this.scrollX): number {
    const s = this.sheet.cols;
    if (x < this.cellsX && this.fc > 0 && x >= this.headerW) return Math.min(this.fc - 1, s.indexAt((x - this.headerW) / this.z));
    return s.indexAt(s.offset(this.fc) + (Math.max(x, this.cellsX) - this.cellsX + scrollX) / this.z);
  }

  rowAtY(y: number, scrollY = this.scrollY): number {
    const s = this.sheet.rows;
    if (y < this.cellsY && this.fr > 0 && y >= this.headerH) return Math.min(this.fr - 1, s.indexAt((y - this.headerH) / this.z));
    return s.indexAt(s.offset(this.fr) + (Math.max(y, this.cellsY) - this.cellsY + scrollY) / this.z);
  }

  /** Kaydırılan bölmedeki ilk/son görünür satır ve sütun. */
  visibleRows(): [number, number] {
    const s = this.sheet.rows;
    const top = s.offset(this.fr) + this.scrollY / this.z;
    const r1 = Math.max(this.fr, s.indexAt(top));
    const r2 = Math.min(MAX_ROWS - 1, s.indexAt(top + this.paneH / this.z));
    return [r1, r2];
  }

  visibleCols(): [number, number] {
    const s = this.sheet.cols;
    const left = s.offset(this.fc) + this.scrollX / this.z;
    const c1 = Math.max(this.fc, s.indexAt(left));
    const c2 = Math.min(MAX_COLS - 1, s.indexAt(left + this.paneW / this.z));
    return [c1, c2];
  }

  quadrants(): Quadrant[] {
    const [r1, r2] = this.visibleRows();
    const [c1, c2] = this.visibleCols();
    const out: Quadrant[] = [];
    const main: Quadrant = { x: this.cellsX, y: this.cellsY, w: this.paneW, h: this.paneH, r1, r2, c1, c2, frozenRows: false, frozenCols: false };
    out.push(main);
    if (this.fr > 0) out.push({ x: this.cellsX, y: this.headerH, w: this.paneW, h: this.frozenH, r1: 0, r2: this.fr - 1, c1, c2, frozenRows: true, frozenCols: false });
    if (this.fc > 0) out.push({ x: this.headerW, y: this.cellsY, w: this.frozenW, h: this.paneH, r1, r2, c1: 0, c2: this.fc - 1, frozenRows: false, frozenCols: true });
    if (this.fr > 0 && this.fc > 0) out.push({ x: this.headerW, y: this.headerH, w: this.frozenW, h: this.frozenH, r1: 0, r2: this.fr - 1, c1: 0, c2: this.fc - 1, frozenRows: true, frozenCols: true });
    return out;
  }

  /** Hücrenin (ya da aralığın) ekran dikdörtgeni. */
  cellRect(r1: number, c1: number, r2 = r1, c2 = c1): { x: number; y: number; w: number; h: number } {
    const x = this.colX(c1);
    const y = this.rowY(r1);
    const x2 = this.colX(c2) + this.colW(c2);
    const y2 = this.rowY(r2) + this.rowH(r2);
    return { x, y, w: x2 - x, h: y2 - y };
  }

  /** Belirli bir hücreyi göstermek için gereken kaydırma. */
  scrollFor(row: number, col: number): { x: number; y: number } {
    let { scrollX: x, scrollY: y } = this;
    const z = this.z;
    if (row >= this.fr) {
      const top = (this.sheet.rows.offset(row) - this.sheet.rows.offset(this.fr)) * z;
      const bottom = top + this.sheet.rows.size(row) * z;
      if (top < y) y = top;
      else if (bottom > y + this.paneH) y = Math.max(top - (this.paneH - this.sheet.rows.size(row) * z), bottom - this.paneH);
    }
    if (col >= this.fc) {
      const left = (this.sheet.cols.offset(col) - this.sheet.cols.offset(this.fc)) * z;
      const right = left + this.sheet.cols.size(col) * z;
      if (left < x) x = left;
      else if (right > x + this.paneW) x = right - this.paneW;
    }
    return { x: Math.max(0, x), y: Math.max(0, y) };
  }

  /** Kaydırma çubukları için toplam içerik boyutu. */
  contentSize(dataRows: number, dataCols: number): { w: number; h: number } {
    const z = this.z;
    const rows = this.sheet.rows;
    const cols = this.sheet.cols;
    const h = (rows.offset(Math.min(MAX_ROWS, Math.max(dataRows, 0) + 1)) - rows.offset(this.fr)) * z;
    const w = (cols.offset(Math.min(MAX_COLS, Math.max(dataCols, 0) + 1)) - cols.offset(this.fc)) * z;
    return { w, h };
  }

  maxScroll(): { x: number; y: number } {
    const z = this.z;
    return {
      x: Math.max(0, (this.sheet.cols.total() - this.sheet.cols.offset(this.fc)) * z - this.paneW),
      y: Math.max(0, (this.sheet.rows.total() - this.sheet.rows.offset(this.fr)) * z - this.paneH),
    };
  }
}
