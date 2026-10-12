/** Sayfa işlemleri: ekle, sil, yeniden adlandır, taşı, kopyala, sekme rengi, gizle. */
import { OpsCore } from "./opsCore";
import { STRUCTURE_MESSAGE } from "./protection";
import type { SheetModel, WorkbookModel } from "./workbook";

export class OpsSheets extends OpsCore {
  /** Kitap yapısı korunurken engellenen işlem için ileti (denetleyici uyarı gösterir). */
  onBlocked: ((message: string) => void) | null = null;

  /** Çalışma kitabı yapı koruması: sayfa ekleme/silme/adlandırma/taşıma/gizleme engellenir. */
  protected blocked(): boolean {
    if (!this.wb.protection?.structure) return false;
    this.onBlocked?.(STRUCTURE_MESSAGE);
    return true;
  }

  /* --------------------------- Sayfalar --------------------------- */

  addSheet(index?: number): SheetModel {
    if (this.blocked()) return this.wb.active;
    let created!: SheetModel;
    this.snapshotOp("Sayfa Ekle", () => {
      created = this.wb.addSheet(undefined, index ?? this.wb.sheets.length);
      this.wb.activeIndex = this.wb.sheets.indexOf(created);
    });
    return created;
  }

  deleteSheet(index: number): void {
    if (this.blocked() || this.wb.sheets.filter((s) => !s.hidden).length <= 1) return;
    this.snapshotOp("Sayfayı Sil", () => {
      this.wb.removeSheet(index);
      this.wb.activeIndex = Math.min(index, this.wb.sheets.length - 1);
      while (this.wb.active.hidden && this.wb.activeIndex > 0) this.wb.activeIndex--;
    });
  }

  renameSheet(index: number, name: string): void {
    if (this.blocked()) return;
    this.snapshotOp("Sayfayı Yeniden Adlandır", () => this.wb.renameSheet(index, name.trim()));
  }

  moveSheet(from: number, to: number): void {
    if (from === to || this.blocked()) return;
    this.snapshotOp("Sayfayı Taşı", () => {
      const active = this.wb.active;
      this.wb.moveSheet(from, to);
      this.wb.activeIndex = this.wb.sheets.indexOf(active);
    });
  }

  copySheet(from: number, to: number): void {
    if (this.blocked()) return;
    this.snapshotOp("Sayfayı Kopyala", () => {
      const wb = this.wb;
      const src = wb.sheets[from];
      let base = src.name.replace(/ \(\d+\)$/, "");
      let n = 2;
      while (wb.sheetByName(`${base} (${n})`)) n++;
      base = `${base} (${n})`.slice(0, 31);
      const copy = wb.addSheet(base, to);
      copy.restoreData(src.cloneData());
      copy.charts = copy.charts.map((c) => ({ ...c, id: wb.newId("chart"), range: c.range.replace(new RegExp(`^'?${src.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}'?!`), `'${base}'!`) }));
      copy.images = copy.images.map((i) => ({ ...i, id: wb.newId("img") }));
      copy.tables = copy.tables.map((t) => ({ ...t, id: wb.newId("tbl"), name: uniqueTableName(wb, t.name) }));
      copy.pivots = copy.pivots.map((p) => {
        const names = new Set(wb.sheets.flatMap((s) => s.pivots.map((x) => x.name.toLocaleUpperCase("tr-TR"))));
        let n = 1;
        while (names.has(`PIVOTTABLE${n}`)) n++;
        return { ...p, id: wb.newId("pivot"), name: `ÖzetTablo${n}`, sourceSheet: p.sourceSheet === src.id ? copy.id : p.sourceSheet };
      });
      wb.engine.forEachCell(src.id, (row, col, rec) => {
        if (rec.formula) wb.engine.setFormula(copy.id, row, col, rec.formula.text);
        else wb.engine.setValue(copy.id, row, col, rec.value);
      });
      wb.activeIndex = wb.sheets.indexOf(copy);
    });
  }

  setTabColor(sheet: SheetModel, color: string | undefined): void {
    if (this.blocked()) return;
    this.sheetOp(sheet, "Sekme Rengi", () => {
      sheet.tabColor = color;
    });
  }

  setSheetHidden(index: number, hidden: boolean): void {
    if (this.blocked()) return;
    const wb = this.wb;
    if (hidden && wb.sheets.filter((s) => !s.hidden).length <= 1) return;
    this.snapshotOp(hidden ? "Sayfayı Gizle" : "Sayfayı Göster", () => {
      wb.sheets[index].hidden = hidden;
      if (hidden && wb.activeIndex === index) {
        const next = wb.sheets.findIndex((s) => !s.hidden);
        wb.activeIndex = next;
      }
      if (!hidden) wb.activeIndex = index;
    });
  }

}

export function uniqueTableName(wb: WorkbookModel, base = "Tablo"): string {
  const names = new Set(wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLocaleUpperCase("tr-TR"))));
  const stem = base.replace(/\d+$/, "") || "Tablo";
  let n = 1;
  while (names.has(`${stem}${n}`.toLocaleUpperCase("tr-TR"))) n++;
  return `${stem}${n}`;
}
