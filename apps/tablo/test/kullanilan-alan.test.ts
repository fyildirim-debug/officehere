/**
 * usedRange önbelleği: ızgara her karede çağırdığı için sonuç önbelleğe alınır. Bu testler
 * önbelleğin her değişiklik türünde geçersizlendiğini (değer, stil, satır/sütun ekle-sil, taşıma,
 * sıralama, geri al/yinele, sayfa, dosya açma) ve değişiklik yokken yeniden tarama yapılmadığını denetler.
 */
import { describe, expect, it, vi } from "vitest";
import { cellKey } from "@officehere/formula";
import { Controller } from "../src/controller";
import { WorkbookModel } from "../src/model/workbook";
import { rect } from "../src/model/rect";

const used = (ctl: Controller) => ctl.wb.usedRange(ctl.sheet);
const yaz = (ctl: Controller, row: number, col: number, input: string | number | null | { formula: string }) => ctl.ops.setInputs(ctl.sheet, [{ row, col, input }], "yaz");

describe("usedRange önbelleği", () => {
  it("hücre yazınca büyür, temizlenince ve silinince küçülür", () => {
    const ctl = new Controller();
    expect(used(ctl)).toBeNull();
    yaz(ctl, 2, 1, "a");
    expect(used(ctl)).toEqual({ r1: 2, c1: 1, r2: 2, c2: 1 });
    yaz(ctl, 9, 4, 5);
    expect(used(ctl)).toEqual({ r1: 2, c1: 1, r2: 9, c2: 4 });
    // İçeriği temizle
    ctl.ops.clear(ctl.sheet, [rect(9, 4)], "contents");
    expect(used(ctl)).toEqual({ r1: 2, c1: 1, r2: 2, c2: 1 });
    // Değer null yazmak hücreyi siler
    yaz(ctl, 2, 1, null);
    expect(used(ctl)).toBeNull();
  });

  it("stil uygulanınca büyür, biçimler temizlenince küçülür", () => {
    const ctl = new Controller();
    yaz(ctl, 0, 0, 1);
    ctl.ops.applyStyle(ctl.sheet, [rect(20, 6)], { bold: true });
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 20, c2: 6 });
    ctl.ops.clear(ctl.sheet, [rect(20, 6)], "formats");
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
  });

  it("satır ve sütun ekle-sil alanı kaydırır", () => {
    const ctl = new Controller();
    yaz(ctl, 1, 1, "x");
    yaz(ctl, 4, 3, "y");
    ctl.ops.insertRows(ctl.sheet, 0, 3);
    expect(used(ctl)).toEqual({ r1: 4, c1: 1, r2: 7, c2: 3 });
    ctl.ops.deleteRows(ctl.sheet, 5, 3);
    expect(used(ctl)).toEqual({ r1: 4, c1: 1, r2: 4, c2: 1 });
    ctl.ops.insertCols(ctl.sheet, 0, 2);
    expect(used(ctl)).toEqual({ r1: 4, c1: 3, r2: 4, c2: 3 });
    ctl.ops.deleteCols(ctl.sheet, 0, 1);
    expect(used(ctl)).toEqual({ r1: 4, c1: 2, r2: 4, c2: 2 });
    // Dolu satırı silmek alanı boşaltır
    ctl.ops.deleteRows(ctl.sheet, 4, 1);
    expect(used(ctl)).toBeNull();
  });

  it("geri al / yinele alanı geri getirir", () => {
    const ctl = new Controller();
    yaz(ctl, 0, 0, 1);
    yaz(ctl, 30, 2, 2);
    expect(used(ctl)?.r2).toBe(30);
    ctl.ops.undo();
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    ctl.ops.redo();
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 30, c2: 2 });
    ctl.ops.insertRows(ctl.sheet, 0, 10);
    expect(used(ctl)?.r2).toBe(40);
    ctl.ops.undo();
    expect(used(ctl)?.r2).toBe(30);
    ctl.ops.applyStyle(ctl.sheet, [rect(50, 0)], { bold: true });
    expect(used(ctl)?.r2).toBe(50);
    ctl.ops.undo();
    expect(used(ctl)?.r2).toBe(30);
  });

  it("her sayfanın alanı ayrı tutulur; sayfa değişince doğru alan döner", () => {
    const ctl = new Controller();
    yaz(ctl, 3, 3, "a");
    const s2 = ctl.ops.addSheet();
    ctl.ops.setInputs(s2, [{ row: 7, col: 0, input: "b" }], "yaz");
    const s1 = ctl.wb.sheets[0];
    expect(ctl.wb.usedRange(s1)).toEqual({ r1: 3, c1: 3, r2: 3, c2: 3 });
    expect(ctl.wb.usedRange(s2)).toEqual({ r1: 7, c1: 0, r2: 7, c2: 0 });
    ctl.ops.setInputs(s2, [{ row: 9, col: 9, input: "c" }], "yaz");
    expect(ctl.wb.usedRange(s1)).toEqual({ r1: 3, c1: 3, r2: 3, c2: 3 });
    expect(ctl.wb.usedRange(s2)).toEqual({ r1: 7, c1: 0, r2: 9, c2: 9 });
  });

  it("dosya açılınca yeni kitabın alanı döner", () => {
    const ctl = new Controller();
    yaz(ctl, 5, 5, "eski");
    expect(used(ctl)?.r2).toBe(5);
    const wb = new WorkbookModel();
    wb.setInputRaw(wb.active, 0, 0, "yeni");
    wb.setInputRaw(wb.active, 2, 1, "yeni");
    ctl.load(wb, "yeni.xlsx", null);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 });
  });

  it("ham yazıcılar (setInputRaw/setStyleRaw) yeniden hesaplamasız da alanı günceller", () => {
    const wb = new WorkbookModel();
    const s = wb.active;
    wb.setInputRaw(s, 1, 1, 1);
    expect(wb.usedRange(s)).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    wb.setStyleRaw(s, 6, 2, 1);
    expect(wb.usedRange(s)).toEqual({ r1: 1, c1: 1, r2: 6, c2: 2 });
    // Aynı sayıda stil kaydı, farklı konum (biri silinir, biri eklenir)
    wb.setStyleRaw(s, 6, 2, undefined);
    wb.setStyleRaw(s, 3, 8, 1);
    expect(wb.usedRange(s)).toEqual({ r1: 1, c1: 1, r2: 3, c2: 8 });
    wb.setInputRaw(s, 1, 1, null);
    expect(wb.usedRange(s)).toEqual({ r1: 3, c1: 8, r2: 3, c2: 8 });
  });

  it("stil haritasına doğrudan ekleme ya da haritanın yenilenmesi de alanı günceller", () => {
    const wb = new WorkbookModel();
    const s = wb.active;
    wb.setInputRaw(s, 0, 0, 1);
    expect(wb.usedRange(s)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    // Okuyucular ve satır ekleme biçim kopyası haritaya doğrudan yazar
    s.styles.set(cellKey(12, 3), 1);
    expect(wb.usedRange(s)).toEqual({ r1: 0, c1: 0, r2: 12, c2: 3 });
    // Geri alma (restoreData) ve tablo oluşturma haritayı yenisiyle değiştirir (aynı boyut)
    s.styles = new Map([[cellKey(4, 9), 1]]);
    expect(wb.usedRange(s)).toEqual({ r1: 0, c1: 0, r2: 4, c2: 9 });
  });

  it("el ile hesaplamada taşıma ve sıralama da alanı günceller", () => {
    const ctl = new Controller();
    const wb = ctl.wb;
    const s = ctl.sheet;
    wb.engine.autoCalc = false;
    yaz(ctl, 0, 0, "a");
    yaz(ctl, 5, 5, "b");
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 5 });
    // Kes-yapıştır: hücre sayısı ve en büyük sınır aynı kalır, sol üst köşe değişir
    wb.engine.moveRange(s.id, 0, 0, 0, 0, s.id, 2, 2);
    expect(used(ctl)).toEqual({ r1: 2, c1: 2, r2: 5, c2: 5 });
    wb.engine.moveRange(s.id, 2, 2, 2, 2, s.id, -2, -2);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 5 });
    // Sıralama: 0..5 satırları; boş satırlar sona gider
    expect(ctl.ops.permuteRowsOp(s, rect(0, 0, 5, 5), [0, 5, 1, 2, 3, 4], "Sırala")).toBe(true);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 1, c2: 5 });
    ctl.ops.undo();
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 5 });
    // Yapısal işlem (yeniden hesaplama yok)
    ctl.ops.insertRows(s, 0, 2);
    expect(used(ctl)).toEqual({ r1: 2, c1: 0, r2: 7, c2: 5 });
    // Yalnız motor düzeyinde (stil haritası yenilenmeden)
    wb.engine.deleteRows(s.id, 0, 2);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 5 });
  });

  it("taşan dizi formülü yeniden hesaplanınca alan büyür ve küçülür", () => {
    const ctl = new Controller();
    yaz(ctl, 0, 0, 3);
    yaz(ctl, 0, 1, { formula: "=SEQUENCE(A1)" });
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 });
    yaz(ctl, 0, 0, 12);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 11, c2: 1 });
    yaz(ctl, 0, 0, 1);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 1 });
    // El ile hesaplama: taşma yalnız yeniden hesaplamada (F9) değişir
    ctl.wb.engine.autoCalc = false;
    yaz(ctl, 0, 0, 6);
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 1 });
    ctl.wb.engine.recalc();
    expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 1 });
  });

  it("değişiklik yokken yeniden taranmaz (dosya açıldıktan sonra da)", () => {
    const wb = new WorkbookModel();
    for (let r = 0; r < 100; r++) wb.setInputRaw(wb.active, r, r % 7, r);
    const ctl = new Controller();
    ctl.load(wb, "kitap.xlsx", null);
    const tara = vi.spyOn(ctl.wb.engine, "forEachValue");
    for (let i = 0; i < 50; i++) expect(used(ctl)).toEqual({ r1: 0, c1: 0, r2: 99, c2: 6 });
    expect(tara).toHaveBeenCalledTimes(1);
    // Seçim ve kaydırma gibi içerik dışı değişiklikler önbelleği bozmaz
    ctl.selectCell(500, 3);
    used(ctl);
    expect(tara).toHaveBeenCalledTimes(1);
    yaz(ctl, 200, 0, "son");
    expect(used(ctl)?.r2).toBe(200);
    used(ctl);
    expect(tara).toHaveBeenCalledTimes(2);
  });

  it("döndürülen dikdörtgeni değiştirmek önbelleği bozmaz", () => {
    const ctl = new Controller();
    yaz(ctl, 4, 4, "a");
    const u = used(ctl)!;
    u.r2 = 999;
    expect(used(ctl)).toEqual({ r1: 4, c1: 4, r2: 4, c2: 4 });
  });
});
