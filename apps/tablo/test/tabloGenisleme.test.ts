// Tablo genişlemesi: alta/sağa yapıştırma, doldurma tutamacı ve Tabloyu Yeniden Boyutlandır.
import { describe, expect, it } from "vitest";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";
import { fillAndGrow, pasteData } from "../src/actions/edit";
import { resizeTable } from "../src/actions/tableAuto";

function type(ctl: Controller, row: number, col: number, text: string): void {
  ctl.selectCell(row, col);
  ctl.startEdit(text);
  ctl.commitEdit(null);
}

/** A1:C4 tablosu (Bölge, Adet, Fiyat). */
function table(): Controller {
  const ctl = new Controller();
  const rows: (string | number)[][] = [
    ["Bölge", "Adet", "Fiyat"],
    ["Ege", 2, 10],
    ["Akdeniz", 3, 12],
    ["Marmara", 4, 15],
  ];
  ctl.ops.setInputs(ctl.sheet, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "örnek");
  D.createTable(ctl, { r1: 0, c1: 0, r2: 3, c2: 2 }, true);
  return ctl;
}

const ref = (ctl: Controller) => ({ ...ctl.sheet.tables[0].ref });
const v = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);

describe("tablonun hemen altına yapıştırma", () => {
  it("tablo yapıştırılan satırları kapsar; tek adımda geri alınır", () => {
    const ctl = table();
    const n = ctl.sheet.tables[0].name;
    type(ctl, 0, 5, `=TOPLA(${n}[Adet])`);
    expect(v(ctl, 0, 5)).toBe(9);
    ctl.selectCell(4, 0);
    pasteData(ctl, "Karadeniz\t5\t20\r\nİç Anadolu\t6\t25", undefined);
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 2 });
    expect(v(ctl, 0, 5)).toBe(20);
    ctl.ops.undo();
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 });
    expect(v(ctl, 4, 0)).toBeNull();
    expect(v(ctl, 0, 5)).toBe(9);
    ctl.ops.redo();
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 5, c2: 2 });
  });

  it("hesaplanmış sütun yeni satırlarda sürer", () => {
    const ctl = table();
    type(ctl, 0, 3, "Tutar");
    type(ctl, 1, 3, "=[@Adet]*[@Fiyat]");
    expect(v(ctl, 3, 3)).toBe(60);
    ctl.selectCell(4, 0);
    pasteData(ctl, "Karadeniz\t5\t20", undefined);
    expect(ref(ctl).r2).toBe(4);
    expect(v(ctl, 4, 3)).toBe(100);
  });

  it("bitişik olmayan satıra ya da toplam satırlı tablonun altına yapıştırma genişletmez", () => {
    const ctl = table();
    ctl.selectCell(5, 0);
    pasteData(ctl, "Uzak\t1\t1", undefined);
    expect(ref(ctl).r2).toBe(3);
    const t = table();
    D.updateTable(t, t.sheet.tables[0], { totalRow: true });
    const r2 = ref(t).r2;
    t.selectCell(r2 + 1, 0);
    pasteData(t, "Karadeniz\t5\t20", undefined);
    expect(ref(t).r2).toBe(r2);
  });

  it("sağdaki sütuna yapıştırma tabloya sütun ekler; başlık sütun adı olur", () => {
    const ctl = table();
    ctl.selectCell(0, 3);
    pasteData(ctl, "Stok\r\n1\r\n2\r\n3", undefined);
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 3, c2: 3 });
    expect(ctl.sheet.tables[0].columns).toEqual(["Bölge", "Adet", "Fiyat", "Stok"]);
  });
});

describe("doldurma tutamacı", () => {
  it("son satırdan aşağı doldurma tabloyu büyütür; koşullu biçim aralığı da büyür", () => {
    const ctl = table();
    ctl.sheet.condFormats.push({ ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], rules: [] } as never);
    fillAndGrow(ctl, ctl.sheet, { r1: 3, c1: 0, r2: 3, c2: 2 }, { r1: 3, c1: 0, r2: 6, c2: 2 }, "copy");
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 6, c2: 2 });
    expect(v(ctl, 6, 0)).toBe("Marmara");
    expect(ctl.sheet.condFormats[0].ranges[0]).toEqual({ r1: 1, c1: 1, r2: 6, c2: 1 });
    ctl.ops.undo();
    expect(ref(ctl).r2).toBe(3);
    expect(v(ctl, 6, 0)).toBeNull();
  });

  it("tablonun dışındaki sütunlarda doldurma tabloya dokunmaz", () => {
    const ctl = table();
    type(ctl, 3, 5, "x");
    fillAndGrow(ctl, ctl.sheet, { r1: 3, c1: 5, r2: 3, c2: 5 }, { r1: 3, c1: 5, r2: 6, c2: 5 }, "copy");
    expect(ref(ctl).r2).toBe(3);
  });
});

describe("Tabloyu Yeniden Boyutlandır", () => {
  it("aşağı ve sağa büyür, küçülür; yeni sütunlar ad alır; geri alınır", () => {
    const ctl = table();
    const t = ctl.sheet.tables[0];
    expect(resizeTable(ctl, t, { r1: 0, c1: 0, r2: 9, c2: 4 })).toBeNull();
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 9, c2: 4 });
    expect(ctl.sheet.tables[0].columns).toEqual(["Bölge", "Adet", "Fiyat", "Sütun4", "Sütun5"]);
    expect(v(ctl, 0, 3)).toBe("Sütun4");
    expect(resizeTable(ctl, ctl.sheet.tables[0], { r1: 0, c1: 0, r2: 2, c2: 1 })).toBeNull();
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 });
    expect(ctl.sheet.tables[0].columns).toEqual(["Bölge", "Adet"]);
    ctl.ops.undo();
    ctl.ops.undo();
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 });
    expect(ctl.sheet.tables[0].columns).toEqual(["Bölge", "Adet", "Fiyat"]);
    expect(v(ctl, 0, 3)).toBeNull();
  });

  it("geçersiz aralıklar nedeniyle reddedilir", () => {
    const ctl = table();
    const t = ctl.sheet.tables[0];
    expect(resizeTable(ctl, t, { r1: 1, c1: 0, r2: 9, c2: 2 })).toMatch(/aynı satırda/);
    expect(resizeTable(ctl, t, { r1: 0, c1: 0, r2: 0, c2: 2 })).toMatch(/en az bir satır/);
    expect(resizeTable(ctl, t, { r1: 0, c1: 5, r2: 9, c2: 6 })).toMatch(/örtüşmelidir/);
    ctl.ops.setInputs(ctl.sheet, [{ row: 8, col: 0, input: "Ad" }, { row: 9, col: 0, input: "x" }], "örnek");
    D.createTable(ctl, { r1: 8, c1: 0, r2: 9, c2: 0 }, true);
    expect(resizeTable(ctl, ctl.sheet.tables[0], { r1: 0, c1: 0, r2: 9, c2: 2 })).toMatch(/çakışamaz/);
    D.updateTable(ctl, ctl.sheet.tables[0], { totalRow: true });
    expect(resizeTable(ctl, ctl.sheet.tables[0], { r1: 0, c1: 0, r2: 6, c2: 2 })).toMatch(/Toplam Satırı/);
    expect(ref(ctl)).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
  });
});
