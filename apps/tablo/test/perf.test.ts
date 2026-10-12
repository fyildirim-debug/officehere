/** Büyük veri performansı (model düzeyi): 50.000 satırda filtre, sıralama, gizle/göster, satır ekle/sil. */
import { describe, expect, it } from "vitest";
import { Controller } from "../src/controller";
import * as D from "../src/actions/data";

const N = 50_000;
const CITIES = ["İstanbul", "Ankara", "İzmir", "Bursa", "Antalya", "Konya", "Çanakkale"];

function big(): Controller {
  const ctl = new Controller();
  const s = ctl.sheet;
  const wb = ctl.wb;
  wb.engine.autoCalc = false;
  const head = ["No", "Ad", "Şehir", "Adet", "Fiyat", "Tutar"];
  head.forEach((h, c) => wb.setInputRaw(s, 0, c, h));
  for (let r = 1; r <= N; r++) {
    wb.setInputRaw(s, r, 0, r);
    wb.setInputRaw(s, r, 1, `Müşteri ${r}`);
    wb.setInputRaw(s, r, 2, CITIES[r % CITIES.length]);
    wb.setInputRaw(s, r, 3, (r * 7) % 50);
    wb.setInputRaw(s, r, 4, 10 + (r % 13));
    wb.setInputRaw(s, r, 5, { formula: `=D${r + 1}*E${r + 1}` });
  }
  wb.setInputRaw(s, 0, 7, { formula: `=SUBTOTAL(109,F2:F${N + 1})` });
  wb.engine.autoCalc = true;
  wb.engine.recalc(true);
  return ctl;
}

function time(fn: () => void): number {
  const t = performance.now();
  fn();
  return performance.now() - t;
}

describe("50.000 satır performansı", () => {
  const ctl = big();
  const s = ctl.sheet;
  const total = ctl.wb.getValue(s, 0, 7) as number;

  it("otomatik filtre uygula / temizle < 1 sn", () => {
    ctl.selectCell(0, 2);
    D.toggleAutoFilter(ctl);
    const f = s.autoFilter!;
    expect(f.range.r2).toBe(N);
    const apply = time(() => D.setColumnFilter(ctl, f, 2, { kind: "values", values: ["İzmir"], blanks: false }));
    const visible = N - s.filterHidden.size;
    expect(visible).toBe(Math.floor((N + 5) / 7));
    expect(s.rows.isHidden(1)).toBe(true);
    // ALTTOPLAM yalnız görünen satırları toplar
    expect(ctl.wb.getValue(s, 0, 7)).toBeLessThan(total);
    const clear = time(() => D.clearAllFilters(ctl));
    expect(s.filterHidden.size).toBe(0);
    expect(s.rows.hiddenCount()).toBe(0);
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(apply).toBeLessThan(1000);
    expect(clear).toBeLessThan(1000);
    // Geri al: filtre geri gelir
    const undo = time(() => ctl.ops.undo());
    expect(s.filterHidden.size).toBe(N - visible);
    expect(undo).toBeLessThan(1000);
    ctl.ops.redo();
  });

  it("satır gizle / göster < 1 sn", () => {
    const rows = Array.from({ length: 40_000 }, (_, i) => i + 5);
    const hide = time(() => ctl.ops.setHidden(s, "row", rows, true));
    expect(s.rows.isHidden(100)).toBe(true);
    // ALTTOPLAM(109) elle gizlenen satırları da dışarıda bırakır
    expect(ctl.wb.getValue(s, 0, 7)).toBeLessThan(total);
    const show = time(() => ctl.ops.setHidden(s, "row", rows, false));
    expect(s.rows.isHidden(100)).toBe(false);
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(hide).toBeLessThan(1000);
    expect(show).toBeLessThan(1000);
  });

  it("sıralama < 2 sn", () => {
    ctl.selectCell(1, 3);
    const t = time(() => D.quickSort(ctl, true));
    expect(ctl.wb.getValue(s, 1, 3)).toBe(49);
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(t).toBeLessThan(2000);
  });

  it("satır ekle / sil < 1,5 sn", () => {
    const ins = time(() => ctl.ops.insertRows(s, 10, 5));
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 6})`);
    const del = time(() => ctl.ops.deleteRows(s, 10, 5));
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 1})`);
    expect(ins).toBeLessThan(1500);
    expect(del).toBeLessThan(1500);
  });
});
