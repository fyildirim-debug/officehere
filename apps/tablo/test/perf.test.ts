/** Büyük veri performansı (model düzeyi): 50.000 satırda filtre, sıralama, gizle/göster, satır ekle/sil. */
import { afterAll, describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
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

/** Ölçülen süreler (rapor için PERF_LOG ortam değişkeniyle dosyaya yazılır). */
const timings: Record<string, number> = {};
function time(fn: () => void, label = ""): number {
  const t = performance.now();
  fn();
  const ms = performance.now() - t;
  if (label) timings[label] = Math.round(ms);
  return ms;
}
afterAll(() => {
  const out = process.env.PERF_LOG;
  if (out) writeFileSync(out, JSON.stringify(timings, null, 1));
});

// Eşikler bilinçli olarak geniş: testler pnpm -r ve CI'da diğer paketlerle paralel koşarken
// süreler birkaç kat uzayabilir. Amaç, eski 170 sn'lik filtre gerilemesi gibi büyüklük
// mertebesindeki yavaşlamaları yakalamak.
const LIMIT_FAST = 5000;
const LIMIT_SLOW = 12000;

describe("50.000 satır performansı", { timeout: 60_000 }, () => {
  const ctl = big();
  const s = ctl.sheet;
  const total = ctl.wb.getValue(s, 0, 7) as number;

  it("otomatik filtre uygula / temizle hızlı (tek başına < 0,4 sn)", () => {
    ctl.selectCell(0, 2);
    D.toggleAutoFilter(ctl);
    const f = s.autoFilter!;
    expect(f.range.r2).toBe(N);
    const apply = time(() => D.setColumnFilter(ctl, f, 2, { kind: "values", values: ["İzmir"], blanks: false }), "filtre uygula");
    const visible = N - s.filterHidden.size;
    expect(visible).toBe(Math.floor((N + 5) / 7));
    expect(s.rows.isHidden(1)).toBe(true);
    // ALTTOPLAM yalnız görünen satırları toplar
    expect(ctl.wb.getValue(s, 0, 7)).toBeLessThan(total);
    const clear = time(() => D.clearAllFilters(ctl), "filtre temizle");
    expect(s.filterHidden.size).toBe(0);
    expect(s.rows.hiddenCount()).toBe(0);
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(apply).toBeLessThan(LIMIT_FAST);
    expect(clear).toBeLessThan(LIMIT_FAST);
    // Geri al: filtre geri gelir
    const undo = time(() => ctl.ops.undo(), "filtre geri al");
    expect(s.filterHidden.size).toBe(N - visible);
    expect(undo).toBeLessThan(LIMIT_FAST);
    ctl.ops.redo();
  });

  it("satır gizle / göster hızlı (tek başına < 0,05 sn)", () => {
    const rows = Array.from({ length: 40_000 }, (_, i) => i + 5);
    const hide = time(() => ctl.ops.setHidden(s, "row", rows, true), "40k satır gizle");
    expect(s.rows.isHidden(100)).toBe(true);
    // ALTTOPLAM(109) elle gizlenen satırları da dışarıda bırakır
    expect(ctl.wb.getValue(s, 0, 7)).toBeLessThan(total);
    const show = time(() => ctl.ops.setHidden(s, "row", rows, false), "40k satır göster");
    expect(s.rows.isHidden(100)).toBe(false);
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(hide).toBeLessThan(LIMIT_FAST);
    expect(show).toBeLessThan(LIMIT_FAST);
  });

  it("sıralama hızlı (tek başına ~1 sn)", () => {
    ctl.selectCell(1, 3);
    const t = time(() => D.quickSort(ctl, true), "sırala");
    expect(ctl.wb.getValue(s, 0, 2)).toBe("Şehir");
    expect(ctl.wb.getValue(s, 1, 3)).toBe(49);
    // Formüller satırlarıyla birlikte taşınır (göreli başvurular kaydırılır)
    const f = ctl.wb.getFormula(s, 1, 5)!;
    expect(f).toBe("=D2*E2");
    expect(ctl.wb.getValue(s, 1, 5)).toBe((ctl.wb.getValue(s, 1, 3) as number) * (ctl.wb.getValue(s, 1, 4) as number));
    expect(ctl.wb.getValue(s, 0, 7)).toBe(total);
    expect(t).toBeLessThan(LIMIT_SLOW);
    // Geri alma özgün sırayı geri getirir
    const u = time(() => ctl.ops.undo(), "sıralamayı geri al");
    expect(ctl.wb.getValue(s, 1, 3)).toBe(7);
    expect(ctl.wb.getValue(s, 1, 0)).toBe(1);
    expect(u).toBeLessThan(LIMIT_SLOW);
  });

  it("satır ekle / sil hızlı (tek başına ~0,8 sn)", () => {
    const ins = time(() => ctl.ops.insertRows(s, 10, 5), "satır ekle");
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 6})`);
    const del = time(() => ctl.ops.deleteRows(s, 10, 5), "satır sil");
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 1})`);
    expect(ins).toBeLessThan(LIMIT_SLOW);
    expect(del).toBeLessThan(LIMIT_SLOW);
  });
});
