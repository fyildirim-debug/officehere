/**
 * Büyük veri performansı (model düzeyi): 50.000 satırda filtre, sıralama, gizle/göster, satır ekle/sil.
 *
 * Normal `pnpm test`te koşmaz (vite.config.ts dışarıda bırakır); CI'da ayrı ve sıralı adımda
 * `pnpm test:perf` ile koşar (vitest.perf.config.ts). Süreler makine hızına göre normalize edilir:
 * önce uygulama kodundan bağımsız sabit bir başvuru iş yükü ölçülür; oran = ölçülen / temeldeki.
 * Her işlem için bütçe = max(temel × 1,5; temel + 100 ms) × oran (kaydedilmiş temele göre %50
 * gerileme kuralı) ve ayrıca gevşek mutlak sınır × oran. Temel: test/perf-temel.json; güncelleme:
 * `pnpm --filter @officehere/tablo test:perf:temel` (yalnız bilinçli hızlanma/yavaşlamadan sonra).
 */
import { afterAll, describe, expect, it } from "vitest";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
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

/** Makine hızı başvurusu: uygulama kodundan bağımsız sabit iş yükü; 5 koşunun en kısası (ms). */
function referansOlc(): number {
  let enKisa = Infinity;
  for (let t = 0; t < 5; t++) {
    const t0 = performance.now();
    let tohum = 12345;
    const rnd = () => (tohum = (tohum * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const sayilar = Array.from({ length: 200_000 }, rnd).sort((a, b) => a - b);
    const nesne = Array.from({ length: 20_000 }, (_, i) => ({ i, ad: `Hücre ${i} ığüşöç`, d: sayilar[i * 10] }));
    const geri = JSON.parse(JSON.stringify(nesne)) as unknown[];
    const m = new Map<string, number>();
    for (let i = 0; i < 100_000; i++) m.set(`k${i % 5000}`, (m.get(`k${i % 5000}`) ?? 0) + i);
    if (geri.length + m.size < 0) throw new Error("olanaksız");
    enKisa = Math.min(enKisa, performance.now() - t0);
  }
  return enKisa;
}

type PerfTemel = { referansMs: number; olcumler: Record<string, number> };
const TEMEL_YOLU = new URL("./perf-temel.json", import.meta.url);
const temel: PerfTemel | null = existsSync(TEMEL_YOLU) ? JSON.parse(readFileSync(TEMEL_YOLU, "utf8")) : null;
const TEMEL_YAZ = import.meta.env.MODE === "temel";
const referansMs = referansOlc();
/** Bu makine temeldekinden kaç kat yavaş (0,5–5 arası; hızlı makinede bütçe en çok yarıya iner). */
const oran = temel ? Math.min(5, Math.max(0.5, referansMs / temel.referansMs)) : 1;

/** Ölçülen süreler (rapor için PERF_LOG ortam değişkeniyle dosyaya yazılır). */
const timings: Record<string, number> = {};
function time(fn: () => void, label = ""): number {
  const t = performance.now();
  fn();
  const ms = performance.now() - t;
  if (label) timings[label] = Math.round(ms);
  return ms;
}
/** Kaydedilmiş temele göre bütçe: %50'den fazla yavaşlama (en az 100 ms) gerileme sayılır. */
function butce(label: string, gevsek: number): number {
  const t = temel?.olcumler[label];
  const kural = t == null ? Infinity : Math.max(t * 1.5, t + 100) * oran;
  return Math.min(kural, gevsek * Math.max(1, oran));
}
function bekle(ms: number, label: string, gevsek: number) {
  expect(ms, `${label}: ${Math.round(ms)} ms (temel ${temel?.olcumler[label] ?? "-"} ms, hız oranı ${oran.toFixed(2)})`).toBeLessThan(butce(label, gevsek));
}
afterAll(() => {
  const out = process.env.PERF_LOG;
  if (out) writeFileSync(out, JSON.stringify({ referansMs: Math.round(referansMs * 10) / 10, oran, olcumler: timings }, null, 1));
  if (TEMEL_YAZ) {
    writeFileSync(TEMEL_YOLU, JSON.stringify({ aciklama: "pnpm --filter @officehere/tablo test:perf:temel ile üretilir", referansMs: Math.round(referansMs * 10) / 10, olcumler: timings }, null, 1) + "\n");
    console.log(`perf temeli yazıldı (referans ${referansMs.toFixed(1)} ms)`);
  }
});

// Gevşek mutlak sınırlar (temel yokken ya da ek güvence): büyüklük mertebesindeki yavaşlamaları
// (ör. eski 170 sn'lik filtre gerilemesi) yakalar.
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
    bekle(apply, "filtre uygula", LIMIT_FAST);
    bekle(clear, "filtre temizle", LIMIT_FAST);
    // Geri al: filtre geri gelir
    const undo = time(() => ctl.ops.undo(), "filtre geri al");
    expect(s.filterHidden.size).toBe(N - visible);
    bekle(undo, "filtre geri al", LIMIT_FAST);
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
    bekle(hide, "40k satır gizle", LIMIT_FAST);
    bekle(show, "40k satır göster", LIMIT_FAST);
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
    bekle(t, "sırala", LIMIT_SLOW);
    // Geri alma özgün sırayı geri getirir
    const u = time(() => ctl.ops.undo(), "sıralamayı geri al");
    expect(ctl.wb.getValue(s, 1, 3)).toBe(7);
    expect(ctl.wb.getValue(s, 1, 0)).toBe(1);
    bekle(u, "sıralamayı geri al", LIMIT_SLOW);
  });

  it("satır ekle / sil hızlı (tek başına ~0,8 sn)", () => {
    const ins = time(() => ctl.ops.insertRows(s, 10, 5), "satır ekle");
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 6})`);
    const del = time(() => ctl.ops.deleteRows(s, 10, 5), "satır sil");
    expect(ctl.wb.getFormula(s, 0, 7)).toBe(`=SUBTOTAL(109,F2:F${N + 1})`);
    bekle(ins, "satır ekle", LIMIT_SLOW);
    bekle(del, "satır sil", LIMIT_SLOW);
  });
});
