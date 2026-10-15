import { describe, expect, it } from "vitest";
import { Errors, FormulaEngine, cellKey, keyCol, keyRow, type Scalar } from "../src";

/** Olağan yolla (setValue/setFormula + tam hesap) kurulmuş motor. */
function regular(cells: Record<string, Scalar | { f: string }>[], names: [string, string][] = []): FormulaEngine {
  const e = new FormulaEngine();
  const ids = cells.map((_, i) => e.addSheet(`S${i + 1}`));
  for (const [n, f] of names) e.defineName(n, f);
  cells.forEach((sheet, i) => {
    for (const [a, v] of Object.entries(sheet)) {
      const { row, col } = at(a);
      if (v !== null && typeof v === "object" && "f" in v) e.setFormula(ids[i], row, col, v.f);
      else e.setValue(ids[i], row, col, v as Scalar);
    }
  });
  e.recalc(true);
  return e;
}

function at(a: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Motorun girdilerini ve değerlerini sayfa başına düz nesneye çevirir. */
function dump(e: FormulaEngine): Record<string, [Scalar, string | undefined]>[] {
  return e.sheetIds().map((id) => {
    const out: Record<string, [Scalar, string | undefined]> = {};
    e.forEachCell(id, (r, c, rec) => (out[`${r},${c}`] = [rec.value, rec.formula?.text]));
    return out;
  });
}

/** Aynı içeriği bulkLoad ile kurar: önbellek değerleri olağan motordan alınır. */
function bulk(src: FormulaEngine, recompute: (sheet: number, row: number, col: number) => boolean = () => false): FormulaEngine {
  const e = new FormulaEngine();
  for (const id of src.sheetIds()) e.addSheet(src.getSheetName(id)!);
  for (const n of src.getNames()) e.defineName(n.name, n.formula, n.sheet);
  for (const id of src.sheetIds()) {
    const load = e.bulkLoad(id);
    for (const [k, rec] of src.cellEntries(id)) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (rec.formula) load.formula(r, c, rec.formula.text, recompute(id, r, c) ? undefined : rec.value);
      else load.value(r, c, rec.value);
    }
    load.finish();
    expect(load.count).toBe(src.cellCount(id));
  }
  return e;
}

describe("toplu yükleme (bulkLoad)", () => {
  it("önbellekli formüller yeniden hesaplanmadan aynı motoru kurar; bağımlılıklar çalışır", () => {
    const src = regular([{ A1: 1, A2: { f: "A1*2" }, A3: { f: "SUM(A1:A2)+S2!B1" }, B1: "metin", C1: true }, { B1: 10, B2: { f: "S1!A3+Oran" } }], [["Oran", "=0.5"]]);
    const e = bulk(src);
    expect(e.pending).toBe(false);
    expect(dump(e)).toEqual(dump(src));
    // Bağımlılık grafiği kurulmuş: girdi değişince zincir güncellenir
    e.setValue(1, 0, 0, 5);
    e.recalc();
    expect(e.getValue(1, 1, 0)).toBe(10);
    expect(e.getValue(1, 2, 0)).toBe(25);
    expect(e.getValue(2, 1, 1)).toBe(25.5);
  });

  it("önbellek değeri yeniden hesaplanmaz (dosyadaki değer korunur)", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("S");
    const load = e.bulkLoad(s);
    load.value(0, 0, 2);
    load.formula(0, 1, "A1*3", 99);
    load.finish();
    expect(e.getValue(s, 0, 1)).toBe(99);
    e.recalc(true);
    expect(e.getValue(s, 0, 1)).toBe(6);
  });

  it("önbelleksiz formül kirlidir; ayrıştırılamayan formül #AD? olur", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("S");
    const load = e.bulkLoad(s);
    load.value(0, 0, 4);
    load.formula(1, 0, "A1+1");
    load.formula(2, 0, "(((", undefined);
    load.finish();
    expect(e.pending).toBe(true);
    expect(e.getValue(s, 1, 0)).toBe(5);
    expect(e.getValue(s, 2, 0)).toEqual(Errors.NAME);
  });

  it("taşan formül önbelleksiz verilince taşma alanı hesapla kurulur", () => {
    const src = regular([{ A1: { f: "SEQUENCE(3)" }, B1: { f: "SUM(A1:A3)" } }]);
    const e = bulk(src, (_s, r, c) => r === 0 && c === 0);
    e.recalc();
    expect(e.getValue(1, 2, 0)).toBe(3);
    expect(e.getSpill(1, 1, 0)).toEqual({ anchorRow: 0, anchorCol: 0, rows: 3, cols: 1 });
    expect(e.getValue(1, 0, 1)).toBe(6);
    expect(dump(e)).toEqual(dump(src));
  });

  it("döngüsel başvurudaki hücreler önbelleksiz verilince döngü bilgisi kurulur", () => {
    const src = regular([{ A1: 1, D1: { f: "E1+1" }, E1: { f: "D1+1" }, F1: { f: "D1+A1" } }]);
    const circ = new Set(src.getCircularRefs().map((c) => cellKey(c.row, c.col)));
    expect(circ.size).toBe(2);
    const e = bulk(src, (_s, r, c) => circ.has(cellKey(r, c)));
    e.recalc();
    expect(e.getCircularRefs().length).toBe(2);
    expect(dump(e)).toEqual(dump(src));
  });

  it("girdi sürümü yalnız sonda bir kez artar; kullanılan alan büyür", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("S");
    const v0 = e.inputVersion;
    const load = e.bulkLoad(s);
    for (let r = 0; r < 1000; r++) load.value(r, r % 7, r);
    expect(e.inputVersion).toBe(v0);
    load.finish();
    expect(e.inputVersion).toBe(v0 + 1);
    expect(e.getUsedExtent(s)).toEqual({ rows: 1000, cols: 7 });
    expect(() => load.value(0, 0, 1)).toThrow();
  });

  it("dolu sayfada olağan yazma gibi davranır: eski formül kaydı silinir, bağımlılar güncellenir", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("S");
    e.setValue(s, 0, 0, 1);
    e.setFormula(s, 0, 1, "=A1*10");
    e.setFormula(s, 0, 2, "=B1+1");
    e.recalc();
    const load = e.bulkLoad(s);
    load.value(0, 1, 7);
    load.value(0, 0, 3);
    load.finish();
    e.recalc();
    expect(e.getFormula(s, 0, 1)).toBeUndefined();
    expect(e.getValue(s, 0, 2)).toBe(8);
    // Eski formülün bağımlılığı silindi: A1 değişince B1 yeniden hesaplanmaz
    e.setValue(s, 0, 0, 100);
    e.recalc();
    expect(e.getValue(s, 0, 1)).toBe(7);
  });

  it("cellEntries parça parça gezilebilir", () => {
    const src = regular([{ A1: 1, B2: "x", C3: { f: "A1+1" } }]);
    const it = src.cellEntries(1);
    const first = it.next().value!;
    expect([keyRow(first[0]), keyCol(first[0]), first[1].value]).toEqual([0, 0, 1]);
    expect([...it].map(([k]) => [keyRow(k), keyCol(k)])).toEqual([
      [1, 1],
      [2, 2],
    ]);
    expect([...src.cellEntries(99)]).toEqual([]);
  });
});
