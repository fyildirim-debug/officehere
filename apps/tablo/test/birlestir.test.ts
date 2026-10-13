/** Birleştir (Consolidate): etiketlere ve konuma göre, sayfalar arası, işlevler ve geri alma. */
import { describe, expect, it } from "vitest";
import { isError } from "@officehere/formula";
import { Controller } from "../src/controller";
import { consolidate, parseSourceRef } from "../src/actions/consolidate";

function book() {
  const ctl = new Controller();
  const a = ctl.sheet;
  ctl.ops.renameSheet(0, "Ocak");
  const b = ctl.ops.addSheet();
  ctl.ops.renameSheet(1, "Şubat");
  const put = (s: typeof a, rows: (string | number | null)[][]) => ctl.ops.setInputs(s, rows.flatMap((r, i) => r.map((v, j) => ({ row: i, col: j, input: v }))), "x");
  put(a, [[null, "Adet", "Tutar"], ["Çay", 2, 10], ["Şeker", 3, 30]]);
  put(b, [[null, "Tutar", "Adet"], ["şeker", 5, 1], ["Kahve", 7, 2]]);
  const out = ctl.ops.addSheet();
  ctl.wb.activeIndex = ctl.wb.sheets.indexOf(out);
  ctl.selectCell(1, 1);
  return { ctl, out };
}
const v = (ctl: Controller, s: ReturnType<typeof book>["out"], r: number, c: number) => ctl.wb.getValue(s, r, c);

describe("Birleştir", () => {
  it("üst satır ve sol sütun etiketleriyle (büyük/küçük harf duyarsız) toplar", () => {
    const { ctl, out } = book();
    expect(consolidate(ctl, ["Ocak!$A$1:$C$3", "'Şubat'!A1:C3"], "SUM", true, true, { row: 1, col: 1 })).toBeNull();
    expect([v(ctl, out, 1, 2), v(ctl, out, 1, 3)]).toEqual(["Adet", "Tutar"]);
    expect([2, 3, 4].map((r) => v(ctl, out, r, 1))).toEqual(["Çay", "Şeker", "Kahve"]);
    expect([2, 3, 4].map((r) => [v(ctl, out, r, 2), v(ctl, out, r, 3)])).toEqual([[2, 10], [4, 35], [2, 7]]);
    expect(v(ctl, out, 1, 1)).toBeNull();
    expect(ctl.wb.getFormula(out, 3, 2)).toBeUndefined();
    ctl.ops.undo();
    expect(v(ctl, out, 3, 2)).toBeNull();
  });

  it("konuma göre birleştirir; ortalama, say ve sayıları say", () => {
    const { ctl, out } = book();
    consolidate(ctl, ["Ocak!B2:C3", "Şubat!B2:C3"], "AVERAGE", false, false, { row: 0, col: 0 });
    expect([[v(ctl, out, 0, 0), v(ctl, out, 0, 1)], [v(ctl, out, 1, 0), v(ctl, out, 1, 1)]]).toEqual([[3.5, 5.5], [5, 16]]);
    consolidate(ctl, ["Ocak!A1:C3", "Şubat!A1:C3"], "COUNT", false, false, { row: 5, col: 0 });
    // Hiçbir kaynakta değer olmayan konum boş kalır
    expect(v(ctl, out, 5, 0)).toBeNull();
    expect(v(ctl, out, 6, 0)).toBe(2);
    consolidate(ctl, ["Ocak!A1:C3", "Şubat!A1:C3"], "COUNTNUMS", false, false, { row: 10, col: 0 });
    expect(v(ctl, out, 11, 0)).toBe(0);
    expect(v(ctl, out, 11, 1)).toBe(2);
    consolidate(ctl, ["Ocak!B2:B2"], "STDEV", false, false, { row: 15, col: 0 });
    expect(isError(v(ctl, out, 15, 0))).toBe(true);
  });

  it("geçersiz başvuru ve boş liste hata iletisi verir", () => {
    const { ctl } = book();
    expect(consolidate(ctl, [], "SUM", false, false, { row: 0, col: 0 })).toMatch(/en az bir/);
    expect(consolidate(ctl, ["Yok!A1:B2"], "SUM", false, false, { row: 0, col: 0 })).toMatch(/geçerli değil/);
    expect(parseSourceRef(ctl, "=Ocak!$A$1:$B$2")?.rect).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
  });
});
