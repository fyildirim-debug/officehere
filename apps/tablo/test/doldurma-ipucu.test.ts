/** Doldurma tutamacı ipucu: sürüklerken gösterilen değer, bırakınca en uç hücreye yazılan değerle aynı. */
import { describe, expect, it } from "vitest";
import { formatValue } from "@officehere/formula";
import { Controller } from "../src/controller";
import { fillTipText } from "../src/grid/gridUtil";
import { rect } from "../src/model/rect";
import type { CellInput } from "../src/model/workbook";
import type { Rect } from "../src/model/types";

function put(ctl: Controller, cells: [number, number, CellInput][]) {
  ctl.ops.setInputs(ctl.sheet, cells.map(([row, col, input]) => ({ row, col, input })), "x");
}

/** İpucunu alır, doldurur ve en uç hücrenin görünen değerini döner. */
function tipAndResult(ctl: Controller, src: Rect, dst: Rect, end: [number, number], copy = false): [string, string] {
  const tip = fillTipText(ctl, src, dst, copy);
  ctl.ops.fill(ctl.sheet, src, dst, copy ? "copy" : "series");
  const v = ctl.wb.getValue(ctl.sheet, end[0], end[1]);
  return [tip, formatValue(v, ctl.wb.styleAt(ctl.sheet, end[0], end[1]).numFmt).text];
}

describe("doldurma ipucu", () => {
  it("5'er artan seride ipucu adımı kullanır (eski hata: son değer + satır sayısı)", () => {
    const ctl = new Controller();
    put(ctl, Array.from({ length: 24 }, (_, i) => [i, 1, (i + 1) * 5] as [number, number, CellInput]));
    const [tip, result] = tipAndResult(ctl, rect(0, 1, 23, 1), rect(0, 1, 27, 1), [27, 1]);
    expect(tip).toBe("140");
    expect(result).toBe("140");
  });

  it("iki değerden seri, yukarı ve sağa doğru da aynı", () => {
    const ctl = new Controller();
    put(ctl, [
      [10, 0, 10],
      [11, 0, 20],
    ]);
    expect(tipAndResult(ctl, rect(10, 0, 11, 0), rect(7, 0, 11, 0), [7, 0])).toEqual(["-20", "-20"]);
    put(ctl, [
      [0, 3, 2],
      [0, 4, 4],
    ]);
    expect(tipAndResult(ctl, rect(0, 3, 0, 4), rect(0, 3, 0, 8), [0, 8])).toEqual(["12", "12"]);
  });

  it("tek sayı kopyalanır; Ctrl ile seri kopyası", () => {
    const ctl = new Controller();
    put(ctl, [[0, 0, 7]]);
    expect(tipAndResult(ctl, rect(0, 0), rect(0, 0, 5, 0), [5, 0])).toEqual(["7", "7"]);
    put(ctl, [
      [0, 2, 1],
      [1, 2, 2],
    ]);
    // Kopya deseni tekrarlar: 1, 2 | 1, 2, 1
    expect(tipAndResult(ctl, rect(0, 2, 1, 2), rect(0, 2, 4, 2), [4, 2], true)).toEqual(["1", "1"]);
  });

  it("metin+sayı, ay adları ve sayı biçimi", () => {
    const ctl = new Controller();
    put(ctl, [
      [0, 0, "Ürün 1"],
      [0, 1, "Ocak"],
    ]);
    expect(tipAndResult(ctl, rect(0, 0), rect(0, 0, 3, 0), [3, 0])).toEqual(["Ürün 4", "Ürün 4"]);
    expect(tipAndResult(ctl, rect(0, 1), rect(0, 1, 13, 1), [13, 1])).toEqual(["Şubat", "Şubat"]);
    put(ctl, [
      [0, 5, 0.1],
      [1, 5, 0.2],
    ]);
    ctl.ops.applyStyle(ctl.sheet, [rect(0, 5, 1, 5)], { numFmt: "0%" });
    expect(tipAndResult(ctl, rect(0, 5, 1, 5), rect(0, 5, 3, 5), [3, 5])).toEqual(["40%", "40%"]);
  });

  it("formül kopyalanırken ve içeri sürüklemede ipucu yok", () => {
    const ctl = new Controller();
    put(ctl, [
      [0, 0, 1],
      [0, 1, { formula: "=A1*5" }],
    ]);
    expect(fillTipText(ctl, rect(0, 1), rect(0, 1, 4, 1))).toBe("");
    put(ctl, [
      [0, 2, 1],
      [1, 2, 2],
      [2, 2, 3],
    ]);
    expect(fillTipText(ctl, rect(0, 2, 2, 2), rect(0, 2, 1, 2))).toBe("");
  });
});
