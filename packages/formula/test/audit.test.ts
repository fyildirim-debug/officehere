import { describe, expect, it } from "vitest";
import { FormulaEngine, FormulaStepper, dependentsOf, precedentsOf } from "../src";

function setup() {
  const e = new FormulaEngine();
  const s = e.addSheet("Sayfa1");
  const t = e.addSheet("Veri");
  e.setValue(s, 0, 0, 10); // A1
  e.setValue(s, 1, 0, 20); // A2
  e.setValue(s, 2, 0, 30); // A3
  e.setFormula(s, 0, 1, "=A1*2"); // B1
  e.setFormula(s, 1, 1, "=SUM(A1:A3)+B1"); // B2
  e.setValue(t, 0, 0, 5);
  e.setFormula(s, 2, 1, "=IF(A1>5,Veri!A1*2,0)"); // B3
  e.setFormula(s, 3, 1, "=(A1+A2)*A3/10"); // B4
  e.recalc();
  return { e, s, t };
}

describe("Formül denetimi", () => {
  it("etkileyenler: hücre ve aralıklar", () => {
    const { e, s } = setup();
    const p = precedentsOf(e, s, 1, 1);
    expect(p).toContainEqual({ sheet: s, r1: 0, c1: 0, r2: 2, c2: 0 });
    expect(p).toContainEqual({ sheet: s, r1: 0, c1: 1, r2: 0, c2: 1 });
    expect(precedentsOf(e, s, 0, 0)).toEqual([]);
  });

  it("etkilenenler: doğrudan bağımlı formüller, başka sayfa dahil", () => {
    const { e, s, t } = setup();
    expect(dependentsOf(e, s, 0, 0).map((d) => [d.row, d.col])).toEqual([
      [0, 1],
      [1, 1],
      [2, 1],
      [3, 1],
    ]);
    expect(dependentsOf(e, t, 0, 0)).toEqual([{ sheet: s, row: 2, col: 1 }]);
  });

  it("Formülü Değerlendir adımları", () => {
    const { e, s } = setup();
    const st = new FormulaStepper(e, s, 3, 1, e.getFormula(s, 3, 1)!);
    const views: string[] = [];
    let v = st.view();
    views.push(v.text);
    expect(v.mark).toEqual([2, 4]); // A1
    let guard = 0;
    while (!st.done && guard++ < 20) {
      st.step();
      views.push(st.view().text);
    }
    expect(views).toEqual(["=(A1+A2)*A3/10", "=(10+A2)*A3/10", "=(10+20)*A3/10", "=30*A3/10", "=30*30/10", "=900/10", "=90"]);
    expect(st.address()).toBe("Sayfa1!$B$4");
  });

  it("EĞER yalnız seçilen dalı değerlendirir; ETOPLA aralığı ve İçeri Adımla", () => {
    const { e, s } = setup();
    const st = new FormulaStepper(e, s, 2, 1, e.getFormula(s, 2, 1)!);
    const seen: string[] = [st.view().text];
    while (!st.done) {
      st.step();
      seen.push(st.view().text);
    }
    expect(seen[0]).toBe("=EĞER(A1>5;Veri!A1*2;0)");
    expect(seen).toContain("=EĞER(DOĞRU;Veri!A1*2;0)");
    expect(seen).toContain("=EĞER(DOĞRU;10;0)");
    expect(seen[seen.length - 1]).toBe("=10");
    const st2 = new FormulaStepper(e, s, 1, 1, e.getFormula(s, 1, 1)!);
    expect(st2.view().text).toBe("=TOPLA(A1:A3)+B1");
    expect(st2.view().mark).toEqual([1, 13]);
    st2.step();
    expect(st2.view().text).toBe("=60+B1");
    expect(st2.stepInTarget()).toEqual({ sheet: s, row: 0, col: 1, formula: "=A1*2" });
  });
});
