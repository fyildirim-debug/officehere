import { it } from "vitest";
import { appendFileSync } from "node:fs";
import { FormulaEngine } from "../../src/engine";
import { adjustFormula } from "../../src/transform";
import { safeParse } from "../../src/engine-types";
const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/fb.log", a.map(String).join(" ") + "\n");
it("b", () => {
  const e = new FormulaEngine();
  const s = e.addSheet("S");
  for (let r = 0; r < 50000; r++) { e.setValue(s, r, 0, r); e.setValue(s, r, 1, r * 2); }
  e.autoCalc = false;
  for (let r = 0; r < 50000; r++) e.setFormula(s, r, 2, `=A${r + 1}*B${r + 1}`);
  e.recalc(true);
  let t = performance.now();
  for (let r = 0; r < 50000; r++) adjustFormula(`A${r + 1}*B${r + 1}`, "S", { sheet: "S", axis: "row", at: 5, count: 3 });
  log("adjust", performance.now() - t);
  t = performance.now();
  for (let r = 0; r < 50000; r++) safeParse(`A${r + 1}*B${r + 1}`);
  log("parse", performance.now() - t);
  t = performance.now();
  e.reregisterAll();
  log("rereg", performance.now() - t);
  t = performance.now();
  e.markAllDirty(); e.recalc();
  log("recalc", performance.now() - t);
  t = performance.now();
  e.insertRows(s, 5, 3);
  log("insertRows", performance.now() - t);
});
