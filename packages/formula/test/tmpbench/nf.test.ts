import { it } from "vitest";
import { appendFileSync } from "node:fs";
import { formatValue, formatFromLocal } from "../../src/numfmt";
import { FormulaEngine } from "../../src/engine";
const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/nf.log", a.map(String).join(" ") + "\n");
it("nf", () => {
  for (const c of ['0.0,,"M"', '0.0,,\\M', "#,##0,", "0,", "#,##0.0,,"]) log("canon", c, "=>", formatValue(1234567.891, c).text);
  for (const c of ['0,0..\\M', '#.##0,', '#.##0.', "0,0..", '[$-tr-TR]D MMMM YYYY DDDD', '[$-en-US]dddd', '[$-409]dddd']) log("local", c, "=>", formatFromLocal(c));
  log("tr-tag", formatValue(45000, "[$-tr-TR]D MMMM YYYY DDDD").text, "|", formatValue(45000, "[$-en-US]dddd").text, "|", formatValue(45000, "[$-409]mmmm").text);
  const e = new FormulaEngine(); const s = e.addSheet("S");
  e.setFormula(s, 0, 0, '=TEXT(1234567.891,"0,0..\\M")');
  e.setFormula(s, 1, 0, '=TEXT(1234567.891,"#.##0,")');
  e.setFormula(s, 2, 0, '=TEXT(DATE(2024,1,1),"[$-en-US]dddd")');
  e.recalc();
  log("TEXT", e.getValue(s, 0, 0), "|", e.getValue(s, 1, 0), "|", e.getValue(s, 2, 0));
});
