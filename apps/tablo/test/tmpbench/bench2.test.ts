import { it } from "vitest";
import { readFileSync, appendFileSync } from "node:fs";
import { readXlsx } from "../../src/io/xlsx";

const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/bench2.log", a.map(String).join(" ") + "\n");
it("bench parts", async () => {
  const buf = readFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/kabul-tablo/files_in/buyuk50k.xlsx");
  const wb = await readXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
  const s = wb.sheets[0];
  log("cells", wb.engine.exportInputs().sheets[0].cells.length, "styles", s.styles.size);
  let t = performance.now();
  const snap = wb.snapshot();
  log("snapshot", performance.now() - t);
  t = performance.now();
  wb.engine.insertRows(s.id, 10, 5);
  log("engine.insertRows", performance.now() - t);
  t = performance.now();
  s.shiftStructure("row", 10, 5);
  log("shiftStructure", performance.now() - t);
  t = performance.now();
  wb.engine.recalc();
  log("recalc", performance.now() - t);
  t = performance.now();
  wb.restore(snap);
  log("restore", performance.now() - t);
}, 600000);
