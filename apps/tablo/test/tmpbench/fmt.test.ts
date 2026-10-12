import { it } from "vitest";
import { readFileSync, appendFileSync } from "node:fs";
import { readXlsx } from "../../src/io/xlsx";
const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/fmt.log", a.map(String).join(" ") + "\n");
it("fmt", async () => {
  for (const f of ["karmasik.xlsx", "karmasik_lo.xlsx"]) {
    const b = readFileSync(__dirname + "/../fixtures/openpyxl/" + f);
    const wb = await readXlsx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
    const s = wb.sheetByName("Biçimler")!;
    for (let r = 0; r < 15; r++) log(f, r, JSON.stringify(wb.getValue(s, r, 0)), JSON.stringify(wb.getValue(s, r, 1)), wb.styleAt(s, r, 1).numFmt);
  }
});
