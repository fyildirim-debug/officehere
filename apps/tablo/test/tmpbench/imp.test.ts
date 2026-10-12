import { it } from "vitest";
import { readFileSync, readdirSync, appendFileSync } from "node:fs";
import { readXlsx } from "../../src/io/xlsx";
const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/imp.log", a.map(String).join(" ") + "\n");
it("imp", async () => {
  const dir = __dirname + "/../fixtures/openpyxl/";
  for (const f of readdirSync(dir)) {
    const b = readFileSync(dir + f);
    try {
      const wb = await readXlsx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
      log(f, "OK sheets", wb.sheets.map((s) => `${s.name}[charts=${s.charts.length} tables=${s.tables.map(t=>t.name).join("/")} notes=${s.notes.size} merges=${s.merges.length} cf=${s.condFormats.map(c=>c.type).join("/")}]`).join(" "), "names", wb.engine.getNames().map(n=>n.name+n.formula).join(","), "warn", wb.loadWarnings.join("|"));
      for (const s of wb.sheets) for (const c of s.charts) log("   chart", c.type, c.range, c.title);
    } catch (e) {
      log(f, "FAIL", (e as Error).message);
    }
  }
}, 60000);
