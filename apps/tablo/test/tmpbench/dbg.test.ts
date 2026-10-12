import { it } from "vitest";
import { readFileSync, appendFileSync } from "node:fs";
import { normalizePackage } from "/home/user/officehere/apps/tablo/src/io/xlsxPackage";
import JSZip from "jszip";
const log = (...a: unknown[]) => appendFileSync("/tmp/claude-0/-home-user-officehere/8aa89398-b631-5201-a590-83027cc0aeca/scratchpad/tablo3/dbg.log", a.map(String).join(" ") + "\n");
it("d", async () => {
  const b = readFileSync("/home/user/officehere/apps/tablo/test/fixtures/openpyxl/aciklama.xlsx");
  const d = await normalizePackage(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
  const z = await JSZip.loadAsync(d);
  log(Object.keys(z.files).join(" "));
  log(await z.file("xl/worksheets/_rels/sheet1.xml.rels")!.async("string"));
  const ExcelJS: any = (await import("exceljs")).default;
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(d);
  const ws = x.worksheets[0];
  log(JSON.stringify(ws.getCell("A2").note), JSON.stringify(ws.getCell("A2").model));
});
