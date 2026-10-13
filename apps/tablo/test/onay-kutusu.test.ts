/** Hücre içi onay kutusu: ekleme, değiştirme, Delete davranışı, geri alma ve dosya gidiş-dönüşü. */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Controller } from "../src/controller";
import { deleteOnCheckboxes, insertCheckboxes, toggleCheckboxAt, toggleCheckboxes } from "../src/actions/checkbox";
import { computeView } from "../src/model/display";
import { rect } from "../src/model/rect";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { readOds } from "../src/io/ods/reader";
import { writeOds } from "../src/io/ods/writer";

function setup() {
  const ctl = new Controller();
  ctl.ops.setInputs(ctl.sheet, [{ row: 0, col: 0, input: "Görev" }, { row: 1, col: 1, input: true }, { row: 3, col: 1, input: { formula: "=COUNTIF(B1:B3,TRUE)" } }], "x");
  ctl.selectRange(rect(0, 1, 2, 1));
  insertCheckboxes(ctl);
  return ctl;
}
const v = (ctl: Controller, r: number, c: number) => ctl.wb.getValue(ctl.sheet, r, c);

describe("onay kutusu", () => {
  it("ekleme boş hücreleri YANLIŞ yapar, DOĞRU korunur; tek adımda geri alınır", () => {
    const ctl = setup();
    expect([0, 1, 2].map((r) => v(ctl, r, 1))).toEqual([false, true, false]);
    expect(ctl.wb.styleAt(ctl.sheet, 2, 1).checkbox).toBe(true);
    expect(v(ctl, 3, 1)).toBe(1);
    const view = computeView(ctl.wb, ctl.sheet, 1, 1, null);
    expect(view.checkbox).toBe(true);
    expect(view.text).toBe("");
    ctl.ops.undo();
    expect(v(ctl, 0, 1)).toBeNull();
    expect(ctl.wb.styleAt(ctl.sheet, 0, 1).checkbox).toBeUndefined();
  });

  it("tıklama ve Ara Çubuğu durumu değiştirir; formül güncellenir", () => {
    const ctl = setup();
    toggleCheckboxAt(ctl, 0, 1);
    expect(v(ctl, 0, 1)).toBe(true);
    expect(v(ctl, 3, 1)).toBe(2);
    // Etkin hücre işaretliyse seçimdeki tüm kutular kaldırılır
    ctl.selectRange(rect(0, 1, 2, 1));
    toggleCheckboxes(ctl);
    expect([0, 1, 2].map((r) => v(ctl, r, 1))).toEqual([false, false, false]);
    toggleCheckboxes(ctl);
    expect([0, 1, 2].map((r) => v(ctl, r, 1))).toEqual([true, true, true]);
    ctl.ops.undo();
    expect(v(ctl, 2, 1)).toBe(false);
  });

  it("Delete: önce işaretleri kaldırır, ikinci Delete onay kutusunu kaldırır", () => {
    const ctl = setup();
    ctl.selectRange(rect(0, 0, 2, 1));
    expect(deleteOnCheckboxes(ctl)).toBe(true);
    expect(v(ctl, 1, 1)).toBe(false);
    expect(v(ctl, 0, 0)).toBeNull();
    expect(ctl.wb.styleAt(ctl.sheet, 1, 1).checkbox).toBe(true);
    expect(deleteOnCheckboxes(ctl)).toBe(true);
    expect(v(ctl, 1, 1)).toBeNull();
    expect(ctl.wb.styleAt(ctl.sheet, 1, 1).checkbox).toBeUndefined();
    ctl.selectRange(rect(5, 5));
    expect(deleteOnCheckboxes(ctl)).toBe(false);
  });

  it("korumalı sayfada kilitli onay kutusu değişmez", () => {
    const ctl = setup();
    ctl.sheet.protected = true;
    let msg = "";
    ctl.alert = (m) => (msg = m);
    toggleCheckboxAt(ctl, 0, 1);
    expect(v(ctl, 0, 1)).toBe(false);
    expect(msg).toMatch(/korumalı/);
  });

  it("XLSX: özellik torbası ve xf uzantısıyla yazılır, geri okunur", async () => {
    const ctl = setup();
    ctl.ops.applyStyle(ctl.sheet, [rect(1, 1)], { bold: true });
    const bytes = await writeXlsx(ctl.wb);
    const zip = await JSZip.loadAsync(bytes);
    const fpb = await zip.file("xl/featurePropertyBag/featurePropertyBag.xml")?.async("string");
    expect(fpb).toContain('<bag type="Checkbox"/>');
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("application/vnd.ms-excel.featurepropertybag+xml");
    expect(await zip.file("xl/_rels/workbook.xml.rels")!.async("string")).toContain("relationships/FeaturePropertyBag");
    const styles = await zip.file("xl/styles.xml")!.async("string");
    expect(styles).toContain('<xfpb:xfComplement i="0"/>');
    const count = Number(/<cellXfs count="(\d+)"/.exec(styles)![1]);
    expect(count).toBe((styles.match(/<xf\b/g) ?? []).length - (/<cellStyleXfs\b[^>]*>([\s\S]*?)<\/cellStyleXfs>/.exec(styles)?.[1].match(/<xf\b/g) ?? []).length);
    const back = await readXlsx(bytes.buffer);
    const s = back.sheets[0];
    expect(back.styleAt(s, 0, 1).checkbox).toBe(true);
    expect(back.styleAt(s, 1, 1)).toMatchObject({ checkbox: true, bold: true });
    expect(back.getValue(s, 1, 1)).toBe(true);
    expect(back.styleAt(s, 0, 0).checkbox).toBeUndefined();
    expect(back.styleAt(s, 3, 1).checkbox).toBeUndefined();
  });

  it("onay kutusu yoksa özellik torbası yazılmaz; ODS gidiş-dönüşünde korunur", async () => {
    const plain = new Controller();
    plain.ops.setInputs(plain.sheet, [{ row: 0, col: 0, input: true }], "x");
    const zip = await JSZip.loadAsync(await writeXlsx(plain.wb));
    expect(zip.file("xl/featurePropertyBag/featurePropertyBag.xml")).toBeNull();
    const ctl = setup();
    const back = await readOds(await writeOds(ctl.wb));
    const s = back.sheets[0];
    expect(back.styleAt(s, 2, 1).checkbox).toBe(true);
    expect(back.getValue(s, 1, 1)).toBe(true);
  });
});
