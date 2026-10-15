/**
 * ODF geçiş katmanı (Tablo): açılan .ods'nin özgün paketi kaynak kaydına girer (`wb.kaynak`); .ods'ye kaydederken
 * settings.xml birleştirilir, manifest.rdf taşınır, Thumbnails/ düşer, sayfaların gömülü grafikleri (alt belgeleri ve
 * yedek görüntüleriyle) korunur. Veri aldığı sayfa değişince grafik yazılmaz (uyarılır). Worker yolu: kaynak baytları
 * IR'deki kimlikle yazan iş parçacığına gider.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { getSource } from "@officehere/core/opc";
import { parseManifest } from "@officehere/core/odf";
import { readOds } from "../src/io/ods/reader";
import { writeOds } from "../src/io/ods/writer";
import { ODF_OBJECTS_XLSX, openBook, writeBook } from "../src/io/formats";
import { serializeDraftSync, workbookFromDraft } from "../src/io/draft";
import { openBookIo, writeBookIo } from "../src/io/client";
import type { XlsxWriteReport } from "../src/io/xlsx";
import type { WorkbookModel } from "../src/model/workbook";

const fixture = (): ArrayBuffer => {
  const b = readFileSync(new URL("./fixtures/lo-stil.ods", import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};
const SLOW = { timeout: 60_000 };

async function pkg(bytes: Uint8Array | ArrayBuffer) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const manifest = parseManifest(await zip.file("META-INF/manifest.xml")!.async("string"));
  return { zip, names, manifest, content: await zip.file("content.xml")!.async("string"), text: (p: string) => zip.file(p)!.async("string") };
}

const objects = (wb: WorkbookModel) => wb.sheets.flatMap((s) => (s.passthrough?.odf ?? []).map((o) => ({ sheet: s.name, ...o })));
const CHARTS = ["Grafik1", "Cizgi", "Pasta"];

describe("ODS geçiş katmanı", () => {
  it("okurken: kaynak kaydı, sayfa grafikleri ve veri aralığı sayfaları; kayıp uyarısı yerine bilgi", SLOW, async () => {
    const report: { notice?: string } = {};
    const wb = await readOds(fixture(), report);
    expect(getSource(wb.kaynak)).toBeTruthy();
    const objs = objects(wb);
    expect(objs.map((o) => o.label)).toEqual(["Grafik", "Grafik", "Grafik"]);
    expect(objs.map((o) => o.sheet).sort()).toEqual(["Satışlar", "Satışlar", "Özet"]);
    expect(objs.find((o) => o.xml.includes("./Pasta"))!.deps).toEqual([["Özet", 0]]);
    expect(objs.every((o) => o.styles.some((s) => s.includes('style:family="graphic"')))).toBe(true);
    expect(wb.loadWarnings.join(" ")).not.toMatch(/grafik/);
    expect(report.notice).toMatch(/3 gömülü grafik/);
  });

  it(".ods'ye kaydederken alt belgeler, ayarlar ve RDF taşınır; Thumbnails düşer; ikinci tur kararlı", SLOW, async () => {
    const wb = await readOds(fixture());
    const out = await pkg(await writeOds(wb));
    for (const c of CHARTS) {
      expect(out.names).toContain(`${c}/content.xml`);
      expect(out.names).toContain(`ObjectReplacements/${c}`);
      expect(out.manifest.find((e) => e.path === `${c}/`)?.mediaType).toBe("application/vnd.oasis.opendocument.chart");
    }
    expect(out.names).toContain("manifest.rdf");
    expect(out.names.some((n) => n.startsWith("Thumbnails/"))).toBe(false);
    const listed = new Set(out.manifest.map((e) => e.path));
    for (const n of out.names) if (n !== "mimetype" && n !== "META-INF/manifest.xml") expect(listed).toContain(n);
    // Yazıcının görünüm ayarları + kaynağın belge ayarları
    const settings = await out.text("settings.xml");
    expect(settings).toContain('config:name="ActiveTable"');
    expect(settings).toContain('config:name="ooo:configuration-settings"');
    expect(settings.match(/config:name="ooo:view-settings"/g)).toHaveLength(1);
    // Çerçeve sayfanın table:shapes'inde, sütunlardan önce; stili yeniden adlandırılmış
    const ozet = out.content.slice(out.content.indexOf('table:name="Özet"'));
    expect(ozet.slice(0, ozet.indexOf("</table:table>"))).toMatch(/^[^>]*><table:shapes><draw:frame [^>]*draw:style-name="oh\d+-gr1"[^>]*>.*xlink:href="\.\/Pasta"/);
    // İkinci tur
    const again = await readOds(await writeOds(wb));
    expect(objects(again)).toHaveLength(3);
    const out2 = await pkg(await writeOds(again));
    expect(out2.names.sort()).toEqual(out.names.sort());
  });

  it("veri aldığı sayfanın yapısı ya da adı değişince grafik yazılmaz ve uyarılır; geri alınınca yine yazılır", SLOW, async () => {
    const wb = await readOds(fixture());
    const satis = wb.sheets.find((s) => s.name === "Satışlar")!;
    satis.structVersion++;
    const warnings: string[] = [];
    let out = await pkg(await writeOds(wb, { warnings }));
    expect(out.names).toContain("Pasta/content.xml");
    expect(out.names.some((n) => n.startsWith("Grafik1/") || n.startsWith("Cizgi/"))).toBe(false);
    expect(warnings).toContain("gömülü grafik (veri aldığı sayfa değiştirildi)");
    satis.structVersion--;
    expect((await pkg(await writeOds(wb))).names).toContain("Cizgi/content.xml");
    wb.sheets.find((s) => s.name === "Özet")!.name = "Özet 2";
    out = await pkg(await writeOds(wb));
    expect(out.names.some((n) => n.startsWith("Pasta/"))).toBe(false);
  });

  it("kaynak yoksa (taslaktan kurtarma) yazılmaz ve uyarılır; .xlsx'e kaydederken uyarılır", SLOW, async () => {
    const wb = await readOds(fixture());
    const draft = workbookFromDraft(serializeDraftSync(wb));
    expect(objects(draft)).toHaveLength(3);
    const warnings: string[] = [];
    const out = await pkg(await writeOds(draft, { warnings }));
    expect(out.names.some((n) => n.startsWith("Cizgi/"))).toBe(false);
    expect(warnings).toContain("gömülü grafik (özgün dosya bulunamadı)");
    const report: XlsxWriteReport = { warnings: [] };
    await writeBook(wb, wb.sheets[0], "xlsx", undefined, report);
    expect(report.warnings).toContain(ODF_OBJECTS_XLSX);
  });

  it("kaynak kaydı kapalıyken grafikler taşınmaz (kayıp uyarısı)", SLOW, async () => {
    const o = await openBook("dogrulama.ods", fixture(), { kaynak: false });
    expect(o.wb.kaynak).toBeUndefined();
    expect(objects(o.wb)).toHaveLength(0);
    expect(o.wb.loadWarnings).toContain("grafikler, resimler ve şekiller");
  });

  it("Worker yolu: kaynak ana iş parçacığında kaydedilir, yazarken baytları gider", SLOW, async () => {
    const o = await openBookIo("lo-stil.ods", fixture());
    expect(getSource(o.wb.kaynak)).toBeTruthy();
    expect(objects(o.wb)).toHaveLength(3);
    const report: XlsxWriteReport = { warnings: [] };
    const out = await pkg((await writeBookIo(o.wb, o.wb.sheets[0], "ods", undefined, report)) as Uint8Array);
    for (const c of CHARTS) expect(out.names).toContain(`${c}/content.xml`);
    expect(out.names).toContain("manifest.rdf");
    expect(report.warnings).toEqual([]);
  });
});
