// @vitest-environment jsdom
/** 3. tur PPTX gidiş-dönüş: özel gösteriler, grafik gruplama/eksen başlıkları/stil/türler, geçiş sesi, prova süreleri. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { makeChart } from "../model/chart";
import { textBody } from "../model/factory";
import { parseDocxOutline } from "../model/outlineImport";
import { newPresentation, newSlide } from "../model/presentation";
import type { ChartShape, Presentation } from "../model/types";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

function deck(n = 3): Presentation {
  const p = newPresentation("varsayilan", 0);
  for (let i = 0; i < n; i++) p.slides.push(newSlide(p, "obj"));
  p.slides.forEach((s, i) => {
    s.shapes = s.shapes.map((sh) => (sh.kind === "sp" && (sh.ph?.type === "title" || sh.ph?.type === "ctrTitle") ? { ...sh, text: textBody(`Slayt ${i + 1}`) } : sh));
  });
  return p;
}

const roundTrip = async (p: Presentation) => readPptx(await writePptx(p));

describe("özel slayt gösterileri", () => {
  it("p:custShowLst olarak yazılır ve sırasıyla okunur", async () => {
    const p = deck();
    p.customShows = [{ id: "x", name: "Yönetici Özeti", slideIds: [p.slides[3].id, p.slides[1].id] }];
    const bytes = await writePptx(p);
    const xml = await (await JSZip.loadAsync(bytes)).file("ppt/presentation.xml")!.async("string");
    expect(xml).toContain('<p:custShow name="Yönetici Özeti"');
    // Şema sırası: custShowLst, notesSz'den sonra ve defaultTextStyle'dan önce.
    expect(xml.indexOf("p:custShowLst")).toBeGreaterThan(xml.indexOf("p:notesSz"));
    expect(xml.indexOf("p:custShowLst")).toBeLessThan(xml.indexOf("p:defaultTextStyle"));
    const r = await readPptx(bytes);
    expect(r.customShows).toHaveLength(1);
    const ids = r.customShows![0].slideIds.map((id) => r.slides.findIndex((s) => s.id === id));
    expect(ids).toEqual([3, 1]);
    expect(r.customShows![0].name).toBe("Yönetici Özeti");
  });
});

describe("grafik ayrıntıları", () => {
  it("yığılmış gruplama, eksen başlıkları, kılavuzlar, stil, palet dışı alanlar korunur", async () => {
    const p = deck(1);
    const ch = makeChart("bar", 100, 100, 600, 400, "Grafik 1", { grouping: "percentStacked" });
    ch.chart = { ...ch.chart, axisTitles: { cat: "Bölge", val: "Pay" }, gridMajor: false, gridMinor: true, style: 3, fill: { rgb: "#F0F0F0" } };
    p.slides[1].shapes.push(ch);
    const r = await roundTrip(p);
    const c = r.slides[1].shapes.find((s): s is ChartShape => s.kind === "chart")!.chart;
    expect(c.type).toBe("bar");
    expect(c.grouping).toBe("percentStacked");
    expect(c.axisTitles).toEqual({ cat: "Bölge", val: "Pay" });
    expect(c.gridMajor).toBe(false);
    expect(c.gridMinor).toBe(true);
    expect(c.style).toBe(3);
    expect(c.fill).toEqual({ rgb: "#F0F0F0" });
  });

  it("dağılım ve radar grafikleri türünü korur", async () => {
    const p = deck(1);
    p.slides[1].shapes.push(makeChart("scatter", 0, 0, 400, 300, "Dağılım"), makeChart("radar", 400, 0, 400, 300, "Radar"), makeChart("combo", 0, 300, 400, 300, "Birleşik"));
    const r = await roundTrip(p);
    const types = r.slides[1].shapes.filter((s): s is ChartShape => s.kind === "chart").map((s) => s.chart.type);
    expect(types).toEqual(["scatter", "radar", "combo"]);
    const combo = r.slides[1].shapes.find((s): s is ChartShape => s.kind === "chart" && s.chart.type === "combo")!.chart;
    expect(combo.series).toHaveLength(3);
    const sc = r.slides[1].shapes.find((s): s is ChartShape => s.kind === "chart" && s.chart.type === "scatter")!.chart;
    expect(sc.series[0].values).toEqual([2.7, 3.2, 0.8]);
  });
});

describe("geçiş sesi ve zamanlamalar", () => {
  it("ses kimliği, 'önceki sesi durdur' ve otomatik ilerleme süresi korunur", async () => {
    const p = deck(2);
    p.slides[1].transition = { type: "fade", duration: 1, advanceOnClick: true, advanceAfter: 7, sound: "bell" };
    p.slides[2].transition = { type: "none", duration: 1, advanceOnClick: true, sound: "stop" };
    const bytes = await writePptx(p);
    const xml = await (await JSZip.loadAsync(bytes)).file("ppt/slides/slide3.xml")!.async("string");
    expect(xml).toContain("<p:endSnd/>");
    const r = await readPptx(bytes);
    expect(r.slides[1].transition?.sound).toBe("bell");
    expect(r.slides[1].transition?.advanceAfter).toBe(7);
    expect(r.slides[2].transition?.sound).toBe("stop");
  });
});

describe("belge ana hattı (.docx)", () => {
  it("Başlık 1 slayt başlığı, Başlık 2 madde olur", async () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const para = (style: string, t: string) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const zip = new JSZip();
    zip.file("word/document.xml", `<w:document xmlns:w="${W}"><w:body>${para("Heading1", "Gündem")}${para("Heading2", "Bütçe")}${para("Normal", "yok sayılır")}${para("Heading1", "Sonuç")}</w:body></w:document>`);
    zip.file("word/styles.xml", `<w:styles xmlns:w="${W}"><w:style w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>`);
    const out = await parseDocxOutline(await zip.generateAsync({ type: "uint8array" }));
    expect(out).toEqual([
      { title: "Gündem", lines: ["Bütçe"] },
      { title: "Sonuç", lines: [] },
    ]);
  });
});
