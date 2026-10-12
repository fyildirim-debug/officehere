// @vitest-environment jsdom
/** 2. tur özellikleri: köprüler, bölümler, grafikler, ortam, asıl slayt, tema efektleri (PPTX gidiş-dönüş). */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { makeChart } from "../model/chart";
import { makeSp, textBody } from "../model/factory";
import { addSectionAt } from "../model/sections";
import { newPresentation, newSlide } from "../model/presentation";
import type { ChartShape, PicShape, Presentation, SpShape } from "../model/types";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
const WAV = "data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA=";

function deck(): Presentation {
  const p = newPresentation("varsayilan", 0);
  for (let i = 0; i < 3; i++) p.slides.push(newSlide(p, "obj"));
  p.slides.forEach((s, i) => {
    s.shapes = s.shapes.map((sh) => (sh.kind === "sp" && (sh.ph?.type === "title" || sh.ph?.type === "ctrTitle") ? { ...sh, text: textBody(`Slayt ${i + 1}`) } : sh));
  });
  return p;
}

async function roundTrip(p: Presentation): Promise<Presentation> {
  return readPptx(await writePptx(p));
}

describe("köprüler", () => {
  it("metin, şekil ve slayt atlama köprüleri korunur", async () => {
    const p = deck();
    const target = p.slides[3].id;
    const box = makeSp("rect", 10, 10, 100, 50, { name: "Düğme", link: `#slide:${target}` });
    const tb = makeSp("rect", 200, 10, 300, 50, {
      txBox: true,
      text: { paragraphs: [{ runs: [{ text: "Site", props: { link: "https://ornek.com.tr/" } }, { text: " ve ", props: {} }, { text: "sonraki", props: { link: "#next" } }], level: 0, props: {} }] },
    });
    p.slides[1].shapes.push(box, tb);
    const r = await roundTrip(p);
    const s2 = r.slides[1];
    const rb = s2.shapes.find((s) => s.name === "Düğme")!;
    expect(rb.link).toBe(`#slide:${r.slides[3].id}`);
    const rt = s2.shapes.find((s) => s.kind === "sp" && s.txBox) as SpShape;
    const runs = rt.text!.paragraphs[0].runs;
    expect(runs[0].props.link).toBe("https://ornek.com.tr/");
    expect(runs.find((x) => x.text === "sonraki")?.props.link).toBe("#next");
  });

  it("slayt atlaması PPTX'te ppaction ile yazılır", async () => {
    const p = deck();
    p.slides[0].shapes.push(makeSp("rect", 0, 0, 10, 10, { link: `#slide:${p.slides[2].id}` }));
    const zip = await JSZip.loadAsync(await writePptx(p));
    const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(xml).toContain('action="ppaction://hlinksldjump"');
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    expect(rels).toContain('Target="slide3.xml"');
  });
});

describe("bölümler", () => {
  it("p14:sectionLst yazılır ve okunur", async () => {
    let p = deck();
    const r1 = addSectionAt(p.slides, 2, "Sonuçlar")!;
    p = { ...p, slides: r1.slides };
    const zip = await JSZip.loadAsync(await writePptx(p));
    const pres = await zip.file("ppt/presentation.xml")!.async("string");
    expect(pres).toContain("p14:sectionLst");
    expect(pres).toContain('name="Varsayılan Bölüm"');
    const r = await roundTrip(p);
    expect(r.slides[0].section?.name).toBe("Varsayılan Bölüm");
    expect(r.slides[2].section?.name).toBe("Sonuçlar");
    expect(r.slides[1].section).toBeUndefined();
  });
});

describe("grafikler", () => {
  it("gerçek grafik parçası ve gömülü çalışma kitabı yazılır, geri okunur", async () => {
    const p = deck();
    const ch = makeChart("col", 100, 150, 600, 400, "Grafik 1");
    ch.chart.series[0].values[1] = 7.25;
    ch.chart.dataLabels = true;
    const pie = makeChart("doughnut", 700, 150, 400, 400, "Grafik 2");
    p.slides[1].shapes.push(ch, pie);
    const bytes = await writePptx(p);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("ppt/charts/chart1.xml")).toBeTruthy();
    expect(zip.file("ppt/embeddings/Calisma_Sayfasi1.xlsx")).toBeTruthy();
    const ct = await zip.file("[Content_Types].xml")!.async("string");
    expect(ct).toContain("drawingml.chart+xml");
    const cx = await zip.file("ppt/charts/chart1.xml")!.async("string");
    expect(cx).toContain('<c:barDir val="col"/>');
    expect(cx).toContain("<c:v>7.25</c:v>");
    // Gömülü çalışma kitabı geçerli bir ZIP'tir ve veriyi içerir.
    const wb = await JSZip.loadAsync(await zip.file("ppt/embeddings/Calisma_Sayfasi1.xlsx")!.async("uint8array"));
    expect(await wb.file("xl/worksheets/sheet1.xml")!.async("string")).toContain("Kategori 1");
    const r = await readPptx(bytes);
    const rc = r.slides[1].shapes.filter((s): s is ChartShape => s.kind === "chart");
    expect(rc).toHaveLength(2);
    expect(rc[0].chart.type).toBe("col");
    expect(rc[0].chart.series).toHaveLength(3);
    expect(rc[0].chart.series[0].values[1]).toBe(7.25);
    expect(rc[0].chart.categories).toEqual(["Kategori 1", "Kategori 2", "Kategori 3", "Kategori 4"]);
    expect(rc[0].chart.dataLabels).toBe(true);
    expect(rc[1].chart.type).toBe("doughnut");
    expect(rc[1].x).toBeCloseTo(700, 0);
  });

  it("tüm grafik türleri gidiş-dönüşte türünü korur", async () => {
    const p = deck();
    const types = ["col", "bar", "line", "pie", "doughnut", "area"] as const;
    types.forEach((t, i) => p.slides[i % 4].shapes.push(makeChart(t, 10 + i * 10, 10, 300, 200)));
    const r = await roundTrip(p);
    const got = r.slides.flatMap((s) => s.shapes).filter((s): s is ChartShape => s.kind === "chart").map((c) => c.chart.type).sort();
    expect(got).toEqual([...types].sort());
  });
});

describe("video ve ses", () => {
  it("ortam dosyası, kapak ve seçenekler korunur", async () => {
    const p = deck();
    const audio: PicShape = { kind: "pic", id: "a1", name: "Ses 1", x: 10, y: 10, w: 48, h: 48, rot: 0, src: PNG, line: { fill: { type: "none" }, width: 1, dash: "solid" }, media: { kind: "audio", src: WAV, mime: "audio/wav", loop: true } };
    p.slides[1].shapes.push(audio);
    const bytes = await writePptx(p);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("ppt/media/media1.wav")).toBeTruthy();
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("<a:audioFile");
    expect(xml).toContain("p14:media");
    const r = await readPptx(bytes);
    const ra = r.slides[1].shapes.find((s): s is PicShape => s.kind === "pic")!;
    expect(ra.media?.kind).toBe("audio");
    expect(ra.media?.src.startsWith("data:audio/wav")).toBe(true);
    expect(ra.media?.loop).toBe(true);
    expect(ra.src.startsWith("data:image/png")).toBe(true);
  });
});

describe("tema efektleri", () => {
  it("effectRef tema efekt stilinden gölge alır, boş effectLst kaldırır", async () => {
    const p = deck();
    const zip = await JSZip.loadAsync(await writePptx(p));
    const slide = await zip.file("ppt/slides/slide2.xml")!.async("string");
    const style = '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
    const sp = (id: number, fx: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Efekt ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fx}</p:spPr>${style}</p:sp>`;
    zip.file("ppt/slides/slide2.xml", slide.replace("</p:spTree>", `${sp(90, "")}${sp(91, "<a:effectLst/>")}</p:spTree>`));
    const r = await readPptx(await zip.generateAsync({ type: "uint8array" }));
    const a = r.slides[1].shapes.find((s) => s.name === "Efekt 90") as SpShape;
    const b = r.slides[1].shapes.find((s) => s.name === "Efekt 91") as SpShape;
    expect(a.effects?.shadow).toBeTruthy();
    expect(a.effects!.shadow!.dist).toBeGreaterThan(0);
    expect(b.effects).toBeUndefined();
  });
});
