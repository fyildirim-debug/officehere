// @vitest-environment jsdom
/** 4. tur PPTX gidiş-dönüş: otomatik küçültme sonucu, eylemler, yakınlaştırma, Dönüştür geçişi, hareket yolları. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applySubtype, chartSubtypes, labelPositions, makeChart } from "../model/chart";
import { applyThemeToSlides } from "../model/designs";
import { textBody } from "../model/factory";
import { THEMES } from "../themes/themes";
import { newPresentation, newSlide } from "../model/presentation";
import type { ChartData, ChartShape, ChartType, Presentation, SpShape } from "../model/types";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

function deck(n = 2): Presentation {
  const p = newPresentation("varsayilan", 0);
  for (let i = 0; i < n; i++) p.slides.push(newSlide(p, "obj"));
  return p;
}

const slideXml = async (bytes: Uint8Array | ArrayBuffer | Blob, n: number) => (await JSZip.loadAsync(bytes)).file(`ppt/slides/slide${n}.xml`)!.async("string");
const bodyOf = (p: Presentation, i: number) => p.slides[i].shapes.find((s): s is SpShape => s.kind === "sp" && s.ph?.type === "obj")!;

describe("otomatik küçültme (normAutofit)", () => {
  it("düzenden miras küçültmede fontScale ve lnSpcReduction yazılır ve geri okunur", async () => {
    const p = deck(1);
    const sh = bodyOf(p, 1);
    sh.text = { ...textBody(Array.from({ length: 12 }, (_, i) => `Satır ${i + 1}`).join("\n")), fontScale: 0.625, lnSpcReduction: 0.2 };
    expect(sh.text.autofit).toBeUndefined();
    const bytes = await writePptx(p);
    const xml = await slideXml(bytes, 2);
    expect(xml).toContain('<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>');
    const r = await readPptx(bytes);
    const t = bodyOf(r, 1).text!;
    expect(t.fontScale).toBeCloseTo(0.625);
    expect(t.lnSpcReduction).toBeCloseTo(0.2);
    expect(t.autofit).toBe("shrink");
  });
  it("küçültülmemiş yer tutucuya autofit öğesi eklenmez", async () => {
    const p = deck(1);
    bodyOf(p, 1).text = textBody("Kısa");
    const xml = await slideXml(await writePptx(p), 2);
    expect(xml).not.toContain("normAutofit");
  });
});

describe("grafik alt türleri ve veri etiketi konumları (O4)", () => {
  it("Pasta, Halka, Dağılım, Radar ve Birleşik alt türleri dolu", () => {
    for (const t of ["pie", "doughnut", "scatter", "radar", "combo"] as ChartType[]) expect(chartSubtypes(t).length).toBeGreaterThan(1);
    expect(labelPositions("col").map((x) => x.label)).toEqual(["Orta", "İç Uç", "İç Taban", "Dış Uç", "Veri Belirtme Çizgisi"]);
    expect(labelPositions("col", true).some((x) => x.pos === "outEnd")).toBe(false);
    expect(applySubtype(makeChart("pie", 0, 0, 1, 1).chart, "pie", { type: "doughnut" }).type).toBe("doughnut");
  });

  const rt = async (patch: Partial<ChartData>, type: ChartType) => {
    const p = deck(1);
    const ch = makeChart(type, 0, 0, 600, 400, "G", patch);
    ch.chart = { ...ch.chart, ...patch };
    p.slides[1].shapes.push(ch);
    const bytes = await writePptx(p);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file(Object.keys(zip.files).find((f) => /charts\/chart\d+\.xml$/.test(f))!)!.async("string");
    const r = await readPptx(bytes);
    return { xml, c: r.slides[1].shapes.find((s): s is ChartShape => s.kind === "chart")!.chart };
  };

  it("sütunda Orta konumu dLblPos ctr olarak gidip gelir", async () => {
    const { xml, c } = await rt({ dataLabels: true, labelPos: "ctr" }, "col");
    expect(xml).toContain('<c:dLblPos val="ctr"/>');
    expect(c.labelPos).toBe("ctr");
  });
  it("pastada veri belirtme çizgisi balon geometrisiyle yazılır ve okunur", async () => {
    const { xml, c } = await rt({ dataLabels: true, labelPos: "callout" }, "pie");
    expect(xml).toContain('prst="wedgeRectCallout"');
    expect(c.labelPos).toBe("callout");
  });
  it("yığılmış sütunda geçersiz Dış Uç konumu yazılmaz", async () => {
    const { xml } = await rt({ dataLabels: true, labelPos: "outEnd", grouping: "stacked" }, "col");
    expect(xml).not.toContain('val="outEnd"');
  });
  it("dağılım, radar ve ikincil eksenli birleşik alt türleri korunur", async () => {
    expect((await rt({ scatterStyle: "smoothMarker" }, "scatter")).c.scatterStyle).toBe("smoothMarker");
    expect((await rt({ scatterStyle: "line" }, "scatter")).c.scatterStyle).toBe("line");
    expect((await rt({}, "scatter")).c.scatterStyle).toBeUndefined();
    expect((await rt({ radarStyle: "filled" }, "radar")).c.radarStyle).toBe("filled");
    expect((await rt({ radarStyle: "marker" }, "radar")).c.radarStyle).toBe("marker");
    const combo = await rt({ secondaryLine: true }, "combo");
    expect(combo.xml).toContain('<c:axPos val="r"/>');
    expect(combo.c.type).toBe("combo");
    expect(combo.c.secondaryLine).toBe(true);
  });
});

describe("Seçili Slaytlara Uygula (D6)", () => {
  it("tema yalnız seçili slayda uygulanır, ikincil asıl slayt olarak gidip gelir", async () => {
    const p = deck(2);
    const other = THEMES.find((t) => t.id !== p.theme.id)!;
    const target = p.slides[1].id;
    const next = applyThemeToSlides(p, other.id, [target]);
    const lay = (pr: Presentation, i: number) => pr.layouts.find((l) => l.id === pr.slides[i].layoutId)!;
    expect(lay(next, 1).theme?.id).toBe(other.id);
    expect(lay(next, 0).theme).toBeUndefined();
    expect(lay(next, 2).theme).toBeUndefined();
    // Ana temaya geri dönünce ikincil düzenler kaldırılır.
    const back = applyThemeToSlides(next, p.theme.id!, [target]);
    expect(back.layouts.some((l) => l.theme)).toBe(false);
    const bytes = await writePptx(next);
    const zip = await JSZip.loadAsync(bytes);
    const presXml = await zip.file("ppt/presentation.xml")!.async("string");
    expect(presXml.match(/<p:sldMasterId /g)).toHaveLength(2);
    expect(zip.file("ppt/slideMasters/slideMaster2.xml")).toBeTruthy();
    const r = await readPptx(bytes);
    expect(lay(r, 1).theme?.name).toBe(lay(next, 1).theme?.name);
    expect(lay(r, 0).theme).toBeUndefined();
    expect(r.slides).toHaveLength(3);
  });
});
