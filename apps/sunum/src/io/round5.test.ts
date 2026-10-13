// @vitest-environment jsdom
/** 5. tur PPTX gidiş-dönüş: gösteri ayarları (p:showPr), yazdırma ayarları (p:prnPr), ortam uzantısı. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { newPresentation, newSlide } from "../model/presentation";
import type { Presentation } from "../model/types";
import { DEFAULT_SHOW_SETTINGS } from "../show/settings";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";
import { readPresProps } from "./presProps";

function deck(n = 3): Presentation {
  const p = newPresentation("varsayilan", 0);
  for (let i = 0; i < n; i++) p.slides.push(newSlide(p, "obj"));
  return p;
}

const part = async (bytes: Uint8Array, name: string) => (await JSZip.loadAsync(bytes)).file(name)!.async("string");

describe("Slayt Gösterisini Ayarla dosyaya yazılır (p:showPr)", () => {
  it("ayar yoksa presProps boş kalır, okununca tanımsızdır", async () => {
    const p = deck(1);
    const bytes = await writePptx(p);
    expect(await part(bytes, "ppt/presProps.xml")).not.toContain("showPr");
    const r = await readPptx(bytes);
    expect(r.showSettings).toBeUndefined();
    expect(r.printSettings).toBeUndefined();
  });

  it("bilgi noktası, döngü, animasyonsuz, el ile ilerleme, aralık ve renkler korunur", async () => {
    const p = deck(4);
    p.showSettings = { ...DEFAULT_SHOW_SETTINGS, type: "kiosk", loop: true, noAnimations: true, noNarration: true, advance: "manual", range: { kind: "range", from: 2, to: 4 }, penColor: "#00B050", laserColor: "#0070C0" };
    const bytes = await writePptx(p);
    const xml = await part(bytes, "ppt/presProps.xml");
    expect(xml).toContain('<p:showPr loop="1" showNarration="0" showAnimation="0" useTimings="0"><p:kiosk/><p:sldRg st="2" end="4"/>');
    expect(xml).toContain('<p:penClr><a:srgbClr val="00B050"/></p:penClr>');
    expect(xml).toContain("p14:laserClr");
    const r = await readPptx(bytes);
    expect(r.showSettings).toMatchObject({ type: "kiosk", loop: true, noAnimations: true, noNarration: true, advance: "manual", range: { kind: "range", from: 2, to: 4 }, penColor: "#00B050", laserColor: "#0070C0" });
  });

  it("özel gösteri seçimi custShow kimliğiyle eşlenir", async () => {
    const p = deck(3);
    p.customShows = [
      { id: "a", name: "Kısa", slideIds: [p.slides[0].id] },
      { id: "b", name: "Yönetim", slideIds: [p.slides[2].id, p.slides[1].id] },
    ];
    p.showSettings = { ...DEFAULT_SHOW_SETTINGS, type: "browsed", range: { kind: "custom", id: "b" } };
    const bytes = await writePptx(p);
    expect(await part(bytes, "ppt/presProps.xml")).toContain('<p:browse showScrollbar="0"/><p:custShow id="1"/>');
    const r = await readPptx(bytes);
    const shown = r.customShows!.find((c) => c.id === (r.showSettings!.range as { id: string }).id);
    expect(shown?.name).toBe("Yönetim");
    expect(r.showSettings!.type).toBe("browsed");
  });

  it("başka uygulamaların showPr'si şema varsayılanlarıyla okunur", () => {
    const xml = `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr showNarration="1"><p:present/><p:sldAll/><p:penClr><a:prstClr val="red"/></p:penClr></p:showPr></p:presentationPr>`;
    const r = readPresProps(xml, new Map());
    expect(r.showSettings).toMatchObject({ type: "speaker", loop: false, noNarration: false, noAnimations: false, advance: "timings", range: { kind: "all" }, penColor: "#FF0000" });
  });
});

describe("yazdırma ayarları (p:prnPr) ve not sayfası yönü", () => {
  it("dinleyici notları, gri tonlama, gizli slaytlar, çerçeve ve yatay yön gidiş-dönüşte korunur", async () => {
    const p = deck(2);
    p.printSettings = { what: "handouts6", color: "gray", hiddenSlides: true, frameSlides: true, orientation: "landscape" };
    const bytes = await writePptx(p);
    expect(await part(bytes, "ppt/presProps.xml")).toContain('<p:prnPr prnWhat="handouts6" clrMode="gray" hiddenSlides="1" frameSlides="1"/>');
    expect(await part(bytes, "ppt/presentation.xml")).toContain('<p:notesSz cx="9144000" cy="6858000"/>');
    const r = await readPptx(bytes);
    expect(r.printSettings).toMatchObject({ what: "handouts6", color: "gray", hiddenSlides: true, frameSlides: true, orientation: "landscape" });
  });
  it("prnPr gösteri ayarlarından önce yazılır (şema sırası)", async () => {
    const p = deck(1);
    p.printSettings = { what: "notes", color: "clr", hiddenSlides: false, frameSlides: false, orientation: "portrait" };
    p.showSettings = { ...DEFAULT_SHOW_SETTINGS };
    const xml = await part(await writePptx(p), "ppt/presProps.xml");
    expect(xml.indexOf("p:prnPr")).toBeGreaterThan(0);
    expect(xml.indexOf("p:prnPr")).toBeLessThan(xml.indexOf("p:showPr"));
  });
});

describe("not ve dinleyici notu üst/alt bilgisi (not asıl sayfası)", () => {
  it("üst bilgi, alt bilgi, sabit tarih ve kapalı sayfa numarası notsuz sunuda da korunur", async () => {
    const p = deck(1);
    p.notesHeaderFooter = { dateTime: true, dateAuto: false, dateFixed: "13 Ekim 2026", header: true, headerText: "Eğitim Notları", pageNumber: false, footer: true, footerText: "Şirket İçi" };
    const bytes = await writePptx(p);
    const nm = await part(bytes, "ppt/notesMasters/notesMaster1.xml");
    expect(nm).toContain('<p:hf sldNum="0"/>');
    expect(nm).toContain("Eğitim Notları");
    expect(nm.indexOf("<p:hf")).toBeGreaterThan(nm.indexOf("p:clrMap"));
    expect(nm.indexOf("<p:hf")).toBeLessThan(nm.indexOf("p:notesStyle"));
    expect(await part(bytes, "ppt/presentation.xml")).toContain("p:notesMasterIdLst");
    const r = await readPptx(bytes);
    expect(r.notesHeaderFooter).toEqual(p.notesHeaderFooter);
  });
  it("otomatik tarih alanla yazılır; ayar yoksa not asıl sayfasında p:hf olmaz", async () => {
    const p = deck(1);
    p.slides[0].notes = "Not";
    expect(await part(await writePptx(p), "ppt/notesMasters/notesMaster1.xml")).not.toContain("<p:hf");
    p.notesHeaderFooter = { dateTime: true, dateAuto: true, dateFixed: "", header: false, headerText: "", pageNumber: true, footer: false, footerText: "" };
    const bytes = await writePptx(p);
    expect(await part(bytes, "ppt/notesMasters/notesMaster1.xml")).toContain('type="datetime1"');
    const r = await readPptx(bytes);
    expect(r.notesHeaderFooter).toMatchObject({ dateTime: true, dateAuto: true, pageNumber: true, header: false, footer: false });
  });
});

describe("ODP: özel gösteriler ve gösteri ayarları (presentation:settings)", () => {
  it("özel gösteri, bilgi noktası/döngü ve el ile ilerleme gidiş-dönüşte korunur", async () => {
    const { writeOdp } = await import("./odp/writer");
    const { readOdp } = await import("./odp/reader");
    const p = deck(3);
    p.customShows = [{ id: "a", name: "Yönetim Özeti", slideIds: [p.slides[2].id, p.slides[0].id] }];
    p.showSettings = { ...DEFAULT_SHOW_SETTINGS, type: "kiosk", advance: "manual", noAnimations: true, range: { kind: "custom", id: "a" } };
    const bytes = await writeOdp(p);
    const content = await part(bytes, "content.xml");
    expect(content).toContain('presentation:show="Yönetim Özeti"');
    expect(content).toContain('presentation:endless="true"');
    const r = await readOdp(bytes);
    expect(r.customShows).toHaveLength(1);
    expect(r.customShows![0].name).toBe("Yönetim Özeti");
    expect(r.customShows![0].slideIds).toEqual([r.slides[2].id, r.slides[0].id]);
    expect(r.showSettings).toMatchObject({ type: "kiosk", loop: true, advance: "manual", noAnimations: true, range: { kind: "custom", id: r.customShows![0].id } });
  });
  it("ayar yoksa presentation:settings yazılmaz; başlangıç sayfası aralığa çevrilir", async () => {
    const { writeOdp } = await import("./odp/writer");
    const { readOdp } = await import("./odp/reader");
    const p = deck(4);
    expect(await part(await writeOdp(p), "content.xml")).not.toContain("presentation:settings");
    p.showSettings = { ...DEFAULT_SHOW_SETTINGS, type: "browsed", range: { kind: "range", from: 2, to: 4 } };
    const r = await readOdp(await writeOdp(p));
    // ODF yalnız başlangıç sayfasını saklar: aralık sona kadar sürer.
    expect(r.showSettings).toMatchObject({ type: "browsed", range: { kind: "range", from: 2, to: r.slides.length } });
  });
});
