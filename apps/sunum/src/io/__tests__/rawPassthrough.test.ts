// @vitest-environment jsdom
/**
 * Geçiş katmanı (K§1.2.4 Sunum): SmartArt/OLE/mürekkep gibi modelin tanımadığı öğeler yedek
 * görünümleriyle okunur, değişmedikçe özgün XML'leri ve parçalarıyla yazılır; slayt/sunu düzeyi
 * tanınmayan öğeler ve paket parçaları (makro, imza politikası, notlar) korunur.
 * Fikstürler: test-derlem/poi (Apache POI test verisi, Apache-2.0).
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { explicitRelIds, OpcPackage, releaseSource } from "@officehere/core/opc";
import { makeSp } from "../../model/factory";
import type { GroupShape, PicShape, Presentation, Shape, SpShape } from "../../model/types";
import { readPresentationFile, writePresentationFormat } from "../formatsImpl";
import { RAW_WARN, rawSig, rawSigMatches } from "../rawPassthrough";

vi.setConfig({ testTimeout: 60_000 });

const POI = `${process.cwd()}/../../test-derlem/poi`;
const poi = (n: string) => new Uint8Array(readFileSync(`${POI}/${n}`));
const EMU = 9525;

async function open(n: string, bytes: Uint8Array = poi(n)) {
  return readPresentationFile(bytes, n);
}
async function save(pres: Presentation, format: "pptx" | "pptm" | "odp" = "pptx") {
  const warnings: string[] = [];
  const bytes = await writePresentationFormat(pres, format, { warnings });
  return { bytes, warnings, zip: await JSZip.loadAsync(bytes) };
}
const text = async (z: JSZip, p: string) => (await z.file(p)?.async("string")) ?? "";
const allShapes = (l: Shape[]): Shape[] => l.flatMap((s) => (s.kind === "group" ? [s, ...allShapes(s.children)] : [s]));
const rawShapes = (p: Presentation) => p.slides.flatMap((s) => allShapes(s.shapes)).filter((s) => s.raw);
const parts = (z: JSZip, re: RegExp) => Object.keys(z.files).filter((n) => re.test(n));

/** OPC tutarlılığı: her iç ilişkinin hedefi var, XML'deki her r:* kimliği parçanın .rels'inde var. */
async function opcErrors(bytes: Uint8Array): Promise<string[]> {
  const pkg = await OpcPackage.open(bytes);
  const out: string[] = [];
  for (const name of pkg.names) {
    if (name.endsWith(".rels") || name === "[Content_Types].xml" || !/\.xml$/i.test(name)) continue;
    const rels = await pkg.rels(name);
    for (const r of rels) if (!r.external && (!r.part || !pkg.has(r.part))) out.push(`${name}: hedefsiz ${r.id} → ${r.target}`);
    const ids = new Set(rels.map((r) => r.id));
    // dsp:dataModelExt relId slaydın ilişkisini anar (parçanın kendi .rels'ini değil): ayrıca denetlenir.
    const xml = ((await pkg.text(name)) ?? "").replace(/\srelId="[^"]*"/g, "");
    for (const id of explicitRelIds(xml)) if (!ids.has(id)) out.push(`${name}: kırık ${id}`);
    if (!pkg.contentType(name)) out.push(`${name}: içerik türü yok`);
  }
  return out;
}

const smartArt = (p: Presentation) => rawShapes(p).find((s): s is GroupShape => s.raw?.kind === "smartart" && s.kind === "group")!;

function moveShape<T extends Shape>(s: T, dx: number, dy: number): T {
  const m = { ...s, x: s.x + dx, y: s.y + dy } as T;
  if (m.kind === "group") (m as GroupShape).children = (m as GroupShape).children.map((c) => moveShape(c, dx, dy));
  return m;
}

describe("SmartArt", () => {
  it("dsp çiziminden grup olarak okunur; özgün XML ve 5 parçasıyla yazılır, relId yamalanır", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const g = smartArt(pres);
    expect(g.children.length).toBeGreaterThan(2);
    expect(g.raw!.refs.map((r) => r.type.split("/").pop()).sort()).toEqual(["diagramColors", "diagramData", "diagramLayout", "diagramQuickStyle"]);
    expect(g.raw!.extra).toHaveLength(1);
    const { bytes, warnings, zip } = await save(pres);
    expect(warnings).toEqual([]);
    expect(await opcErrors(bytes)).toEqual([]);
    const slide = await text(zip, "ppt/slides/slide1.xml");
    expect(slide).toContain("dgm:relIds");
    const rels = await text(zip, "ppt/slides/_rels/slide1.xml.rels");
    for (const t of ["diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors", "diagramDrawing"]) expect(rels).toContain(`/${t}"`);
    // dsp:dataModelExt relId yeni slaydın çizim ilişkisini anar.
    const dataPath = parts(zip, /^ppt\/diagrams\/.*data1\.xml$/)[0];
    const relId = /dataModelExt[^>]*relId="([^"]+)"/.exec(await text(zip, dataPath))![1];
    expect(rels).toMatch(new RegExp(`Id="${relId}"[^>]*diagramDrawing"`));
    // İkinci tur: yine SmartArt, parçalar korunur.
    const again = await open("a.pptx", bytes);
    expect(smartArt(again.pres)).toBeTruthy();
    const second = await save(again.pres);
    expect(parts(second.zip, /^ppt\/diagrams\/[^/]+\.xml$/)).toHaveLength(5);
  });

  it("taşınınca özgün XML kalır, konum güncellenir", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const g = smartArt(pres);
    const moved = moveShape(g, 50, 20);
    pres.slides[0].shapes = pres.slides[0].shapes.map((s) => (s === g ? moved : s));
    const { warnings, zip } = await save(pres);
    expect(warnings).toEqual([]);
    const slide = await text(zip, "ppt/slides/slide1.xml");
    expect(slide).toContain(`<a:off x="${Math.round(moved.x * EMU)}" y="${Math.round(moved.y * EMU)}"/>`);
    expect(parts(zip, /^ppt\/diagrams\/[^/]+\.xml$/)).toHaveLength(5);
  });

  it("alt şeklin metni düzenlenince düz grup + uyarı yazılır", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const g = smartArt(pres);
    const i = g.children.findIndex((c) => c.kind === "sp" && c.text?.paragraphs.some((p) => p.runs.length));
    const child = g.children[i] as SpShape;
    const edited: GroupShape = { ...g, children: g.children.map((c, k) => (k === i ? { ...child, text: { ...child.text!, paragraphs: [{ runs: [{ text: "Değişti", props: {} }], level: 0, props: {} }] } } : c)) };
    expect(rawSigMatches(edited)).toBe(false);
    pres.slides[0].shapes = pres.slides[0].shapes.map((s) => (s === g ? edited : s));
    const { bytes, warnings, zip } = await save(pres);
    expect(warnings).toContain(RAW_WARN.smartart);
    expect(parts(zip, /^ppt\/diagrams\//)).toEqual([]);
    const slide = await text(zip, "ppt/slides/slide1.xml");
    expect(slide).toContain("<p:grpSp>");
    expect(slide).toContain("Değişti");
    expect(slide).not.toContain("dgm:relIds");
    expect(await opcErrors(bytes)).toEqual([]);
  });

  it("çoğaltılan SmartArt kendi parçalarını ve yeni model kimliklerini alır", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const g = smartArt(pres);
    const copy = moveShape({ ...structuredClone(g), id: "kopya" }, 30, 30);
    pres.slides[0].shapes.push(copy);
    const { bytes, warnings, zip } = await save(pres);
    expect(warnings).toEqual([]);
    expect(await opcErrors(bytes)).toEqual([]);
    const datas = parts(zip, /^ppt\/diagrams\/.*data1\.xml$/);
    expect(datas).toHaveLength(2);
    const ids = await Promise.all(datas.map(async (p) => new Set([...(await text(zip, p)).matchAll(/modelId="(\{[^"]+\})"/g)].map((m) => m[1]))));
    expect(ids[0].size).toBeGreaterThan(0);
    expect([...ids[0]].filter((x) => ids[1].has(x))).toEqual([]);
    // Çizim parçası da aynı yeni kimlikleri kullanır.
    const drawings = parts(zip, /^ppt\/diagrams\/.*drawing1\.xml$/);
    for (const d of drawings) {
      const own = [...(await text(zip, d)).matchAll(/modelId="(\{[^"]+\})"/g)].map((m) => m[1]);
      expect(own.length).toBeGreaterThan(0);
      expect(ids.some((set) => own.every((x) => set.has(x)))).toBe(true);
    }
    const rels = await text(zip, "ppt/slides/_rels/slide1.xml.rels");
    for (const p of datas) {
      const relId = /dataModelExt[^>]*relId="([^"]+)"/.exec(await text(zip, p))![1];
      expect(rels).toMatch(new RegExp(`Id="${relId}"[^>]*diagramDrawing"`));
    }
  });

  it("silinen slaydın SmartArt parçaları taşınmaz", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const blank = { ...structuredClone(pres.slides[0]), id: "bos", shapes: [], raw: undefined };
    pres.slides = [blank];
    const { bytes, zip } = await save(pres);
    expect(parts(zip, /^ppt\/diagrams\//)).toEqual([]);
    expect(await opcErrors(bytes)).toEqual([]);
  });

  it("ODP'ye yedek görünüm (şekil grubu) yazılır ve uyarılır", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const words = smartArt(pres)
      .children.flatMap((c) => (c.kind === "sp" ? (c.text?.paragraphs ?? []) : []))
      .flatMap((p) => p.runs.map((r) => r.text))
      .filter(Boolean);
    expect(words.length).toBeGreaterThan(0);
    const { warnings, zip } = await save(pres, "odp");
    expect(warnings).toContain(RAW_WARN.odp);
    const content = await text(zip, "content.xml");
    expect(content).toContain("<draw:g");
    for (const w of words) expect(content).toContain(w);
  });

  it("kaynak paket bellekte yoksa yedek görünüm + uyarı", async () => {
    const { pres } = await open("smartart-simple.pptx");
    releaseSource(pres.kaynak);
    const { bytes, warnings, zip } = await save(pres);
    expect(warnings).toContain(RAW_WARN.source);
    expect(parts(zip, /^ppt\/diagrams\//)).toEqual([]);
    expect(await text(zip, "ppt/slides/slide1.xml")).toContain("<p:grpSp>");
    expect(await opcErrors(bytes)).toEqual([]);
  });
});

describe("OLE, makro, imza", () => {
  it("OLE nesneleri yedek resimle görünür ve .pptm'de makroyla korunur", async () => {
    const { pres, format } = await open("PPTWithAttachments.pptm");
    expect(format).toBe("pptm");
    const ole = rawShapes(pres).filter((s) => s.raw?.kind === "ole");
    expect(ole.length).toBeGreaterThan(0);
    for (const s of ole) expect(s.kind === "pic" && !!(s as PicShape).src).toBe(true);
    const { bytes, warnings, zip } = await save(pres, "pptm");
    expect(warnings).toEqual([]);
    expect(await opcErrors(bytes)).toEqual([]);
    expect(await text(zip, "[Content_Types].xml")).toContain("application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml");
    expect(parts(zip, /vbaProject\.bin$/)).toHaveLength(1);
    expect(parts(zip, /^ppt\/embeddings\//).length).toBeGreaterThanOrEqual(4);
    // Eski tür OLE önizlemeleri slaydın VML çiziminde: çizim de taşınır.
    expect(parts(zip, /\.vml$/).length).toBeGreaterThan(0);
  });

  it(".pptm .pptx olarak kaydedilince makro düşer ve uyarılır", async () => {
    const { pres } = await open("SimpleMacro.pptm");
    const { warnings, zip } = await save(pres);
    expect(parts(zip, /vbaProject\.bin$/)).toEqual([]);
    expect(warnings.some((w) => w.includes("Makrolar"))).toBe(true);
    expect(await text(zip, "[Content_Types].xml")).toContain("presentationml.presentation.main+xml");
  });

  it("imzalı sunu: açılışta bildirim, kaydedince imza düşer ve uyarılır", async () => {
    const { pres, notices, warnings: loadWarnings } = await open("hello-world-signed.pptx");
    expect(loadWarnings).toEqual([]);
    expect(notices.some((n) => n.includes("imza"))).toBe(true);
    const { bytes, warnings, zip } = await save(pres);
    expect(parts(zip, /^_xmlsignatures\//)).toEqual([]);
    expect(warnings.some((w) => w.includes("imza"))).toBe(true);
    expect(await opcErrors(bytes)).toEqual([]);
  });
});

describe("sunu ve slayt düzeyi", () => {
  it("presentation.xml yaması: dinleyici notu asıl sayfası ve kılavuz uzantısı korunur", async () => {
    const att = await save((await open("PPTWithAttachments.pptm")).pres, "pptm");
    const p1 = await text(att.zip, "ppt/presentation.xml");
    expect(p1).toContain("handoutMasterIdLst");
    expect(parts(att.zip, /^ppt\/handoutMasters\/[^/]+\.xml$/)).toHaveLength(1);
    expect(await opcErrors(att.bytes)).toEqual([]);
    const bar = await save((await open("bar-chart.pptx")).pres);
    const p2 = await text(bar.zip, "ppt/presentation.xml");
    expect(p2).toContain("{EFAFB233-063F-42B5-8137-9DF3F51BA10A}");
    expect(p2).toContain("sldGuideLst");
    // Şema sırası: extLst en sonda, defaultTextStyle'dan sonra.
    expect(p2.indexOf("defaultTextStyle")).toBeLessThan(p2.indexOf("sldGuideLst"));
  });

  it("slayt extLst (p14:creationId) taşınır; çoğaltılan slaytta yinelenmez", async () => {
    const { pres } = await open("smartart-simple.pptx");
    const src = new TextDecoder().decode((await (await JSZip.loadAsync(poi("smartart-simple.pptx"))).file("ppt/slides/slide1.xml")!.async("uint8array")));
    const val = /creationId[^>]*val="(\d+)"/.exec(src)![1];
    pres.slides.push({ ...structuredClone(pres.slides[0]), id: "kopya" });
    const { bytes, zip } = await save(pres);
    expect(await text(zip, "ppt/slides/slide1.xml")).toContain(`val="${val}"`);
    const second = /creationId[^>]*val="(\d+)"/.exec(await text(zip, "ppt/slides/slide2.xml"))![1];
    expect(second).not.toBe(val);
    expect(await opcErrors(bytes)).toEqual([]);
  });

  it("kaynakta not sayfası olan slayt notsuz da not sayfasıyla yazılır", async () => {
    const { pres } = await open("layouts.pptx");
    const src = await JSZip.loadAsync(poi("layouts.pptx"));
    const { zip } = await save(pres);
    expect(parts(zip, /^ppt\/notesSlides\/[^/]+\.xml$/).length).toBe(parts(src, /^ppt\/notesSlides\/[^/]+\.xml$/).length);
    expect(parts(zip, /^ppt\/notesMasters\/[^/]+\.xml$/)).toHaveLength(1);
  });

  it("tanınmayan öğe (mürekkep) yer tutucu resimle görünür, özgün XML ve parçasıyla yazılır", async () => {
    const z = await JSZip.loadAsync(poi("shapes.pptx"));
    const slide = (await z.file("ppt/slides/slide1.xml")!.async("string")).replace(
      "</p:spTree>",
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:contentPart p14:bwMode="auto" r:id="rId99"><p14:nvContentPartPr><p14:cNvPr id="777" name="Mürekkep 9"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr><p14:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></p14:xfrm></p:contentPart></mc:Choice><mc:Fallback/></mc:AlternateContent></p:spTree>',
    );
    z.file("ppt/slides/slide1.xml", slide);
    const rels = (await z.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).replace("</Relationships>", '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../ink/ink1.xml"/></Relationships>');
    z.file("ppt/slides/_rels/slide1.xml.rels", rels);
    z.file("ppt/ink/ink1.xml", '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>');
    const ct = (await z.file("[Content_Types].xml")!.async("string")).replace("</Types>", '<Override PartName="/ppt/ink/ink1.xml" ContentType="application/inkml+xml"/></Types>');
    z.file("[Content_Types].xml", ct);
    const { pres } = await open("murekkep.pptx", await z.generateAsync({ type: "uint8array" }));
    const ink = rawShapes(pres).find((s) => s.raw?.kind === "murekkep") as PicShape;
    expect(ink.kind).toBe("pic");
    expect(ink.src).toMatch(/^data:image\/svg\+xml;base64,/);
    expect(ink.w).toBeCloseTo(200);
    const { bytes, zip } = await save(pres);
    expect(await opcErrors(bytes)).toEqual([]);
    const out = await text(zip, "ppt/slides/slide1.xml");
    expect(out).toContain("p:contentPart");
    expect(out).toMatch(/xmlns:p14="http:\/\/schemas.microsoft.com\/office\/powerpoint\/2010\/main"[^>]*>|Requires="p14"/);
    expect(parts(zip, /^ppt\/ink\/.*ink1\.xml$/)).toHaveLength(1);
  });
});

// 1×1 PNG
const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));

/** Parçaya resim ilişkisi ekler (yoksa .rels kurar); ilişki kimliği "rIdP1". */
async function addImageRel(z: JSZip, part: string, media: string): Promise<void> {
  const i = part.lastIndexOf("/");
  const relsPath = `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
  const rels = (await z.file(relsPath)?.async("string")) ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  z.file(relsPath, rels.replace("</Relationships>", `<Relationship Id="rIdP1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${media}"/></Relationships>`));
  z.file(`ppt/media/${media}`, PNG);
  const ct = (await z.file("[Content_Types].xml")!.async("string")).replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>');
  if (!/Extension="png"/.test(await z.file("[Content_Types].xml")!.async("string"))) z.file("[Content_Types].xml", ct);
}

describe("tema, asıl slayt, grafik ve metin geçişi", () => {
  it("içe aktarılan tema değişmediyse özgün tema (resimli dolgu stiliyle) yazılır; değişince yeniden üretilir", async () => {
    const z = await JSZip.loadAsync(poi("smartart-simple.pptx"));
    const theme = (await z.file("ppt/theme/theme1.xml")!.async("string")).replace(/<a:bgFillStyleLst>/, '<a:bgFillStyleLst><a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdP1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>');
    z.file("ppt/theme/theme1.xml", theme);
    await addImageRel(z, "ppt/theme/theme1.xml", "tema-doku.png");
    const { pres } = await open("tema.pptx", await z.generateAsync({ type: "uint8array" }));
    const kept = await save(pres);
    expect(await text(kept.zip, "ppt/theme/theme1.xml")).toMatch(/<a:blipFill><a:blip [^>]*r:embed="rId\d+"/);
    expect(parts(kept.zip, /^ppt\/media\/.*tema-doku\.png$/)).toHaveLength(1);
    expect(await opcErrors(kept.bytes)).toEqual([]);
    const changed = await save({ ...pres, theme: { ...pres.theme, colors: { ...pres.theme.colors, accent1: "#123456" } } });
    const t = await text(changed.zip, "ppt/theme/theme1.xml");
    expect(t).toContain('val="123456"');
    expect(t).not.toContain("blipFill");
  });

  it("ikincil asıl slaytlar ayrı asıl slayt olarak kalır", async () => {
    const src = await JSZip.loadAsync(new Uint8Array(readFileSync(`${process.cwd()}/src/io/__fixtures__/lo-proje.pptx`)));
    // Düzeni olan asıl slaytlar (düzensiz asıl slayt modelde temsil edilmez).
    const rels = await Promise.all(parts(src, /^ppt\/slideMasters\/_rels\/[^/]+\.rels$/).map((p) => text(src, p)));
    const masters = rels.filter((r) => r.includes("/slideLayout")).length;
    expect(masters).toBeGreaterThan(1);
    const { pres } = await open("lo-proje.pptx", new Uint8Array(readFileSync(`${process.cwd()}/src/io/__fixtures__/lo-proje.pptx`)));
    const { bytes, zip } = await save(pres);
    expect(parts(zip, /^ppt\/slideMasters\/[^/]+\.xml$/).length).toBe(masters);
    expect(await opcErrors(bytes)).toEqual([]);
    // İkinci tur aynı yapıda.
    const again = await save((await open("b.pptx", bytes)).pres);
    expect(parts(again.zip, /^ppt\/slideMasters\/[^/]+\.xml$/).length).toBe(masters);
  });

  it("yeniden üretilen grafiğe kaynaktaki biçem ve renk parçaları bağlanır", async () => {
    const z = await JSZip.loadAsync(poi("bar-chart.pptx"));
    const MS = "http://schemas.microsoft.com/office/2011/relationships/";
    const rels = (await z.file("ppt/charts/_rels/chart1.xml.rels")!.async("string")).replace(
      "</Relationships>",
      `<Relationship Id="rId8" Type="${MS}chartStyle" Target="style1.xml"/><Relationship Id="rId9" Type="${MS}chartColorStyle" Target="colors1.xml"/></Relationships>`,
    );
    z.file("ppt/charts/_rels/chart1.xml.rels", rels);
    z.file("ppt/charts/style1.xml", '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>');
    z.file("ppt/charts/colors1.xml", '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>');
    const ct = (await z.file("[Content_Types].xml")!.async("string")).replace(
      "</Types>",
      '<Override PartName="/ppt/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/ppt/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>',
    );
    z.file("[Content_Types].xml", ct);
    const { pres } = await open("grafik.pptx", await z.generateAsync({ type: "uint8array" }));
    const shapes = pres.slides[0].shapes.map((s) => (s.kind === "chart" ? { ...s, chart: { ...s.chart, title: "Yeni başlık" } } : s));
    const { bytes, zip } = await save({ ...pres, slides: [{ ...pres.slides[0], shapes }] });
    const out = parts(zip, /^ppt\/charts\/chart\d+\.xml$/);
    expect(out).toHaveLength(1);
    expect(await text(zip, out[0])).toContain("Yeni başlık");
    const outRels = await text(zip, out[0].replace("charts/", "charts/_rels/") + ".rels");
    expect(outRels).toContain(`${MS}chartStyle`);
    expect(outRels).toContain(`${MS}chartColorStyle`);
    expect(await opcErrors(bytes)).toEqual([]);
  });

  it("resimli madde işareti ve resim dolgulu metin korunur", async () => {
    const z = await JSZip.loadAsync(poi("shapes.pptx"));
    const slide = await z.file("ppt/slides/slide1.xml")!.async("string");
    const i = slide.indexOf("<a:r>");
    expect(i).toBeGreaterThan(0);
    const patched = `${slide.slice(0, i)}<a:pPr><a:buBlip><a:blip r:embed="rIdP1"/></a:buBlip></a:pPr><a:r><a:rPr lang="tr-TR"><a:blipFill><a:blip r:embed="rIdP1"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill></a:rPr>${slide.slice(i + "<a:r>".length).replace(/^<a:rPr[^>]*\/>|^<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/, "")}`;
    z.file("ppt/slides/slide1.xml", patched);
    await addImageRel(z, "ppt/slides/slide1.xml", "madde.png");
    const { pres } = await open("madde.pptx", await z.generateAsync({ type: "uint8array" }));
    const paras = pres.slides[0].shapes.flatMap((s) => (s.kind === "sp" ? (s.text?.paragraphs ?? []) : []));
    expect(paras.some((p) => p.props.bullet?.type === "char" && !!p.props.bullet.pic)).toBe(true);
    expect(paras.some((p) => p.runs.some((r) => r.props.picFill?.mode === "tile"))).toBe(true);
    const { bytes, zip } = await save(pres);
    const out = await text(zip, "ppt/slides/slide1.xml");
    expect(out).toContain("<a:buBlip><a:blip r:embed=");
    expect(out).toMatch(/<a:rPr[^>]*>(<a:ln[^>]*>[\s\S]*?<\/a:ln>)?<a:blipFill/);
    expect(await opcErrors(bytes)).toEqual([]);
  });
});

describe("görünüm imzası", () => {
  const group = (): GroupShape => {
    const a = makeSp("rect", 100, 100, 50, 40, { name: "a" });
    const b = makeSp("ellipse", 160, 120, 40, 40, { name: "b" });
    return { kind: "group", id: "g", name: "g", x: 100, y: 100, w: 100, h: 60, rot: 0, children: [a, b] };
  };
  const scale = (g: GroupShape, k: number, dx: number): GroupShape => ({
    ...g,
    x: g.x + dx,
    w: g.w * k,
    h: g.h * k,
    children: g.children.map((c) => ({ ...c, id: `${c.id}-x`, x: g.x + dx + (c.x - g.x) * k, y: g.y + (c.y - g.y) * k, w: c.w * k, h: c.h * k })),
  });

  it("taşıma, ölçekleme ve kimlik değişimi imzayı bozmaz; içerik değişimi bozar", () => {
    const g = group();
    g.raw = { kind: "smartart", src: "x", xml: "", refs: [], extra: [], sig: rawSig(g) };
    expect(rawSigMatches(scale(g, 1.37, 33.3))).toBe(true);
    const recolored = { ...g, children: g.children.map((c, i) => (i ? { ...c, fill: { type: "solid", color: { rgb: "#FF0000" } } } : c)) } as GroupShape;
    expect(rawSigMatches(recolored)).toBe(false);
    expect(rawSigMatches({ ...g, rot: 15 })).toBe(false);
    const nudged = { ...g, children: g.children.map((c, i) => (i ? { ...c, x: c.x + 5 } : c)) };
    expect(rawSigMatches(nudged)).toBe(false);
  });
});
