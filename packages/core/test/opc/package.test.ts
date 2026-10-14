import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  ContentTypes,
  explicitRelIds,
  OpcPackage,
  parseRels,
  partOfRels,
  relativeTarget,
  relsPathOf,
  RelsBuilder,
  resolveTarget,
  rewriteRelIds,
} from "../../src/opc/package";
import { OpcError } from "../../src/opc/zip";
import { BASE_DEFAULTS, contentTypes, CT_MAIN_DOCX, DOC_XML, poi, poiDosyalari, R, rels, T, zipOf } from "./paket";

describe("içerik türleri", () => {
  it("Override önce, sonra uzantı varsayılanı; büyük/küçük harf ve yüzde kodlaması", () => {
    const ct = ContentTypes.parse(
      contentTypes({ ...BASE_DEFAULTS, PNG: "image/png" }, { "word/document.xml": CT_MAIN_DOCX, "word/media/resim%201.png": "image/x-ozel" }),
    );
    expect(ct.get("word/document.xml")).toBe(CT_MAIN_DOCX);
    expect(ct.get("WORD/Document.XML")).toBe(CT_MAIN_DOCX);
    expect(ct.get("word/styles.xml")).toBe("application/xml");
    expect(ct.get("word/media/a.png")).toBe("image/png");
    expect(ct.get("word/media/resim 1.png")).toBe("image/x-ozel");
    expect(ct.get("dosya")).toBeUndefined();
  });

  it("yazılıp yeniden okunduğunda aynı; boş ya da bozuk metin boş model", () => {
    const ct = ContentTypes.parse(contentTypes(BASE_DEFAULTS, { "a b/ç.xml": "x/y" }));
    ct.setDefault("bin", "application/vnd.ms-office.vbaProject");
    ct.setOverride("word/vbaData.xml", "application/vnd.ms-word.vbaData+xml");
    const again = ContentTypes.parse(ct.toXml());
    expect(again.get("a b/ç.xml")).toBe("x/y");
    expect(again.get("x/vbaProject.bin")).toBe("application/vnd.ms-office.vbaProject");
    expect(again.overrideEntries()).toEqual(ct.overrideEntries());
    expect(ContentTypes.parse("<Types><Default Extension=").defaultEntries()).toEqual([]);
    expect(ContentTypes.parse(null).get("a.xml")).toBeUndefined();
  });
});

describe("ilişki yolları", () => {
  it("hedef çözme: göreli, mutlak, üst dizin, parça bölümü, yüzde kodu", () => {
    expect(resolveTarget("word/document.xml", "media/image1.png")).toBe("word/media/image1.png");
    expect(resolveTarget("word/document.xml", "../customXml/item1.xml")).toBe("customXml/item1.xml");
    expect(resolveTarget("word/document.xml", "/word/theme/theme1.xml")).toBe("word/theme/theme1.xml");
    expect(resolveTarget("", "word/document.xml")).toBe("word/document.xml");
    expect(resolveTarget("word/document.xml", "./a/./b.xml#parca")).toBe("word/a/b.xml");
    expect(resolveTarget("word/document.xml", "media/resim%201.png")).toBe("word/media/resim 1.png");
  });

  it("kökün dışına çıkan, boş ve şemalı hedef null", () => {
    expect(resolveTarget("word/document.xml", "../../x.xml")).toBeNull();
    expect(resolveTarget("word/document.xml", "#_ftn1")).toBeNull();
    expect(resolveTarget("word/document.xml", "")).toBeNull();
    expect(resolveTarget("word/document.xml", "http://ornek.com/a.xml")).toBeNull();
    expect(resolveTarget("word/document.xml", "file:///C:/a.xml")).toBeNull();
  });

  it("göreli hedef, çözülünce aynı parçayı verir (rastgele yollar)", () => {
    const seg = fc.stringMatching(/^[A-Za-z0-9 çğıöşüİ_.%#-]{1,8}$/).filter((s) => s !== "." && s !== ".." && !s.includes("#"));
    fc.assert(
      fc.property(fc.array(seg, { minLength: 1, maxLength: 4 }), fc.array(seg, { minLength: 1, maxLength: 4 }), (a, b) => {
        const source = a.join("/");
        const part = b.join("/");
        expect(resolveTarget(source, relativeTarget(source, part))).toBe(part);
        expect(resolveTarget("", relativeTarget("", part))).toBe(part);
      }),
      { numRuns: 500 },
    );
  });

  it("ilişki dosyası yolu ve tersi", () => {
    expect(relsPathOf("")).toBe("_rels/.rels");
    expect(relsPathOf("word/document.xml")).toBe("word/_rels/document.xml.rels");
    expect(partOfRels("word/_rels/document.xml.rels")).toBe("word/document.xml");
    expect(partOfRels("_rels/.rels")).toBe("");
    expect(partOfRels("word/document.xml")).toBeNull();
  });
});

describe("ilişkiler", () => {
  it("bozuk ve tuhaf .rels: okunabilenler döner, eksik öznitelikli ve yinelenen Id atlanır", () => {
    const xml = `<Relationships><Relationship Id='rId1' Type='${T("theme")}' Target='theme/theme1.xml'/>
      <Relationship Id="rId2" Type="${T("hyperlink")}" Target="http://a.com/?x=1&amp;y=2" TargetMode="EXTERNAL"/>
      <Relationship Id="rId1" Type="${T("styles")}" Target="styles.xml"/>
      <Relationship Type="${T("settings")}" Target="settings.xml"/>
      <Relationship Id="rId9" Type="${T("x")}"`;
    const list = parseRels(xml, "word/document.xml");
    expect(list.map((r) => r.id)).toEqual(["rId1", "rId2"]);
    expect(list[0].part).toBe("word/theme/theme1.xml");
    expect(list[1]).toMatchObject({ external: true, part: null, target: "http://a.com/?x=1&y=2" });
    expect(parseRels("çöp <<<>>>", "")).toEqual([]);
    expect(parseRels(null, "")).toEqual([]);
  });

  it("RelsBuilder eski kimlikleri asla yeniden kullanmaz, hedefi göreli yazar", () => {
    const b = new RelsBuilder("word/document.xml", parseRels(rels([{ id: "rId1", type: T("styles"), target: "styles.xml" }, { id: "rId3", type: T("x"), target: "x.xml" }]), "word/document.xml"));
    const a = b.addPart(T("theme"), "word/theme/theme1.xml");
    const c = b.addPart(T("customXml"), "customXml/item1.xml");
    const d = b.addExternal(T("hyperlink"), "https://ornek.com");
    expect(new Set([a, c, d, "rId1", "rId3"]).size).toBe(5);
    const again = parseRels(b.xml(), "word/document.xml");
    expect(again.find((r) => r.id === c)).toMatchObject({ target: "../customXml/item1.xml", part: "customXml/item1.xml" });
    expect(again.find((r) => r.id === d)).toMatchObject({ external: true });
  });
});

describe("açık ilişki kimlikleri", () => {
  it("ilişki ad alanına bağlı her önek ve öznitelik; VML o:relid; öneksiz relId", () => {
    const xml = `<w:document xmlns:w="w" xmlns:r="${R}" xmlns:rel="${R}" xmlns:s="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">
      <a:blip r:embed="rId5" r:link='rId6'/><c:chart rel:id="rId7"/><x s:id="rId8"/><v:imagedata o:relid="rId9" w:id="rId99"/>
      <dsp:dataModelExt relId="rId10"/><w:bookmarkStart w:id="0" w:name="rId11"/>`;
    expect([...explicitRelIds(xml)].sort()).toEqual(["rId10", "rId5", "rId6", "rId7", "rId8", "rId9"]);
  });

  it("bildirimsiz parçacıkta r öneki varsayılır; ek önek verilebilir", () => {
    expect([...explicitRelIds(`<a:blip r:embed="rId2"/>`)]).toEqual(["rId2"]);
    expect([...explicitRelIds(`<a:blip q:embed="rId2"/>`, { prefixes: ["q"] })]).toEqual(["rId2"]);
    expect(explicitRelIds(`<c r="A1" t="s"><v>rId1</v></c>`).size).toBe(0);
  });

  it("kimlikleri tek geçişte yeniden yazar (takas güvenli), başka öznitelikleri değiştirmez", () => {
    const xml = `<p xmlns:r="${R}"><a r:id="rId1" w:id="rId1"/><b r:embed="rId2"/></p>`;
    const out = rewriteRelIds(xml, new Map([["rId1", "rId2"], ["rId2", "rId1"]]));
    expect(out).toBe(`<p xmlns:r="${R}"><a r:id="rId2" w:id="rId1"/><b r:embed="rId1"/></p>`);
  });
});

describe("OpcPackage", () => {
  it("ana parça, içerik türü, ilişkiler, metin; olmayan parça null", async () => {
    const b = await zipOf({
      "[Content_Types].xml": contentTypes(BASE_DEFAULTS, { "word/document.xml": CT_MAIN_DOCX }),
      "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "/word/document.xml" }]),
      "word/document.xml": DOC_XML(`<w:p><w:r><w:drawing><a:blip r:embed="rId4"/></w:drawing></w:r></w:p>`),
      "word/_rels/document.xml.rels": rels([{ id: "rId4", type: T("image"), target: "media/a.png" }, { id: "rId5", type: T("theme"), target: "theme/theme1.xml" }]),
    });
    const pkg = await OpcPackage.open(b);
    expect(pkg.isOpc).toBe(true);
    expect(pkg.mainPart).toBe("word/document.xml");
    expect(pkg.mainContentType).toBe(CT_MAIN_DOCX);
    expect(pkg.macroEnabled).toBe(false);
    expect(pkg.signed).toBe(false);
    expect((await pkg.rels("word/document.xml")).map((r) => r.part)).toEqual(["word/media/a.png", "word/theme/theme1.xml"]);
    expect([...(await pkg.explicitRelIds("word/document.xml"))]).toEqual(["rId4"]);
    expect(await pkg.text("WORD/DOCUMENT.XML")).toContain("<w:body>");
    expect(await pkg.data("yok.xml")).toBeNull();
    expect(pkg.has("Word/Document.xml")).toBe(true);
  });

  it("çizge erişilebilirliği döngülü .rels'te sonlanır", async () => {
    const b = await zipOf({
      "[Content_Types].xml": contentTypes(BASE_DEFAULTS, {}),
      "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "a.xml" }]),
      "a.xml": "<a/>",
      "b.xml": "<b/>",
      "yetim.xml": "<y/>",
      "_rels/a.xml.rels": rels([{ id: "r1", type: T("x"), target: "b.xml" }, { id: "r2", type: T("x"), target: "a.xml" }, { id: "r3", type: T("x"), target: "yok.xml" }]),
      "_rels/b.xml.rels": rels([{ id: "r1", type: T("x"), target: "a.xml" }, { id: "r2", type: T("x"), target: "/b.xml" }]),
    });
    const pkg = await OpcPackage.open(b);
    expect([...(await pkg.graph.reachable())].sort()).toEqual(["a.xml", "b.xml"]);
  });

  it("ZIP değilse OpcError; içerik türleri yoksa hata değil", async () => {
    await expect(OpcPackage.open(new TextEncoder().encode("düz metin dosyası, zip değil................"))).rejects.toBeInstanceOf(OpcError);
    const pkg = await OpcPackage.open(await zipOf({ "a.txt": "x" }));
    expect(pkg.isOpc).toBe(false);
    expect(pkg.mainPart).toBeNull();
  });

  it("parça boyutu sınırı: büyük parça 'too-large', paket yine açılır", async () => {
    const pkg = await OpcPackage.open(await zipOf({ "[Content_Types].xml": contentTypes(BASE_DEFAULTS, {}), "buyuk.xml": "x".repeat(200_000) }), { maxPartSize: 100_000 });
    await expect(pkg.text("buyuk.xml")).rejects.toMatchObject({ code: "too-large" });
    expect(pkg.sizeOf("buyuk.xml")).toBe(200_000);
  });

  it("POI dosyaları: ana parça, makro ve imza bilgisi", async () => {
    for (const f of poiDosyalari(/\.(docx|docm|xlsx|xlsm|pptx|pptm)$/)) {
      const pkg = await OpcPackage.open(poi(f));
      expect(pkg.mainPart, f).toBeTruthy();
      expect(pkg.macroEnabled, f).toBe(/\.(docm|xlsm|pptm)$/.test(f));
      expect(pkg.signed, f).toBe(/signed/.test(f));
      // tüm .rels dosyaları okunabiliyor, hedefler çözülüyor
      const reach = await pkg.graph.reachable();
      expect(reach.has(pkg.mainPart!), f).toBe(true);
    }
  });
});
