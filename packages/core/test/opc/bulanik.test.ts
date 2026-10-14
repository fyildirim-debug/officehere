// Bulanık dayanıklılık: bozuk ve döngülü .rels, mutlak / kökün dışına çıkan / şemalı hedefler, ayrılmış adlar,
// eksik içerik türleri, bozuk ZIP girdileri ve çok büyük parçalarla taşıma çökmez, takılmaz ve paketi bozmaz.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { Carry, carryParts } from "../../src/opc/carry";
import { OpcPackage } from "../../src/opc/package";
import { SourceRegistry } from "../../src/opc/source";
import { ZipWriter } from "../../src/opc/zip";
import { BASE_DEFAULTS, contentTypes, CT_MAIN_DOCX, DOC_XML, minimalDocx, MST, opcHatalari, rels, T } from "./paket";

const PARCALAR = ["word/document.xml", "word/theme/theme1.xml", "word/a.xml", "word/alt/b.xml", "customXml/item1.xml", "customXml/itemProps1.xml", "c.xml", "word/media/r.png", "word/vbaProject.bin", "_xmlsignatures/origin.sigs"];
const HEDEFLER = [
  "theme/theme1.xml", "a.xml", "alt/b.xml", "../customXml/item1.xml", "itemProps1.xml", "../c.xml", "/c.xml", "/word/a.xml", "media/r.png", "vbaProject.bin", "../_xmlsignatures/origin.sigs",
  "../../../../dışarı.xml", "yok.xml", "", "#yer", "http://ornek.com/x", "file:///C:/a.xml", "[Content_Types].xml", "../[Content_Types].xml", "_rels/document.xml.rels",
  "../_rels/.rels", "document.xml", "A.XML", "a.xml#parca", "ç%C3%A7.xml", "alt/../a.xml", "./a.xml",
];
const TURLER = [
  T("theme"), T("fontTable"), T("customXml"), T("customXmlProps"), T("styles"), T("image"), T("officeDocument"), T("glossaryDocument"), T("custom-properties"), T("worksheet"), T("slide"),
  MST("2006/relationships/vbaProject"), MST("2011/relationships/people"), "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin", "http://ornek.com/bilinmeyen", "",
];
const ICERIK = [
  "<a/>", DOC_XML(`<w:p><w:r><a:blip r:embed="rId1"/><v:imagedata o:relid="rId2"/></w:r></w:p>`), `<k:Kok xmlns:k="urn:x"/>`, "", "çöp <<< > &&& \u0000",
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property pid="2" name="A"><vt:lpwstr>1</vt:lpwstr></property>`,
];

const relArb = fc.record({
  id: fc.constantFrom("rId1", "rId2", "rId3", "rId1", "", "rId 9", "r&amp;"),
  type: fc.constantFrom(...TURLER),
  target: fc.constantFrom(...HEDEFLER),
  external: fc.boolean(),
});
const relsArb = fc.oneof(
  { weight: 4, arbitrary: fc.array(relArb, { maxLength: 6 }).map((l) => rels(l.map((r) => ({ ...r, external: r.external && r.target.startsWith("http") })))) },
  { weight: 1, arbitrary: fc.tuple(fc.array(relArb, { maxLength: 4 }), fc.nat()).map(([l, n]) => { const x = rels(l); return x.slice(0, n % (x.length + 1)); }) },
  { weight: 1, arbitrary: fc.string({ maxLength: 40 }) },
);
const kaynakArb = fc.record({
  parcalar: fc.dictionary(fc.constantFrom(...PARCALAR), fc.constantFrom(...ICERIK), { minKeys: 2, maxKeys: PARCALAR.length }),
  belge: fc.constantFrom(...ICERIK),
  // kök ve ana parça ilişkileri çoğunlukla okunabilir, ara sıra bozuk
  kok: fc.oneof(
    { weight: 4, arbitrary: fc.array(relArb, { maxLength: 3 }).map((l) => rels([{ id: "rIdK", type: T("officeDocument"), target: "word/document.xml" }, ...l.map((r) => ({ ...r, id: `${r.id}k`, external: false }))])) },
    { weight: 1, arbitrary: relsArb },
  ),
  belgeIliski: fc.oneof({ weight: 4, arbitrary: fc.array(relArb, { minLength: 1, maxLength: 8 }).map((l) => rels(l.map((r) => ({ ...r, external: r.external && r.target.startsWith("http") })))) }, { weight: 1, arbitrary: relsArb }),
  iliskiler: fc.dictionary(fc.constantFrom(...PARCALAR.filter((p) => p !== "word/document.xml")), relsArb, { maxKeys: 6 }),
  ct: fc.oneof(
    { weight: 3, arbitrary: fc.constant(contentTypes(BASE_DEFAULTS, { "word/document.xml": CT_MAIN_DOCX })) },
    { weight: 1, arbitrary: fc.constant(contentTypes({ xml: "application/xml" }, {})) },
    { weight: 1, arbitrary: fc.constant("") },
    { weight: 1, arbitrary: fc.constant('<Types><Override PartName="/word/a.xml"') },
  ),
  bozukGirdi: fc.boolean(),
});

interface Kaynak {
  parcalar: Record<string, string>;
  belge: string;
  kok: string;
  belgeIliski: string;
  iliskiler: Record<string, string>;
  ct: string;
  bozukGirdi: boolean;
}

async function kaynakKur(k: Kaynak): Promise<Uint8Array> {
  const z = new ZipWriter();
  if (k.ct) z.add("[Content_Types].xml", k.ct);
  z.add("_rels/.rels", k.kok);
  z.add("word/document.xml", k.belge);
  z.add("word/_rels/document.xml.rels", k.belgeIliski);
  for (const [ad, x] of Object.entries(k.parcalar)) if (ad !== "word/document.xml") z.add(ad, x);
  for (const [ad, x] of Object.entries(k.iliskiler)) {
    const i = ad.lastIndexOf("/");
    z.add(`${ad.slice(0, i + 1)}_rels/${ad.slice(i + 1)}.rels`, x);
  }
  const b = await z.finish();
  if (k.bozukGirdi) {
    // ilk girdinin verisini boz: yerel başlıktan sonraki baytlar
    for (let i = 60; i < Math.min(b.length - 100, 90); i++) b[i] = 0xff;
  }
  return b;
}

/** Taşınan parçaların kendi XML'inden gelen kırık kimlikler (kaynağın sorunu) dışında paket hatası yok. */
const paketHatalari = (b: Uint8Array, macro = false) => opcHatalari(b, macro ? ".docm" : ".docx").filter((h) => h.tur !== "kırık-rId");

describe("bulanık dayanıklılık", () => {
  it("rastgele bozuk kaynak: taşıma çökmez, takılmaz, çıktı paketi tutarlı", async () => {
    const out = await minimalDocx();
    const temiz = paketHatalari(out).filter((h) => h.tur !== "ct-ilk-değil");
    expect(temiz).toEqual([]);
    const sayac = { kosu: 0, tasinan: 0, dusen: 0 };
    await fc.assert(
      fc.asyncProperty(kaynakArb, fc.boolean(), async (k, macro) => {
        const b = await kaynakKur(k);
        let src: OpcPackage | undefined;
        try {
          src = await OpcPackage.open(b);
        } catch (e) {
          expect((e as Error).name).toBe("OpcError");
          return;
        }
        const t0 = performance.now();
        const { bytes, report } = await carryParts({ app: "belge", source: src, output: out, macro });
        sayac.kosu++;
        if (report.carried.length) sayac.tasinan++;
        if (report.dropped.length) sayac.dusen++;
        expect(performance.now() - t0).toBeLessThan(2000);
        expect(paketHatalari(bytes, macro)).toEqual([]);
        // ikinci tur: çıktı yeniden kaynak olunca da tutarlı
        const again = await carryParts({ app: "belge", source: await OpcPackage.open(bytes), output: out, macro });
        expect(paketHatalari(again.bytes, macro)).toEqual([]);
        for (const c of report.carried) if (c.kind !== "external") expect(c.to.toLowerCase().endsWith(".rels")).toBe(false);
      }),
      { numRuns: 300 },
    );
    // üretici anlamlı: koşuların önemli bir kısmında taşıma ve düşürme gerçekleşiyor
    // Üreteç işe yarar örnek üretiyor mu: ortalama ~%20; eşik %10 (rastgele tohumla ~4 standart sapma pay,
    // %20 eşiği ortalamanın tam üstündeydi ve testi yarı yarıya düşürüyordu).
    expect(sayac.tasinan).toBeGreaterThan(sayac.kosu / 10);
    expect(sayac.dusen).toBeGreaterThan(sayac.kosu / 10);
  }, 120_000);

  it("rastgele bozuk kaynak: öğe düzeyi kopyalar ve çöp toplama da tutarlı", async () => {
    await fc.assert(
      fc.asyncProperty(kaynakArb, fc.array(fc.constantFrom(...PARCALAR), { maxLength: 4 }), async (k, kopyalar) => {
        const b = await kaynakKur(k);
        const src = await OpcPackage.open(b).catch(() => undefined);
        if (!src) return;
        const reg = new SourceRegistry();
        const id = reg.register(src);
        const carry = new Carry(reg);
        const inst = carry.instance(id);
        const docRels = [{ id: "rId1", type: T("styles"), target: "styles.xml" }];
        kopyalar.forEach((p, i) => {
          const to = inst.copy(p);
          // yarısını yazıcı bağlar, yarısı çöp toplanmalı
          if (to && i % 2 === 0) docRels.push({ id: `rId${10 + i}`, type: "http://ornek.com/opak", target: `/${to}` });
        });
        const out = await minimalDocx({ "word/_rels/document.xml.rels": rels(docRels) });
        const { bytes } = await carryParts({ app: "belge", source: id, output: out, carry, registry: reg });
        expect(paketHatalari(bytes)).toEqual([]);
      }),
      { numRuns: 200 },
    );
  }, 120_000);

  it("çok büyük parçalar: taşınan büyük parça açılmadan ham kopyalanır; büyük ana parça ve .rels sınırla kesilir", async () => {
    const buyuk = new Uint8Array(48 * 1024 * 1024); // sıfırlar: ~48 KB'a sıkışır
    const z = new ZipWriter();
    z.add("[Content_Types].xml", contentTypes({ ...BASE_DEFAULTS, bin: "application/octet-stream" }, { "word/document.xml": CT_MAIN_DOCX }));
    z.add("_rels/.rels", rels([{ id: "rId1", type: T("officeDocument"), target: "word/document.xml" }, { id: "rId2", type: "http://ornek.com/buyuk", target: "buyuk.bin" }]));
    z.add("word/document.xml", DOC_XML());
    z.add("word/_rels/document.xml.rels", rels([{ id: "rId1", type: T("theme"), target: "theme/theme1.xml" }]));
    z.add("word/theme/theme1.xml", "<a:theme/>");
    z.add("buyuk.bin", buyuk);
    const src = await OpcPackage.open(await z.finish());
    let t0 = performance.now();
    const { bytes, report } = await carryParts({ app: "belge", source: src, output: await minimalDocx() });
    expect(performance.now() - t0).toBeLessThan(1000);
    expect(report.carried.map((c) => c.to)).toEqual(["buyuk.bin", "word/theme/theme1.xml"]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.entry("buyuk.bin")!.size).toBe(buyuk.length);
    expect(pkg.entry("buyuk.bin")!.compressedSize).toBe(src.entry("buyuk.bin")!.compressedSize);

    // Ana parça sınırı aşarsa açık/örtük ayrımı yapılamaz: ana parçanın ilişkileri taşınmaz (güvenli olan), çökmez
    const z2 = new ZipWriter();
    z2.add("[Content_Types].xml", contentTypes(BASE_DEFAULTS, { "word/document.xml": CT_MAIN_DOCX }));
    z2.add("_rels/.rels", rels([{ id: "rId1", type: T("officeDocument"), target: "word/document.xml" }]));
    z2.add("word/document.xml", DOC_XML("<w:p/>".repeat(400_000)));
    z2.add("word/_rels/document.xml.rels", rels([{ id: "rId1", type: T("theme"), target: "theme/theme1.xml" }]));
    z2.add("word/theme/theme1.xml", "<a:theme/>");
    // .rels bombası: 20 MB boşluk (sınır 16 MB) → bozuk sayılır
    z2.add("word/theme/_rels/theme1.xml.rels", " ".repeat(20 * 1024 * 1024));
    const src2 = await OpcPackage.open(await z2.finish(), { maxPartSize: 1024 * 1024 });
    t0 = performance.now();
    const r2 = await carryParts({ app: "belge", source: src2, output: await minimalDocx() });
    expect(performance.now() - t0).toBeLessThan(2000);
    expect(r2.report.carried).toEqual([]);
    expect(r2.report.explicit.map((e) => e.part)).toEqual(["word/theme/theme1.xml"]);
    expect(paketHatalari(r2.bytes)).toEqual([]);
    // sınır yüksekken tema taşınır, bomba .rels'i boş sayılır
    const src3 = await OpcPackage.open(await z2.finish());
    const r3 = await carryParts({ app: "belge", source: src3, output: await minimalDocx() });
    expect(r3.report.carried.map((c) => c.from)).toEqual(["word/theme/theme1.xml"]);
    expect(paketHatalari(r3.bytes)).toEqual([]);
  }, 60_000);

  it("kaynak yoksa (null) yalnız içerik türleri ilk parça olacak biçimde yeniden kurulur", async () => {
    const out = await minimalDocx();
    const { bytes, report } = await carryParts({ app: "belge", source: null, output: out });
    expect(report.carried).toEqual([]);
    expect(opcHatalari(bytes)).toEqual([]);
    // ikinci kez: zaten ilk parçaysa baytlar aynen döner
    const again = await carryParts({ app: "belge", source: null, output: bytes });
    expect(again.bytes).toBe(bytes);
  });

  it("çıktı ZIP değilse OpcError", async () => {
    await expect(carryParts({ app: "belge", source: null, output: new TextEncoder().encode("zip değil".repeat(10)) })).rejects.toMatchObject({ name: "OpcError" });
  });
});
