// @vitest-environment jsdom
/**
 * PPTX okuyucunun parça modülleri (io/pptxReader/*): asıl slayt ve düzenler,
 * tema, ortam/grafik, slaytlar, açıklamalar, bölümler ve sunu bölümü. Özellikle
 * kurtarma ve uyarı yolları; tam dosya davranışı robustness/pptx testlerinde.
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { Presentation, SpShape } from "../../model/types";
import { LoadWarnings } from "../loadWarnings";
import { PptxError, readPptx } from "../pptxReader";
import { PackageError, ZIP_LIMITS } from "../zipSafe";

vi.setConfig({ testTimeout: 30_000 });

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const load = (n: string) => new Uint8Array(readFileSync(`${FIX}/${n}`));

async function mutate(name: string, fn: (z: JSZip) => void | Promise<void>): Promise<Uint8Array> {
  const z = await JSZip.loadAsync(load(name));
  await fn(z);
  return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
/** Parçanın metnini değiştirir (yoksa test hatası). */
async function edit(z: JSZip, path: string, fn: (xml: string) => string): Promise<void> {
  const f = z.file(path);
  if (!f) throw new Error(`fikstürde yok: ${path}`);
  z.file(path, fn(await f.async("string")));
}
async function read(bytes: Uint8Array): Promise<{ pres: Presentation; warnings: string[]; details: string[] }> {
  const w = new LoadWarnings();
  const pres = await readPptx(bytes, { warnings: w });
  return { pres, warnings: w.list(), details: w.details };
}
const layoutName = (p: Presentation, i: number) => p.layouts.find((l) => l.id === p.slides[i].layoutId)?.name;
const placeholder = (shapes: Presentation["layouts"][number]["shapes"], type: string) => shapes.find((s): s is SpShape => s.kind === "sp" && s.ph?.type === type);

const LIMITS = { ...ZIP_LIMITS };
afterEach(() => {
  Object.assign(ZIP_LIMITS, LIMITS);
  vi.restoreAllMocks();
});

describe("pptxReader: dış API", () => {
  it("dizin içe aktarması readPptx ve PptxError'ı verir; ZIP olmayan veri Türkçe PptxError", async () => {
    expect(new PptxError("x")).toBeInstanceOf(PackageError);
    const e = await readPptx(new TextEncoder().encode("bu bir zip değil")).catch((x) => x);
    expect(e).toBeInstanceOf(PptxError);
    expect(e.message).toBe("Dosya geçerli bir PPTX (ZIP) paketi değil.");
  });

  it("uyarı nesnesi verilmeden de okur", async () => {
    const pres = await readPptx(load("lo-bicim.ppsx"));
    expect(pres.slides).toHaveLength(4);
  });
});

describe("pptxReader: asıl slaytlar ve düzenler", () => {
  it("ikincil asıl slaydın farklı metin stilleri ve teması yalnız kendi düzenine taşınır", async () => {
    const bytes = await mutate("lo-proje.pptx", async (z) => {
      await edit(z, "ppt/slideMasters/slideMaster2.xml", (x) => x.replace("</p:sldMaster>", '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="5400"/></a:lvl1pPr></p:titleStyle></p:txStyles></p:sldMaster>'));
      await edit(z, "ppt/theme/theme2.xml", (x) => x.replace('<a:accent1><a:srgbClr val="087e8b"/>', '<a:accent1><a:srgbClr val="123456"/>'));
    });
    const { pres, warnings } = await read(bytes);
    expect(warnings).toEqual([]);
    const themed = pres.layouts.filter((l) => l.theme);
    expect(themed).toHaveLength(1);
    expect(themed[0].theme!.colors.accent1).toBe("#123456");
    expect(themed[0].showMasterShapes).toBe(false);
    expect(placeholder(themed[0].shapes, "title")?.text?.levels?.[0]?.run?.size).toBe(54);
    // Ana asıl slayt ve onun düzenleri değişmez.
    expect(pres.theme.colors.accent1).not.toBe("#123456");
    expect(pres.master.textStyles.title[0].run.size).not.toBe(54);
  });

  it("okunamayan ikincil asıl slayt atlanır; düzeni olmayan slaytlar ilk düzene bağlanır", async () => {
    const orig = await read(load("lo-proje.pptx"));
    const { pres, warnings, details } = await read(await mutate("lo-proje.pptx", (z) => void z.file("ppt/slideMasters/slideMaster2.xml", "<bozuk")));
    expect(warnings).toContain("Bazı asıl slaytlar okunamadı; onları kullanan slaytlar başka bir düzene bağlandı.");
    expect(details.join("\n")).toContain("ppt/slideMasters/slideMaster2.xml");
    expect(pres.slides).toHaveLength(orig.pres.slides.length);
    expect(pres.layouts.length).toBeLessThan(orig.pres.layouts.length);
  });

  it("asıl slaydın ilişkileri yoksa düzenler paket içinden bulunur", async () => {
    const orig = await read(load("lo-bicim.ppsx"));
    const { pres, warnings } = await read(await mutate("lo-bicim.ppsx", (z) => void z.remove("ppt/slideMasters/_rels/slideMaster1.xml.rels")));
    expect(warnings.join(" ")).toMatch(/Asıl slaydın ilişki bölümü eksik ya da bozuk; düzenler paket içinden bulundu/);
    expect(pres.layouts.map((l) => l.name)).toEqual(orig.pres.layouts.map((l) => l.name));
    expect(pres.slides.map((_, i) => layoutName(pres, i))).toEqual(orig.pres.slides.map((_, i) => layoutName(orig.pres, i)));
  });

  it("bozuk düzen ve ilişkisi eksik düzen uyarı verir; hiç düzen okunamazsa boş düzen eklenir", async () => {
    const r1 = await read(
      await mutate("lo-bicim.ppsx", async (z) => {
        z.file("ppt/slideLayouts/slideLayout1.xml", "<bozuk");
        await edit(z, "ppt/slideLayouts/slideLayout2.xml", (x) => x.replace("<p:cNvPr ", '<p:cNvPr r:id="rId9" '));
        z.remove("ppt/slideLayouts/_rels/slideLayout2.xml.rels");
      }),
    );
    expect(r1.warnings).toContain("Bazı slayt düzenleri okunamadı ya da bozuk; ilgili slaytlar başka bir düzene bağlanmış olabilir.");
    expect(r1.details.join("\n")).toContain("ppt/slideLayouts/slideLayout1.xml");
    expect(r1.details.join("\n")).toContain("ppt/slideLayouts/_rels/slideLayout2.xml.rels: pakette yok ya da okunamadı");
    expect(r1.pres.layouts).toHaveLength(1);

    const r2 = await read(
      await mutate("lo-bicim.ppsx", (z) => {
        z.remove("ppt/slideLayouts/slideLayout1.xml");
        z.remove("ppt/slideLayouts/slideLayout2.xml");
      }),
    );
    expect(r2.pres.layouts).toEqual([expect.objectContaining({ type: "blank", name: "Boş", shapes: [] })]);
    expect(r2.pres.slides.every((s) => s.layoutId === r2.pres.layouts[0].id)).toBe(true);
  });
});

describe("pptxReader: tema", () => {
  it("asıl slaydın tema ilişkisi bozuksa not asıl sayfasının teması değil, sahipsiz tema kullanılır", async () => {
    const orig = await read(load("lo-bicim.ppsx"));
    const { pres, warnings, details } = await read(
      await mutate("lo-bicim.ppsx", (z) => edit(z, "ppt/slideMasters/_rels/slideMaster1.xml.rels", (x) => x.replace("../theme/theme1.xml", "../theme/theme9.xml"))),
    );
    expect(details).toContain("tema ppt/theme/theme9.xml yerine ppt/theme/theme1.xml kullanıldı");
    expect(warnings).toContain("Sununun ilişki bölümü bozuk; asıl slayt ve tema paket içinden bulunarak kullanıldı.");
    expect(pres.theme).toEqual(orig.pres.theme);
  });

  it("yalnız başka asıl sayfaların temaları kaldıysa varsayılan tema kullanılır", async () => {
    const { pres, warnings } = await read(await mutate("lo-bicim.ppsx", (z) => void z.remove("ppt/theme/theme1.xml")));
    expect(warnings).toContain("Tema okunamadı; varsayılan tema renkleri ve yazı tipleri kullanıldı.");
    expect(pres.theme.fonts).toEqual({ major: "Calibri Light", minor: "Calibri" });
  });
});

describe("pptxReader: ortam dosyaları ve grafikler", () => {
  it("sınırı aşan resim boş çerçeveyle, okunamayan grafik uyarıyla açılır", async () => {
    ZIP_LIMITS.media = 16;
    const { pres, warnings, details } = await read(await mutate("lo-bicim.ppsx", (z) => void z.file("ppt/charts/chart1.xml", "")));
    expect(warnings).toContain("1 resim ya da ortam dosyası bozuk ya da çok büyük olduğu için yüklenemedi; yerlerinde boş çerçeve görünür.");
    expect(warnings).toContain("Bazı grafikler okunamadı; varsa yedek resimleri gösteriliyor.");
    expect(details.join("\n")).toContain("ppt/media/image1.png");
    expect(pres.slides).toHaveLength(4);
  });
});

describe("pptxReader: slaytlar", () => {
  it("5.000'den çok slayt girdisi kırpılır ve uyarılır", async () => {
    const bytes = await mutate("lo-bicim.ppsx", (z) =>
      edit(z, "ppt/presentation.xml", (x) => x.replace("</p:sldIdLst>", `${'<p:sldId id="300" r:id="rId4"/>'.repeat(4998)}</p:sldIdLst>`)),
    );
    const { pres, warnings } = await read(bytes);
    expect(warnings).toContain("Sunuda 5.002 slayt girdisi var; yalnız ilk 5.000 slayt açıldı.");
    expect(pres.slides).toHaveLength(5000);
  });

  it("arka plan başvurusu (p:bgRef) düz dolgu olarak okunur", async () => {
    const bytes = await mutate("lo-bicim.ppsx", (z) => edit(z, "ppt/slides/slide1.xml", (x) => x.replace("<p:cSld>", '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg>')));
    const { pres } = await read(bytes);
    expect(pres.slides[0].background).toEqual({ type: "solid", color: { scheme: "accent2" } });
  });
});

describe("pptxReader: açıklamalar", () => {
  it("bilinmeyen yazar 'Yazar', geçersiz tarih okuma anı olur", async () => {
    vi.spyOn(Date, "now").mockReturnValue(1_700_000_000_000);
    const bytes = await mutate("lo-yorum.pptx", (z) => edit(z, "ppt/comments/comment1.xml", (x) => x.replace('authorId="0" dt="2026-10-12T00:52:45.982000000"', 'authorId="7" dt="dün"')));
    const { pres } = await read(bytes);
    const [c1, c2] = pres.slides[0].comments!;
    expect(c1).toMatchObject({ author: "Yazar", date: 1_700_000_000_000, text: "İlk açıklama: başlık rengi koyulaşsın." });
    expect(c2.author).toBe("Kullanıcı");
  });

  it("bozuk açıklama parçası slaydı düşürmez, uyarı verir", async () => {
    const { pres, warnings } = await read(await mutate("lo-yorum.pptx", (z) => void z.file("ppt/comments/comment1.xml", "<bozuk")));
    expect(pres.slides).toHaveLength(1);
    expect(pres.slides[0].comments).toBeUndefined();
    expect(warnings).toEqual(["1. slaydın açıklamaları okunamadı."]);
  });
});

describe("pptxReader: bölümler", () => {
  it("p14:sectionLst bölümleri ilk slaytlara işlenir; diğer uzantılar ve bilinmeyen slaytlar yok sayılır", async () => {
    const ext =
      '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext>' +
      '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
      '<p14:section name="Giriş" id="{1}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>' +
      '<p14:section name="Kayıp" id="{2}"><p14:sldIdLst><p14:sldId id="999"/></p14:sldIdLst></p14:section>' +
      '<p14:section name="Sonuç" id="{3}"><p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/></p14:sldIdLst></p14:section>' +
      "</p14:sectionLst></p:ext></p:extLst>";
    const bytes = await mutate("lo-bicim.ppsx", (z) => edit(z, "ppt/presentation.xml", (x) => x.replace(/<p:extLst>[\s\S]*<\/p:extLst>/, "").replace("</p:presentation>", `${ext}</p:presentation>`)));
    const { pres } = await read(bytes);
    expect(pres.slides.map((s) => s.section?.name)).toEqual(["Giriş", undefined, "Sonuç", undefined]);
  });
});

describe("pptxReader: sunu bölümü", () => {
  it("kısmen okunabilen presentation.xml uyarıyla açılır", async () => {
    const bytes = await mutate("lo-bicim.ppsx", (z) => edit(z, "ppt/presentation.xml", (x) => x.replace(/(<p:presentation[^>]*>)/, '$1<parsererror xmlns="http://www.w3.org/1999/xhtml"/>')));
    const { pres, warnings } = await read(bytes);
    expect(warnings).toContain("Sunu bölümü (presentation.xml) bozuk; yalnız okunabilen kısmı kullanıldı.");
    expect(pres.slides).toHaveLength(4);
  });

  it("olağandışı büyük ya da bozuk presentation.xml anlaşılır hatayla reddedilir", async () => {
    const bozuk = await readPptx(await mutate("lo-bicim.ppsx", (z) => void z.file("ppt/presentation.xml", "<bozuk"))).catch((x) => x);
    expect(bozuk).toBeInstanceOf(PptxError);
    expect(bozuk.message).toBe("Sunu bölümü (presentation.xml) bozuk; dosya açılamadı.");
    ZIP_LIMITS.xml = 64;
    const buyuk = await readPptx(load("lo-bicim.ppsx")).catch((x) => x);
    expect(buyuk).toBeInstanceOf(PptxError);
    expect(buyuk.message).toMatch(/olağandışı büyük/);
  });

  it("belge özellikleri ilişkisi yoksa alışılmış yoldan okunur", async () => {
    const orig = await read(load("lo-bicim.ppsx"));
    const { pres } = await read(await mutate("lo-bicim.ppsx", (z) => edit(z, "_rels/.rels", (x) => x.replace(/<Relationship[^>]*core-properties[^>]*\/>/, ""))));
    expect(pres.props).toEqual(orig.pres.props);
    expect(pres.props).toMatchObject({ title: "Biçim Örneği", creator: "OfficeHere Test" });
  });
});
