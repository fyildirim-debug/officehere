// @vitest-environment jsdom
// ODF geçiş katmanı (Belge): açılan .odt'nin özgün paketi kaynak kaydına girer; .odt'ye kaydederken settings.xml ve
// manifest.rdf taşınır, Thumbnails/ düşer, gömülü grafik alt belgeleri (Object N/, ObjectReplacements/) opak öğe
// olarak korunur ve manifest bunlarla yeniden kurulur. Başka biçimde ve kaynaksız yazımda yedek görünüm yazılır.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { getSource, releaseSource } from "@officehere/core/opc";
import { parseManifest } from "@officehere/core/odf";
import { readOdt, validateContent, writeDocx, writeOdt } from "../index";
import { LoadWarnings } from "../warnings";
import { all, fixture, text } from "./bicim-yardimci";

async function open(name: string, opts?: { kaynak?: boolean }) {
  const warnings = new LoadWarnings();
  const doc = await readOdt(fixture(name), warnings, opts);
  return { doc, warnings };
}

const objects = (c: JSONContent) => all(c, "opaqueInline").filter((n) => n.attrs?.odf);

async function pkg(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const manifest = parseManifest(await zip.file("META-INF/manifest.xml")!.async("string"));
  return { zip, names, manifest, content: await zip.file("content.xml")!.async("string") };
}

describe("ODT geçiş katmanı: paket düzeyi", () => {
  it("özgün paket kaydedilir; aynı biçimde settings.xml ve manifest.rdf taşınır, Thumbnails düşer, manifest tam", async () => {
    const { doc } = await open("lo-ornek.odt");
    expect(doc.settings.kaynak).toMatch(/^[0-9a-f]{64}$/);
    expect(getSource(doc.settings.kaynak)).toBeTruthy();
    const out = await pkg(await writeOdt(doc));
    expect(out.names[0]).toBe("mimetype");
    expect(out.names).toContain("settings.xml");
    expect(out.names).toContain("manifest.rdf");
    expect(out.names.some((n) => n.startsWith("Thumbnails/"))).toBe(false);
    const listed = new Set(out.manifest.map((e) => e.path));
    for (const n of out.names) if (n !== "mimetype" && n !== "META-INF/manifest.xml") expect(listed).toContain(n);
    const src = await JSZip.loadAsync(fixture("lo-ornek.odt"));
    expect(await out.zip.file("settings.xml")!.async("string")).toBe(await src.file("settings.xml")!.async("string"));
  });

  it("kaynak bırakılınca ya da başka biçimden (.docx) gelen belgede taşınmaz", async () => {
    const { doc } = await open("lo-ornek.odt");
    releaseSource(doc.settings.kaynak);
    expect((await pkg(await writeOdt(doc))).names).not.toContain("settings.xml");
    const { doc: free } = await open("lo-ornek.odt", { kaynak: false });
    expect(free.settings.kaynak).toBeFalsy();
    expect((await pkg(await writeOdt(free))).names).not.toContain("manifest.rdf");
  });
});

describe("ODT geçiş katmanı: gömülü nesneler", () => {
  it("grafik alt belgesi opak öğe olarak okunur (kayıp uyarısı yok)", async () => {
    const { doc, warnings } = await open("lo-nesne.odt");
    validateContent(doc.content);
    const [o, ...rest] = objects(doc.content);
    expect(rest).toHaveLength(0);
    expect(o.attrs).toMatchObject({ kind: "ole", label: "Gömülü nesne (grafik)", src: doc.settings.kaynak, xml: "" });
    expect(o.attrs!.odf.xml).toContain('xlink:href="./Object 1"');
    expect(o.attrs!.odf.styles.join("")).toContain('style:name="fr1"');
    expect(warnings.has("objects")).toBe(false);
    expect(text(doc.content)).toContain("Gömülü grafik ve formül içeren belge.");
  });

  it(".odt'ye kaydederken alt belge, yedek görüntü ve stili korunur; ikinci tur kararlı", async () => {
    const { doc } = await open("lo-nesne.odt");
    const bytes = await writeOdt(doc);
    const out = await pkg(bytes);
    for (const n of ["Object 1/content.xml", "Object 1/styles.xml", "Object 1/meta.xml", "ObjectReplacements/Object 1", "settings.xml", "manifest.rdf"]) expect(out.names).toContain(n);
    const man = new Map(out.manifest.map((e) => [e.path, e.mediaType]));
    expect(man.get("Object 1/")).toBe("application/vnd.oasis.opendocument.chart");
    expect(man.get("ObjectReplacements/Object 1")).toContain("gdimetafile");
    expect(out.content).toMatch(/<draw:frame [^>]*draw:style-name="(oh\d+-fr1)"/);
    const style = /draw:style-name="(oh\d+-fr1)"/.exec(out.content)![1];
    expect(out.content).toContain(`style:name="${style}"`);
    const src = await JSZip.loadAsync(fixture("lo-nesne.odt"));
    expect(await out.zip.file("Object 1/content.xml")!.async("string")).toBe(await src.file("Object 1/content.xml")!.async("string"));
    // Yeniden açılır ve aynı biçimde bir daha kaydedilince aynı paket yapısı.
    const again = await readOdt(bytes);
    expect(objects(again.content)).toHaveLength(1);
    const out2 = await pkg(await writeOdt(again));
    expect(out2.names.sort()).toEqual(out.names.sort());
    expect(out2.manifest.length).toBe(out.manifest.length);
  });

  it("çoğaltılan nesne kendi kopyasını alır (Object 2)", async () => {
    const { doc } = await open("lo-nesne.odt");
    const para = all(doc.content, "paragraph").find((p) => p.content?.some((c) => c.type === "opaqueInline"))!;
    const node = para.content!.find((c) => c.type === "opaqueInline")!;
    para.content!.push(structuredClone(node));
    const out = await pkg(await writeOdt(doc));
    expect(out.names).toContain("Object 2/content.xml");
    expect(out.names).toContain("ObjectReplacements/Object 2");
    expect(out.content).toContain('xlink:href="./Object 2"');
  });

  it("kaynak yoksa yer tutucu yazılır ve bildirilir; .docx'e yedek görünümü yazılır", async () => {
    const { doc } = await open("lo-nesne.odt");
    releaseSource(doc.settings.kaynak);
    let lost = false;
    const out = await pkg(await writeOdt(doc, { onReport: (r) => (lost = r.objectsLost) }));
    expect(lost).toBe(true);
    expect(out.names.some((n) => n.startsWith("Object"))).toBe(false);
    expect(out.content).toContain("[Desteklenmeyen öğe: Gömülü nesne (grafik)]");
    const { doc: d2 } = await open("lo-nesne.odt");
    const docx = await JSZip.loadAsync(await writeDocx(d2));
    expect(await docx.file("word/document.xml")!.async("string")).toContain("Desteklenmeyen öğe: Gömülü nesne (grafik)");
  });

  it("kaynak kaydı kapalıyken nesne kaydedince kaybolacağı için uyarılır", async () => {
    const { doc, warnings } = await open("lo-nesne.odt", { kaynak: false });
    expect(objects(doc.content)).toHaveLength(0);
    expect(warnings.has("objects")).toBe(true);
  });
});
