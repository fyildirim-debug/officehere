// @vitest-environment jsdom
/**
 * ODF geçiş katmanı (Sunum): açılan .odp'nin özgün paketi kaynak kaydına girer (`pres.kaynak`); .odp'ye kaydederken
 * settings.xml taşınır, Thumbnails/ düşer, gömülü grafikler (Object N/ alt belgeleri ve yedek görüntüleri) özgün
 * çerçeveleriyle yazılır. Görünümü düzenlenen ya da kaynağı bulunamayan nesne yedek görünümüyle (resim) yazılır.
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { parseManifest } from "@officehere/core/odf";
import { getSource } from "@officehere/core/opc";
import { flatShapes } from "../../model/ops";
import type { Presentation, Shape } from "../../model/types";
import { readPresentationFile, writePresentationFormat } from "../formats";
import { kaynakBirak } from "../rawPassthrough";
import { ODF_OBJECTS_EDITED, ODF_OBJECTS_FALLBACK, ODF_OBJECTS_SOURCE } from "./objects";

const LO = (n: string) => new Uint8Array(readFileSync(`${process.cwd()}/../../test-derlem/lo/${n}`));
const objects = (p: Presentation) => p.slides.flatMap((s) => flatShapes(s.shapes)).filter((s) => s.odf);

async function pkg(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const manifest = parseManifest(await zip.file("META-INF/manifest.xml")!.async("string"));
  return { zip, names, manifest, content: await zip.file("content.xml")!.async("string") };
}

async function write(p: Presentation, format: "odp" | "pptx" = "odp") {
  const warnings: string[] = [];
  const bytes = await writePresentationFormat(p, format, { warnings });
  return { bytes, warnings };
}

describe("ODP geçiş katmanı", () => {
  it("okurken: kaynak kaydı, gömülü grafikler yedek görünümlü resim (özgün çerçeveyle)", async () => {
    const { pres } = await readPresentationFile(LO("lo-sunu24.odp"), "lo-sunu24.odp");
    expect(getSource(pres.kaynak)).toBeTruthy();
    const objs = objects(pres);
    expect(objs).toHaveLength(3);
    for (const o of objs) {
      expect(o.kind).toBe("pic");
      expect(o.odf!.src).toBe(pres.kaynak);
      expect(o.odf!.xml).toMatch(/<draw:frame [^>]*>.*xlink:href="\.\/Object \d"/s);
      expect(o.odf!.styles.length).toBeGreaterThan(0);
    }
  });

  it(".odp'ye kaydederken alt belgeler ve ayarlar taşınır, Thumbnails düşer, manifest tam; ikinci tur kararlı", async () => {
    const { pres } = await readPresentationFile(LO("lo-sunu24.odp"), "lo-sunu24.odp");
    const { bytes, warnings } = await write(pres);
    expect(warnings).toEqual([]);
    const out = await pkg(bytes);
    for (const n of [1, 2, 3]) {
      expect(out.names).toContain(`Object ${n}/content.xml`);
      expect(out.names).toContain(`ObjectReplacements/Object ${n}`);
      expect(out.manifest.find((e) => e.path === `Object ${n}/`)?.mediaType).toBe("application/vnd.oasis.opendocument.chart");
    }
    expect(out.names).toContain("settings.xml");
    expect(out.names.some((n) => n.startsWith("Thumbnails/"))).toBe(false);
    const listed = new Set(out.manifest.map((e) => e.path));
    for (const n of out.names) if (n !== "mimetype" && n !== "META-INF/manifest.xml") expect(listed).toContain(n);
    expect(out.content).toMatch(/<draw:frame [^>]*draw:style-name="oh\d+-gr\d+"/);
    const again = (await readPresentationFile(bytes, "kayit.odp")).pres;
    expect(objects(again)).toHaveLength(3);
    const out2 = await pkg((await write(again)).bytes);
    expect(out2.names.sort()).toEqual(out.names.sort());
  });

  it("taşınan ve boyutlandırılan nesne yeni konumuyla yazılır; çoğaltılan kendi kopyasını alır", async () => {
    const { pres } = await readPresentationFile(LO("lo-sunu.odp"), "lo-sunu.odp");
    const [o] = objects(pres);
    o.x += 96;
    o.w = 480;
    const slide = pres.slides.find((s) => s.shapes.includes(o))!;
    slide.shapes.push({ ...structuredClone(o), id: "kopya" } as Shape);
    const out = await pkg((await write(pres)).bytes);
    expect(out.names).toContain("Object 1/content.xml");
    expect(out.names).toContain("Object 2/content.xml");
    expect(out.content).toContain('svg:width="12.7cm"');
  });

  it("görünümü düzenlenen ya da kaynağı olmayan nesne resim olarak yazılır ve uyarılır", async () => {
    const { pres } = await readPresentationFile(LO("lo-sunu.odp"), "lo-sunu.odp");
    const [o] = objects(pres);
    if (o.kind !== "pic") throw new Error("resim bekleniyordu");
    o.src = "data:image/png;base64,iVBORw0KGgo=";
    let r = await write(pres);
    expect(r.warnings).toContain(ODF_OBJECTS_EDITED);
    expect((await pkg(r.bytes)).names.some((n) => n.startsWith("Object"))).toBe(false);
    const fresh = (await readPresentationFile(LO("lo-sunu.odp"), "lo-sunu.odp")).pres;
    kaynakBirak(fresh.kaynak);
    r = await write(fresh);
    expect(r.warnings).toContain(ODF_OBJECTS_SOURCE);
    expect((await pkg(r.bytes)).names).not.toContain("settings.xml");
  });

  it(".pptx'e kaydederken yedek görünümü yazılır ve uyarılır; kaynak kaydı kapalıyken nesne taşınmaz", async () => {
    const { pres } = await readPresentationFile(LO("lo-sunu.odp"), "lo-sunu.odp");
    const r = await write(pres, "pptx");
    expect(r.warnings).toContain(ODF_OBJECTS_FALLBACK);
    expect((await JSZip.loadAsync(r.bytes)).file("ppt/presentation.xml")).not.toBeNull();
    const off = (await readPresentationFile(LO("lo-sunu.odp"), "lo-sunu.odp", { kaynak: false })).pres;
    expect(off.kaynak).toBeUndefined();
    expect(objects(off)).toHaveLength(0);
  });
});
