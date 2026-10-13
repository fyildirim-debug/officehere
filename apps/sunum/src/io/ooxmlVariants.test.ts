// @vitest-environment jsdom
/** OOXML sunu çeşitleri: .pptm, .potx, .potm, .ppsx, .ppsm (içerik türleri farklı, ana bölüm aynı). */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { plainText } from "../model/textStyle";
import type { Presentation } from "../model/types";
import { TEMPLATES } from "../templates/templates";
import { FORMATS, readPresentationFile, saveFilters, sniffFormat, writePresentationFormat, type SunuFormat } from "./formats";
import { MAIN_CT, writePptx } from "./pptxWriter";

const MACRO_CT: Record<string, string> = {
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  ppsm: "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
  potm: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
  potx: MAIN_CT.template,
  ppsx: MAIN_CT.slideshow,
};

const textOf = (p: Presentation) => p.slides.map((s) => s.shapes.map((sh) => (sh.kind === "sp" ? plainText(sh.text) : "")).join(" ")).join(" ");

/** Yazılan .pptx'i başka bir çeşide çevirir (makrolu çeşitlere VBA parçası eklenir). */
async function variant(bytes: Uint8Array, ext: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  let ct = await zip.file("[Content_Types].xml")!.async("string");
  ct = ct.replace(MAIN_CT.presentation, MACRO_CT[ext]);
  if (ext.endsWith("m")) {
    ct = ct.replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
    zip.file("ppt/vbaProject.bin", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]));
    let rels = await zip.file("ppt/_rels/presentation.xml.rels")!.async("string");
    rels = rels.replace("</Relationships>", '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
    zip.file("ppt/_rels/presentation.xml.rels", rels);
  }
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "uint8array" });
}

describe("OOXML çeşitleri", () => {
  it.each(["pptm", "potx", "potm", "ppsx", "ppsm"] as SunuFormat[])("%s içerik türüyle tanınır ve açılır (makrolar yok sayılır)", async (ext) => {
    const src = TEMPLATES.find((t) => t.id === "project")!.build();
    const bytes = await variant(await writePptx(src), ext);
    // Uzantı yanlış olsa da içerik türü belirleyicidir.
    expect(await sniffFormat(bytes, "dosya.pptx")).toBe(ext);
    const r = await readPresentationFile(bytes, `Örnek Şablon.${ext}`);
    expect(r.format).toBe(ext);
    expect(r.pres.slides).toHaveLength(src.slides.length);
    expect(textOf(r.pres)).toBe(textOf(src));
    const info = FORMATS[ext];
    expect(!!info.template).toBe(ext.startsWith("pot"));
    expect(!!info.show).toBe(ext.startsWith("pps"));
    expect(!!info.macro).toBe(ext.endsWith("m"));
  });

  it("LibreOffice'in yazdığı .ppsx açılır", async () => {
    const data = new Uint8Array(readFileSync(`${process.cwd()}/src/io/__fixtures__/lo-bicim.ppsx`));
    const r = await readPresentationFile(data, "lo-bicim.ppsx");
    expect(r.format).toBe("ppsx");
    expect(r.pres.slides).toHaveLength(4);
    expect(textOf(r.pres)).toContain("Çeyrek Sonuçları — Iğdır Şubesi");
    // Not sayfasında yer tutucu yoksa ilk metin kutusu not olur.
    expect(r.pres.slides[0].notes).toBe("Açılış: herkese hoş geldiniz deyin.");
  });

  it("Farklı Kaydet: .ppsx ve .potx ana içerik türüyle yazılır; .odp yazılır", async () => {
    const src = TEMPLATES.find((t) => t.id === "meeting")!.build();
    for (const [fmt, ct] of [
      ["ppsx", MAIN_CT.slideshow],
      ["potx", MAIN_CT.template],
      ["pptx", MAIN_CT.presentation],
    ] as const) {
      const bytes = await writePresentationFormat(src, fmt);
      const types = await (await JSZip.loadAsync(bytes)).file("[Content_Types].xml")!.async("string");
      expect(types).toContain(`PartName="/ppt/presentation.xml" ContentType="${ct}"`);
      expect(await sniffFormat(bytes, `x.${fmt}`)).toBe(fmt);
      const back = await readPresentationFile(bytes, `x.${fmt}`);
      expect(textOf(back.pres)).toBe(textOf(src));
    }
    const odp = await writePresentationFormat(src, "odp");
    expect(await sniffFormat(odp, "x.odp")).toBe("odp");
    expect(saveFilters("odp")[0]).toEqual({ name: "OpenDocument Sunu", extensions: ["odp"] });
    expect(saveFilters("pptx").map((f) => f.extensions[0])).toEqual(["pptx", "odp", "ppsx", "potx"]);
  });

  it("kelime işlemci ve hesap tablosu paketleri sunu diye açılmaz", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    await expect(sniffFormat(await zip.generateAsync({ type: "uint8array" }), "a.pptx")).rejects.toThrow(/kelime işlemci/);
    await expect(sniffFormat(new TextEncoder().encode("düz metin"), "a.pptx")).rejects.toThrow(/tanınmadı/);
  });

  it("slayt adı .pptx gidiş-dönüşünde korunur", async () => {
    const src = TEMPLATES.find((t) => t.id === "blank")!.build();
    src.slides[0] = { ...src.slides[0], name: "Açılış Slaytı" };
    const back = await readPresentationFile(await writePptx(src), "x.pptx");
    expect(back.pres.slides[0].name).toBe("Açılış Slaytı");
  });
});
