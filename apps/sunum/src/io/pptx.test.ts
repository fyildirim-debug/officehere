// @vitest-environment jsdom
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { makePicture, makeSp, makeTable, textBody } from "../model/factory";
import { newPresentation, newSlide } from "../model/presentation";
import { groupShapes } from "../model/ops";
import type { GroupShape, PicShape, Presentation, SpShape, TableShape } from "../model/types";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

// 1×1 kırmızı PNG
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

function build(): Presentation {
  const pres = newPresentation("kiyi", 1);
  const s1 = pres.slides[0];
  s1.shapes = s1.shapes.map((s) => (s.kind === "sp" && s.ph?.type === "ctrTitle" ? { ...s, text: textBody("Merhaba Dünya") } : s));
  s1.notes = "Birinci not satırı\nİkinci satır";
  s1.transition = { type: "push", option: "l", duration: 1.6, advanceOnClick: true, advanceAfter: 4 };

  const s2 = newSlide(pres, "obj");
  s2.shapes = s2.shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph) return s;
    if (s.ph.type === "title") return { ...s, text: textBody("Gündem") };
    return { ...s, text: textBody("Birinci madde\n\tAlt madde\nÇalışma ğüşiöç İ") };
  });
  const rect = makeSp("roundRect", 100, 400, 300, 150, {
    name: "Kutu",
    adj: { adj: 25000 },
    fill: { type: "solid", color: { scheme: "accent2", lumMod: 0.75 } },
    line: { fill: { type: "solid", color: { rgb: "#112233" } }, width: 2.25, dash: "dash" },
    effects: { shadow: { color: { rgb: "#000000", alpha: 0.4 }, blur: 5, dist: 3, dir: 45 } },
    rot: 15,
    text: {
      ...textBody("Kalın metin"),
      anchor: "ctr",
      paragraphs: [{ runs: [{ text: "Kalın ", props: { bold: true, size: 24, color: { rgb: "#FF0000" } } }, { text: "italik", props: { italic: true, underline: true, font: "Georgia" } }], level: 0, props: { align: "ctr", lineSpacing: 150 } }],
    },
  });
  const arrow = makeSp("straightConnector1", 500, 300, 200, 100, { name: "Ok", line: { fill: { type: "solid", color: { scheme: "accent1" } }, width: 3, dash: "solid", tail: "triangle" }, flipV: true, connector: true });
  const pic = makePicture(PNG, 800, 100, 200, 150);
  pic.crop = { l: 0.1, t: 0, r: 0.2, b: 0 };
  const tbl = makeTable(3, 3, 100, 100, 600);
  tbl.rows[0].cells[0].text = textBody("Başlık");
  tbl.rows[1].cells[1].text = textBody("Hücre");
  tbl.rows[2].cells[0].gridSpan = 2;
  tbl.rows[2].cells[1].hMerge = true;
  tbl.style.styleId = "light1-accent3";
  s2.shapes.push(rect, arrow, pic, tbl);
  const ga = makeSp("ellipse", 900, 500, 100, 100, { fill: { type: "solid", color: { scheme: "accent4" } } });
  const gb = makeSp("star5", 1020, 520, 80, 80, { fill: { type: "gradient", kind: "linear", angle: 45, stops: [{ pos: 0, color: { rgb: "#FFFFFF" } }, { pos: 1, color: { scheme: "accent1" } }] } });
  s2.shapes.push(ga, gb);
  s2.shapes = groupShapes(s2.shapes, [ga.id, gb.id]).shapes;
  s2.animations = [
    { id: "a1", shapeId: rect.id, cls: "entr", effect: "flyIn", option: "l", trigger: "click", duration: 0.75, delay: 0 },
    { id: "a2", shapeId: pic.id, cls: "entr", effect: "fade", trigger: "after", duration: 1, delay: 0.5 },
    { id: "a3", shapeId: rect.id, cls: "emph", effect: "spin", trigger: "click", duration: 2, delay: 0 },
  ];
  const s3 = newSlide(pres, "blank");
  s3.hidden = true;
  s3.background = { type: "solid", color: { rgb: "#202020" } };
  pres.slides.push(s2, s3);
  return pres;
}

describe("PPTX gidiş-dönüş", () => {
  it("yaz → oku sonrası içerik korunur", async () => {
    const pres = build();
    const bytes = await writePptx(pres);
    expect(bytes.length).toBeGreaterThan(1000);
    const back = await readPptx(bytes);

    expect(back.width).toBeCloseTo(1280, 0);
    expect(back.height).toBeCloseTo(720, 0);
    expect(back.slides).toHaveLength(3);
    expect(back.theme.id).toBe("kiyi");
    expect(back.theme.variant).toBe(1);
    expect(back.theme.colors.accent1).toBe(pres.theme.colors.accent1);
    expect(back.layouts).toHaveLength(pres.layouts.length);
    expect(back.layouts.map((l) => l.type)).toEqual(pres.layouts.map((l) => l.type));

    // Slayt 1
    const b1 = back.slides[0];
    const t1 = b1.shapes.find((s) => s.kind === "sp" && s.ph?.type === "ctrTitle") as SpShape;
    expect(t1.text!.paragraphs[0].runs[0].text).toBe("Merhaba Dünya");
    expect(b1.notes).toBe("Birinci not satırı\nİkinci satır");
    expect(b1.transition).toMatchObject({ type: "push", option: "l", duration: 1.6, advanceAfter: 4, advanceOnClick: true });
    expect(back.layouts.find((l) => l.id === b1.layoutId)?.type).toBe("title");

    // Slayt 2
    const b2 = back.slides[1];
    const body = b2.shapes.find((s) => s.kind === "sp" && s.ph?.type === "obj") as SpShape;
    expect(body.text!.paragraphs.map((p) => [p.level, p.runs.map((r) => r.text).join("")])).toEqual([
      [0, "Birinci madde"],
      [1, "Alt madde"],
      [0, "Çalışma ğüşiöç İ"],
    ]);
    const kutu = b2.shapes.find((s) => s.name === "Kutu") as SpShape;
    expect(kutu.geom).toBe("roundRect");
    expect(kutu.adj).toEqual({ adj: 25000 });
    expect(kutu.rot).toBeCloseTo(15, 3);
    expect(kutu.x).toBeCloseTo(100, 1);
    expect(kutu.w).toBeCloseTo(300, 1);
    expect(kutu.fill).toEqual({ type: "solid", color: { scheme: "accent2", lumMod: 0.75 } });
    expect(kutu.line).toMatchObject({ width: 2.25, dash: "dash", fill: { type: "solid", color: { rgb: "#112233" } } });
    expect(kutu.effects?.shadow?.dir).toBeCloseTo(45, 3);
    expect(kutu.text!.anchor).toBe("ctr");
    const p0 = kutu.text!.paragraphs[0];
    expect(p0.props).toMatchObject({ align: "ctr", lineSpacing: 150 });
    expect(p0.runs[0]).toMatchObject({ text: "Kalın ", props: { bold: true, size: 24, color: { rgb: "#FF0000" } } });
    expect(p0.runs[1]).toMatchObject({ text: "italik", props: { italic: true, underline: true, font: "Georgia" } });

    const ok = b2.shapes.find((s) => s.name === "Ok") as SpShape;
    expect(ok.geom).toBe("straightConnector1");
    expect(ok.flipV).toBe(true);
    expect(ok.line.tail).toBe("triangle");

    const pic = b2.shapes.find((s) => s.kind === "pic") as PicShape;
    expect(pic.src.startsWith("data:image/png;base64,")).toBe(true);
    expect(pic.crop?.l).toBeCloseTo(0.1, 5);
    expect(pic.crop?.r).toBeCloseTo(0.2, 5);

    const tbl = b2.shapes.find((s) => s.kind === "table") as TableShape;
    expect(tbl.rows).toHaveLength(3);
    expect(tbl.cols).toHaveLength(3);
    expect(tbl.style.styleId).toBe("light1-accent3");
    expect(tbl.style.firstRow).toBe(true);
    expect(tbl.rows[0].cells[0].text.paragraphs[0].runs[0].text).toBe("Başlık");
    expect(tbl.rows[2].cells[0].gridSpan).toBe(2);
    expect(tbl.rows[2].cells[1].hMerge).toBe(true);

    const grp = b2.shapes.find((s) => s.kind === "group") as GroupShape;
    expect(grp.children).toHaveLength(2);
    expect(grp.children[0].x).toBeCloseTo(900, 1);
    expect(grp.children[1].x).toBeCloseTo(1020, 1);
    const star = grp.children[1] as SpShape;
    expect(star.fill.type).toBe("gradient");

    expect(b2.animations.map((a) => [a.effect, a.option, a.trigger, a.duration, a.delay])).toEqual([
      ["flyIn", "l", "click", 0.75, 0],
      ["fade", undefined, "after", 1, 0.5],
      ["spin", undefined, "click", 2, 0],
    ]);
    expect(b2.animations[0].shapeId).toBe(kutu.id);
    expect(b2.animations[1].shapeId).toBe(pic.id);

    // Slayt 3
    expect(back.slides[2].hidden).toBe(true);
    expect(back.slides[2].background).toEqual({ type: "solid", color: { rgb: "#202020" } });
  });

  it("paket yapısı ve içerik türleri geçerli", async () => {
    const bytes = await writePptx(build());
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files);
    expect(names[0]).toBe("[Content_Types].xml");
    for (const p of [
      "ppt/presentation.xml",
      "ppt/slides/slide1.xml",
      "ppt/slideMasters/slideMaster1.xml",
      "ppt/slideLayouts/slideLayout1.xml",
      "ppt/theme/theme1.xml",
      "ppt/notesSlides/notesSlide1.xml",
      "ppt/notesMasters/notesMaster1.xml",
      "ppt/tableStyles.xml",
      "docProps/core.xml",
      "ppt/media/image1.png",
    ])
      expect(names).toContain(p);
    const ct = await zip.file("[Content_Types].xml")!.async("string");
    expect(ct).toContain('Extension="png"');
    expect(ct).toContain("/ppt/slides/slide3.xml");
    // Tüm XML parçaları iyi biçimli olmalı.
    for (const n of names.filter((x) => x.endsWith(".xml") || x.endsWith(".rels"))) {
      const xml = await zip.file(n)!.async("string");
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      expect(doc.getElementsByTagName("parsererror").length, n).toBe(0);
    }
  });

  it("isteğe bağlı: örnek dosyayı diske yaz (OH_DUMP)", async () => {
    const env = (globalThis as unknown as { process?: { env: Record<string, string | undefined> } }).process?.env;
    const out = env?.OH_DUMP;
    if (!out) return;
    const fs = (await import(/* @vite-ignore */ "node:" + "fs")) as { writeFileSync: (p: string, d: Uint8Array) => void };
    fs.writeFileSync(out, await writePptx(build()));
  });

  it("mürekkep çizimi (serbest yol, yuvarlak uç, saydamlık) korunur", async () => {
    const pres = build();
    const ink = makeSp("custom", 100, 120, 200, 80, {
      name: "Mürekkep 1",
      fill: { type: "none" },
      line: { fill: { type: "solid", color: { rgb: "#FFFF00", alpha: 0.5 } }, width: 12, dash: "solid", cap: "rnd" },
      paths: [{ w: 200, h: 80, fill: false, cmds: [{ c: "M", pts: [0, 40] }, { c: "Q", pts: [50, 0, 100, 40] }, { c: "Q", pts: [150, 80, 200, 40] }] }],
    }) as SpShape;
    ink.ink = true;
    pres.slides[0].shapes.push(ink);
    const back = await readPptx(await writePptx(pres));
    const r = back.slides[0].shapes.find((x) => x.name === "Mürekkep 1") as SpShape;
    expect(r.ink).toBe(true);
    expect(r.geom).toBe("custom");
    expect(r.line.cap).toBe("rnd");
    expect(r.line.width).toBeCloseTo(12, 1);
    expect(r.line.fill.type === "solid" && r.line.fill.color.alpha).toBeCloseTo(0.5, 2);
    expect(r.paths![0].cmds.map((c) => c.c)).toEqual(["M", "Q", "Q"]);
    expect(r.paths![0].fill).toBe(false);
  });

  it("açıklamalar (yazar, metin, konum) korunur", async () => {
    const pres = build();
    pres.slides[1].comments = [
      { id: "c1", author: "Ayşe Yılmaz", text: "Bu başlığı kısaltalım & <netleştirelim>", x: 640, y: 120, date: Date.parse("2026-10-01T09:30:00Z") },
      { id: "c2", author: "Mehmet Öz", text: "Katılıyorum.", x: 20, y: 20, date: Date.parse("2026-10-01T10:00:00Z") },
    ];
    pres.slides[2].comments = [{ id: "c3", author: "Ayşe Yılmaz", text: "Tablo güncel mi?", x: 300, y: 400, date: Date.parse("2026-10-02T08:00:00Z") }];
    const bytes = await writePptx(pres);
    const zip = await JSZip.loadAsync(bytes);
    expect(Object.keys(zip.files)).toContain("ppt/commentAuthors.xml");
    expect(Object.keys(zip.files)).toContain("ppt/comments/comment2.xml");
    const authors = await zip.file("ppt/commentAuthors.xml")!.async("string");
    expect(authors).toContain('name="Ayşe Yılmaz"');
    expect(authors).toContain('lastIdx="2"');
    const back = await readPptx(bytes);
    expect(back.slides[0].comments).toBeUndefined();
    const c = back.slides[1].comments!;
    expect(c).toHaveLength(2);
    expect(c[0]).toMatchObject({ author: "Ayşe Yılmaz", text: "Bu başlığı kısaltalım & <netleştirelim>", x: 640, y: 120 });
    expect(c[0].date).toBe(Date.parse("2026-10-01T09:30:00Z"));
    expect(c[1].author).toBe("Mehmet Öz");
    expect(back.slides[2].comments![0].text).toBe("Tablo güncel mi?");
  });

  it("4:3 slayt boyutu korunur", async () => {
    const pres = newPresentation("varsayilan", 0, { width: 960, height: 720 });
    const back = await readPptx(await writePptx(pres));
    expect(back.width).toBeCloseTo(960, 0);
    expect(back.height).toBeCloseTo(720, 0);
  });

  it("geçersiz dosyada anlaşılır hata", async () => {
    await expect(readPptx(new Uint8Array([1, 2, 3]))).rejects.toThrow(/PPTX/);
  });
});
