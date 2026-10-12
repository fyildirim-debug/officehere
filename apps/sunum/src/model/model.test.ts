import { describe, expect, it } from "vitest";
import { textBody } from "./factory";
import { bodyShape, bodyToLines, linesToBody, titleShape } from "./outline";
import { newPresentation, newSlide, resizePresentation } from "./presentation";
import { findPlaceholder } from "./textStyle";
import type { SpShape } from "./types";

describe("Ana hat dönüşümü", () => {
  it("düzeyler sekmeyle gösterilir ve geri okunur", () => {
    const body = textBody("Birinci\n\tAlt madde\n\t\tDaha alt\nİkinci");
    const lines = bodyToLines(body);
    expect(lines).toBe("Birinci\n\tAlt madde\n\t\tDaha alt\nİkinci");
    const back = linesToBody(lines, body);
    expect(back.paragraphs.map((p) => p.level)).toEqual([0, 1, 2, 0]);
    expect(back.paragraphs[1].runs[0].text).toBe("Alt madde");
  });

  it("paragraf ve karakter biçimi satır sırasına göre korunur", () => {
    const body = textBody("Kalın satır\nİkinci");
    body.paragraphs[0].runs[0].props = { bold: true, size: 32 };
    body.paragraphs[0].props = { align: "ctr" };
    const next = linesToBody("Kalın satır (düzenlendi)\nİkinci\n\tYeni", body);
    expect(next.paragraphs[0].runs[0].props).toMatchObject({ bold: true, size: 32 });
    expect(next.paragraphs[0].props.align).toBe("ctr");
    expect(next.paragraphs[2]).toMatchObject({ level: 1, runs: [{ text: "Yeni" }] });
  });

  it("boş gövde boş metin verir; başlık ve gövde yer tutucuları bulunur", () => {
    const pres = newPresentation();
    const s = newSlide(pres, "obj");
    expect(bodyToLines(bodyShape(s)?.text)).toBe("");
    expect(titleShape(s)?.ph?.type).toBe("title");
    expect(bodyShape(s)?.ph?.type).toBe("obj");
  });
});

describe("Slayt boyutu değişimi", () => {
  it("düzenden konum alan yer tutucular yeni düzene oturur, özel konumlular ölçeklenir", () => {
    const pres = newPresentation("kiyi", 0);
    const s = newSlide(pres, "obj");
    pres.slides.push(s);
    const moved = s.shapes.find((x) => x.kind === "sp" && x.ph?.type === "obj") as SpShape;
    moved.x -= 40; // kullanıcı taşımış
    const q = resizePresentation(pres, 960, 720);
    expect(q.width).toBe(960);
    const qs = q.slides[1];
    const layout = q.layouts.find((l) => l.id === qs.layoutId)!;
    const title = qs.shapes.find((x) => x.kind === "sp" && x.ph?.type === "title") as SpShape;
    const lt = findPlaceholder(layout.shapes, title.ph!)!;
    expect([title.x, title.y, title.w, title.h]).toEqual([lt.x, lt.y, lt.w, lt.h]);
    const body = qs.shapes.find((x) => x.kind === "sp" && x.ph?.type === "obj") as SpShape;
    // 0,75 oranında ölçeklenip dikeyde ortalanır.
    expect(body.x).toBeCloseTo(moved.x * 0.75, 3);
    expect(body.y).toBeCloseTo(moved.y * 0.75 + 90, 3);
    // Tüm şekiller slayt içinde kalır.
    for (const sh of qs.shapes) expect(sh.x + sh.w).toBeLessThanOrEqual(960 + 0.01);
  });
});
