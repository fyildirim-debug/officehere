// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readPptx } from "../io/pptxReader";
import { writePptx } from "../io/pptxWriter";
import { TEMPLATES } from "./templates";
import { THEMES } from "../themes/themes";
import { applyTheme, newPresentation } from "../model/presentation";

describe("Şablonlar", () => {
  it("en az 4 hazır şablon + boş sunu", () => {
    expect(TEMPLATES.length).toBeGreaterThanOrEqual(5);
    expect(TEMPLATES[0].id).toBe("blank");
  });

  for (const t of TEMPLATES) {
    it(`${t.label}: oluşturulur ve PPTX gidiş-dönüşünde slayt/şekil sayısı korunur`, async () => {
      const pres = t.build();
      expect(pres.slides.length).toBeGreaterThan(0);
      const back = await readPptx(await writePptx(pres));
      expect(back.slides.length).toBe(pres.slides.length);
      back.slides.forEach((s, i) => {
        expect(s.shapes.length).toBe(pres.slides[i].shapes.length);
        expect(s.animations.length).toBe(pres.slides[i].animations.length);
      });
      const env = (globalThis as unknown as { process?: { env: Record<string, string | undefined> } }).process?.env;
      if (env?.OH_DUMP_DIR) {
        const fs = (await import(/* @vite-ignore */ "node:" + "fs")) as { writeFileSync: (p: string, d: Uint8Array) => void };
        fs.writeFileSync(`${env.OH_DUMP_DIR}/${t.id}.pptx`, await writePptx(pres));
      }
    });
  }
});

describe("Temalar", () => {
  it("en az 8 tema, her biri 4 çeşitleme", () => {
    expect(THEMES.length).toBeGreaterThanOrEqual(8);
    for (const t of THEMES) expect(t.variants.length).toBe(4);
  });

  it("tema değişince tüm slaytlar yeni düzenlere bağlanır", () => {
    const p = TEMPLATES[1].build();
    const q = applyTheme(p, "orman", 2);
    expect(q.theme.id).toBe("orman");
    expect(q.theme.colors.accent1).toBe(THEMES.find((x) => x.id === "orman")!.variants[2].colors.accent1);
    for (const s of q.slides) expect(q.layouts.some((l) => l.id === s.layoutId)).toBe(true);
    // Yer tutucu metinleri korunur.
    const t0 = q.slides[0].shapes.find((s) => s.kind === "sp" && s.ph?.type === "ctrTitle");
    expect(t0 && t0.kind === "sp" && t0.text?.paragraphs[0].runs[0].text).toBe("Müşteri Portalı Yenileme Projesi");
  });

  it("boş sunu tek başlık slaytıyla başlar", () => {
    const p = newPresentation();
    expect(p.slides).toHaveLength(1);
    expect(p.layouts.find((l) => l.id === p.slides[0].layoutId)?.type).toBe("title");
  });
});
