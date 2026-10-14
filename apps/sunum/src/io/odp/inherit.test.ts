// @vitest-environment jsdom
/**
 * Devralınan metin biçimi: yazıcılar yalnız açıkça ayarlanmış çalıştırma
 * özelliklerini yazar; devralınan değerler (tema/asıl slayt) çalıştırmalara
 * sabitlenmez. Gidiş-dönüşte açık özellik sayısı ve görünüm korunur; tema
 * değişince metin rengi de değişir.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import { resolveHex } from "../../model/color";
import { flatShapes } from "../../model/ops";
import { applyTheme } from "../../model/presentation";
import { levelResolver, resolveFont, effectiveRun } from "../../model/textStyle";
import type { Presentation, SpShape } from "../../model/types";
import { TEMPLATES } from "../../templates/templates";
import { THEMES } from "../../themes/themes";
import { readPptx } from "../pptxReader";
import { writePptx } from "../pptxWriter";
import { readOdp } from "./reader";
import { writeOdp } from "./writer";

// Paket ayrıştırma ağır: yüklü makinede (paralel CI işleri) 5 sn varsayılan yetmeyebilir.
vi.setConfig({ testTimeout: 30_000 });

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const load = (n: string) => new Uint8Array(readFileSync(`${FIX}/${n}`));

/** Karakter ağırlıklı açık çalıştırma özelliği sayıları (ölçüm aracındaki "renk/boyut/yazıtipi"). */
function explicitCounts(p: Presentation, tables = true) {
  const c = { renk: 0, boyut: 0, yazıtipi: 0, kalın: 0, italik: 0 };
  for (const s of p.slides)
    for (const sh of flatShapes(s.shapes)) {
      const bodies = sh.kind === "table" ? (tables ? sh.rows.flatMap((r) => r.cells.map((x) => x.text)) : []) : sh.kind === "sp" && sh.text ? [sh.text] : [];
      for (const b of bodies)
        for (const para of b?.paragraphs ?? [])
          for (const r of para.runs) {
            if (r.props.color) c.renk += r.text.length;
            if (r.props.size) c.boyut += r.text.length;
            if (r.props.font) c.yazıtipi += r.text.length;
            if (r.props.bold) c.kalın += r.text.length;
            if (r.props.italic) c.italik += r.text.length;
          }
    }
  return c;
}

/** Yer tutucu/metin kutusu çalıştırmalarının etkin görünümü (renk onaltılık, yazı tipi adı). */
function effective(p: Presentation): string[] {
  const out: string[] = [];
  for (const s of p.slides) {
    const layout = p.layouts.find((l) => l.id === s.layoutId);
    for (const sh of flatShapes(s.shapes)) {
      if (sh.kind !== "sp" || !sh.text) continue;
      const lvl = levelResolver(sh, { master: p.master, layout });
      for (const para of sh.text.paragraphs)
        for (const r of para.runs) {
          if (!r.text.trim()) continue;
          const e = effectiveRun(r.props, para, lvl);
          out.push(`${r.text}|${e.size}|${!!e.bold}|${!!e.italic}|${resolveHex(e.color, p.theme)}|${resolveFont(e.font, p.theme)}`);
        }
    }
  }
  return out;
}

function titleColor(p: Presentation): string | undefined {
  for (const s of p.slides) {
    const layout = p.layouts.find((l) => l.id === s.layoutId);
    const t = s.shapes.find((x): x is SpShape => x.kind === "sp" && (x.ph?.type === "title" || x.ph?.type === "ctrTitle") && !!x.text?.paragraphs.some((q) => q.runs.some((r) => r.text)));
    if (!t) continue;
    const para = t.text!.paragraphs.find((q) => q.runs.some((r) => r.text))!;
    return resolveHex(effectiveRun(para.runs[0].props, para, levelResolver(t, { master: p.master, layout })).color, p.theme);
  }
  return undefined;
}

describe("ODP yazıcısı devralınan biçimi çalıştırmalara sabitlemez", () => {
  it("LibreOffice ODP gidiş-dönüşünde açık özellik sayısı ve etkin görünüm korunur", async () => {
    const a = await readOdp(load("lo-bicim.odp"));
    const b = await readOdp(await writeOdp(a));
    expect(explicitCounts(b)).toEqual(explicitCounts(a));
    expect(effective(b)).toEqual(effective(a));
    // İkinci tur da kararlı.
    const c = await readOdp(await writeOdp(b));
    expect(explicitCounts(c)).toEqual(explicitCounts(a));
  });

  it("uygulama şablonları: ODP gidiş-dönüşünde açık özellik eklenmez, görünüm korunur", async () => {
    for (const t of TEMPLATES) {
      const a = t.build();
      const b = await readOdp(await writeOdp(a));
      // Tablo stilleri ODF'de yok: tablo stilinden gelen hücre biçimi açık yazılır (görünüm korunur).
      expect(explicitCounts(b, false), t.id).toEqual(explicitCounts(a, false));
      expect(effective(b), t.id).toEqual(effective(a));
    }
  });

  it("ODP'den açılan sunuda tema değişince metin rengi de değişir", async () => {
    const a = TEMPLATES.find((t) => t.id === "project")!.build();
    const b = await readOdp(await writeOdp(a));
    expect(b.theme.colors).toEqual(a.theme.colors);
    const before = titleColor(b);
    expect(before).toBe(titleColor(a));
    const other = THEMES.map((x) => x.id).find((id) => titleColor(applyTheme(a, id)) !== before)!;
    expect(other).toBeDefined();
    const themed = applyTheme(b, other);
    expect(titleColor(themed)).not.toBe(before);
    // Yerel sunuda da aynı renge geçer (ODP'den gelen metin temayı izler).
    expect(titleColor(themed)).toBe(titleColor(applyTheme(a, other)));
  });
});

describe("PPTX yazıcısı da yalnız açık özellikleri yazar", () => {
  it("şablonlar ve LibreOffice PPTX'i: gidiş-dönüşte açık özellik sayısı korunur", async () => {
    for (const t of TEMPLATES) {
      const a = t.build();
      const b = await readPptx(await writePptx(a));
      expect(explicitCounts(b), t.id).toEqual(explicitCounts(a));
      for (const th of ["orman", "kiyi"]) expect(titleColor(applyTheme(b, th)), t.id).toBe(titleColor(applyTheme(a, th)));
    }
    const a = await readPptx(load("lo-proje.pptx"));
    const b = await readPptx(await writePptx(a));
    expect(explicitCounts(b)).toEqual(explicitCounts(a));
  });
});
