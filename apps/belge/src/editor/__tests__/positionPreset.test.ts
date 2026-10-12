import { describe, expect, it } from "vitest";
import { POSITION_PRESETS, nearestVertical, positionPresetId } from "../imageWrap";

describe("Konum önayarları", () => {
  it("dokuz konum; ortadaki sütun üst ve alt, yanlar kare kaydırmayla", () => {
    expect(POSITION_PRESETS.map((p) => p.short)).toEqual(["Sol Üst", "Orta Üst", "Sağ Üst", "Sol Orta", "Orta", "Sağ Orta", "Sol Alt", "Orta Alt", "Sağ Alt"]);
    for (const p of POSITION_PRESETS) {
      expect(p.label).toBe(`${p.short}, ${p.wrapLabel}`);
      expect(p.wrapLabel).toBe(p.h === "center" ? "Üst ve Alt Metin Kaydırma ile" : "Kare Metin Kaydırma ile");
    }
  });

  it("resmin sayfadaki dikey yeri: en yakın kenar ya da orta", () => {
    const content = { top: 100, bottom: 1000 };
    expect(nearestVertical({ top: 104, bottom: 204 }, content)).toBe("top");
    expect(nearestVertical({ top: 550, bottom: 650 }, content)).toBe("middle");
    expect(nearestVertical({ top: 896, bottom: 996 }, content)).toBe("bottom");
    // Büyük resim: alt kenarı içeriğin altına oturduysa üst kenarı ortanın üstünde olsa da "alt".
    expect(nearestVertical({ top: 390, bottom: 996 }, content)).toBe("bottom");
  });

  it("galeride seçili öğe yerleşimden okunur; elle taşınmış resimde seçim yok", () => {
    expect(positionPresetId({ wrap: "inline" }, null)).toBe("inline");
    expect(positionPresetId({}, "top")).toBe("inline");
    expect(positionPresetId({ wrap: "square", floatSide: "right", posX: 0, posY: 0 }, "top")).toBe("top-right");
    expect(positionPresetId({ wrap: "topBottom", floatSide: "center", posX: 0, posY: 0 }, "middle")).toBe("middle-center");
    expect(positionPresetId({ wrap: "square", floatSide: "left", posX: 12, posY: 0 }, "top")).toBeNull();
    expect(positionPresetId({ wrap: "behind", floatSide: "left", posX: 0, posY: 0 }, "top")).toBeNull();
    expect(positionPresetId({ wrap: "tight", floatSide: "left" }, "bottom")).toBeNull();
    expect(positionPresetId({ wrap: "square", floatSide: "left" }, null)).toBeNull();
  });
});
