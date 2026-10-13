import { describe, expect, it } from "vitest";
import { makePicture, makeSp, makeTable, makeTextBox } from "../../model/factory";
import type { ChartShape, GroupShape } from "../../model/types";
import { objectMenuKind } from "./kind";

describe("bağlam menüsü türü", () => {
  const rect = makeSp("rect", 0, 0, 100, 50);
  const pic = makePicture("data:image/png;base64,AA==", 0, 0, 10, 10);

  it("seçim yoksa none, birden çoksa multi", () => {
    expect(objectMenuKind([])).toBe("none");
    expect(objectMenuKind([rect, pic])).toBe("multi");
  });

  it("resim, ortam, grafik, tablo ve grup ayrılır", () => {
    expect(objectMenuKind([pic])).toBe("picture");
    expect(objectMenuKind([{ ...pic, media: { kind: "video", src: "data:video/mp4;base64,", mime: "video/mp4" } }])).toBe("media");
    const chart = { kind: "chart", id: "c", name: "Grafik", x: 0, y: 0, w: 1, h: 1, rot: 0 } as unknown as ChartShape;
    expect(objectMenuKind([chart])).toBe("chart");
    expect(objectMenuKind([makeTable(2, 2, 0, 0, 200)])).toBe("table");
    const group: GroupShape = { kind: "group", id: "g", name: "Grup", x: 0, y: 0, w: 1, h: 1, rot: 0, children: [rect] };
    expect(objectMenuKind([group])).toBe("group");
  });

  it("metin kutusu ve yer tutucu metin menüsü; çizgi, bağlayıcı ve mürekkep çizgi menüsü alır", () => {
    expect(objectMenuKind([makeTextBox(0, 0, 100, 40, "Merhaba")])).toBe("text");
    expect(objectMenuKind([{ ...rect, ph: { type: "title" } }])).toBe("text");
    expect(objectMenuKind([makeSp("line", 0, 0, 100, 0)])).toBe("line");
    expect(objectMenuKind([{ ...rect, connector: true }])).toBe("line");
    expect(objectMenuKind([{ ...rect, geom: "custom", ink: true }])).toBe("line");
    expect(objectMenuKind([rect])).toBe("shape");
  });
});
