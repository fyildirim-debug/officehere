import { describe, expect, it } from "vitest";
import { addRecentEntry, normalizeRecent, pinRecentEntry, removeRecentEntry, type RecentEntry } from "../src/main/recentList";

const f = (name: string, openedAt: number, pinned?: boolean): RecentEntry => ({ path: `/d/${name}.docx`, name: `${name}.docx`, app: "belge", openedAt, ...(pinned ? { pinned } : {}) });

describe("son kullanılanlar listesi", () => {
  it("diskten okunan bozuk ve yinelenen kayıtları atar, sabitlenenleri başa alır", () => {
    const raw = [f("a", 3), null, { path: 5 }, { ...f("b", 2), app: "bilinmeyen" }, f("c", 1, true), f("a", 0)];
    expect(normalizeRecent(raw).map((x) => x.name)).toEqual(["c.docx", "a.docx"]);
    expect(normalizeRecent("bozuk")).toEqual([]);
  });

  it("ekleme yeni dosyayı başa alır, sabitlemeyi korur, sınır sabitlenenleri atmaz", () => {
    let list = [f("p1", 1, true), f("p2", 2, true), f("a", 5), f("b", 4)];
    list = addRecentEntry(list, f("c", 10), 2);
    expect(list.map((x) => x.name)).toEqual(["p2.docx", "p1.docx", "c.docx", "a.docx"]);
    list = addRecentEntry(list, f("p1", 20), 2);
    expect(list.map((x) => [x.name, x.pinned === true])).toEqual([
      ["p1.docx", true],
      ["p2.docx", true],
      ["c.docx", false],
      ["a.docx", false],
    ]);
  });

  it("sabitleme sırayı değiştirir, kaldırma listeden çıkarır", () => {
    let list = [f("a", 3), f("b", 2), f("c", 1)];
    list = pinRecentEntry(list, "/d/c.docx", true);
    expect(list.map((x) => x.name)).toEqual(["c.docx", "a.docx", "b.docx"]);
    list = pinRecentEntry(list, "/d/b.docx", true);
    expect(list.map((x) => x.name)).toEqual(["b.docx", "c.docx", "a.docx"]);
    list = pinRecentEntry(list, "/d/b.docx", false);
    expect(list.map((x) => x.name)).toEqual(["c.docx", "a.docx", "b.docx"]);
    expect(list[2]).not.toHaveProperty("pinned");
    list = removeRecentEntry(list, "/d/a.docx");
    expect(list.map((x) => x.name)).toEqual(["c.docx", "b.docx"]);
  });
});
