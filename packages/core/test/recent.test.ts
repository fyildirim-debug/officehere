import { describe, expect, it } from "vitest";
import {
  filterRecent,
  folderOf,
  formatRelativeTime,
  formatShortDate,
  matchesRecentQuery,
  orderRecent,
  shortenPath,
  truncateMiddle,
  withRecentAdded,
  withRecentPinned,
  withRecentRemoved,
} from "../src/recent";
import type { RecentFile } from "../src/bridge";

const at = (y: number, m: number, d: number, h = 12, min = 0) => new Date(y, m - 1, d, h, min).getTime();

describe("formatRelativeTime", () => {
  const now = at(2026, 10, 12, 15, 30);

  it("bir dakikadan yeni ve gelecekteki zamanlar 'az önce'dir", () => {
    expect(formatRelativeTime(now, now)).toBe("az önce");
    expect(formatRelativeTime(now - 59_000, now)).toBe("az önce");
    expect(formatRelativeTime(now + 5_000, now)).toBe("az önce");
  });

  it("dakika ve saat", () => {
    expect(formatRelativeTime(now - 60_000, now)).toBe("1 dakika önce");
    expect(formatRelativeTime(now - 5 * 60_000, now)).toBe("5 dakika önce");
    expect(formatRelativeTime(now - 59 * 60_000, now)).toBe("59 dakika önce");
    expect(formatRelativeTime(now - 2 * 3_600_000, now)).toBe("2 saat önce");
    expect(formatRelativeTime(at(2026, 10, 12, 0, 5), now)).toBe("15 saat önce");
  });

  it("takvim gününe göre dün ve gün sayısı", () => {
    expect(formatRelativeTime(at(2026, 10, 11, 23, 50), now)).toBe("dün");
    expect(formatRelativeTime(at(2026, 10, 11, 8, 0), now)).toBe("dün");
    expect(formatRelativeTime(at(2026, 10, 9, 18, 0), now)).toBe("3 gün önce");
    expect(formatRelativeTime(at(2026, 10, 6, 9, 0), now)).toBe("6 gün önce");
  });

  it("bir haftadan eski tarih gg.aa.yyyy yazılır", () => {
    expect(formatRelativeTime(at(2026, 10, 5, 9, 0), now)).toBe("05.10.2026");
    expect(formatRelativeTime(at(2025, 1, 3), now)).toBe("03.01.2025");
    expect(formatShortDate(at(2026, 12, 31))).toBe("31.12.2026");
  });
});

describe("folderOf", () => {
  it("Windows ve Unix yollarının klasörü", () => {
    expect(folderOf("C:\\Users\\furkan\\Belgeler\\rapor.docx")).toBe("C:\\Users\\furkan\\Belgeler");
    expect(folderOf("C:\\rapor.docx")).toBe("C:\\");
    expect(folderOf("/home/furkan/rapor.docx")).toBe("/home/furkan");
    expect(folderOf("/rapor.docx")).toBe("/");
    expect(folderOf("rapor.docx")).toBe("");
  });
});

describe("shortenPath", () => {
  it("kısa yol olduğu gibi kalır", () => {
    expect(shortenPath("C:\\Belgeler", 48)).toBe("C:\\Belgeler");
  });

  it("kökü ve son klasörleri koruyarak ortadan kısaltır", () => {
    const p = "C:\\Users\\furkan\\OneDrive\\Belgeler\\Projeler\\2026\\Ekim";
    const s = shortenPath(p, 30);
    expect(s).toBe("C:\\…\\Projeler\\2026\\Ekim");
    expect(s.length).toBeLessThanOrEqual(30);
    expect(shortenPath("/home/furkan/Belgeler/Projeler/OfficeHere/Raporlar", 32)).toBe("/home/…/OfficeHere/Raporlar");
  });

  it("son klasör bile sığmazsa karakter düzeyinde kısaltır", () => {
    const p = "/home/" + "a".repeat(80);
    const s = shortenPath(p, 20);
    expect(s).toHaveLength(20);
    expect(s).toContain("…");
    expect(s.startsWith("/home/")).toBe(true);
  });

  it("truncateMiddle baştan ve sondan eşit tutar", () => {
    expect(truncateMiddle("abcdefghij", 7)).toBe("abc…hij");
    expect(truncateMiddle("abc", 7)).toBe("abc");
  });
});

describe("arama", () => {
  const files = [
    { name: "İSTANBUL Raporu.docx", path: "/belgeler/İSTANBUL Raporu.docx" },
    { name: "IŞIK ölçümleri.xlsx", path: "/veri/IŞIK ölçümleri.xlsx" },
    { name: "Invoice.pptx", path: "/sunular/Invoice.pptx" },
    { name: "bütçe.xlsx", path: "/Muhasebe/2026/bütçe.xlsx" },
  ];

  it("Türkçe büyük/küçük harf duyarsız", () => {
    expect(filterRecent(files, "istanbul").map((f) => f.name)).toEqual(["İSTANBUL Raporu.docx"]);
    expect(filterRecent(files, "ışık").map((f) => f.name)).toEqual(["IŞIK ölçümleri.xlsx"]);
    expect(filterRecent(files, "invoice").map((f) => f.name)).toEqual(["Invoice.pptx"]);
    expect(filterRecent(files, "BÜTÇE").map((f) => f.name)).toEqual(["bütçe.xlsx"]);
  });

  it("yolda da arar, sözcüklerin hepsi eşleşmeli", () => {
    expect(filterRecent(files, "muhasebe").map((f) => f.name)).toEqual(["bütçe.xlsx"]);
    expect(matchesRecentQuery(files[3], "2026 bütçe")).toBe(true);
    expect(matchesRecentQuery(files[3], "2026 rapor")).toBe(false);
    expect(filterRecent(files, "  ")).toHaveLength(4);
  });
});

describe("liste işlemleri", () => {
  const f = (name: string, extra: Partial<RecentFile> = {}): RecentFile => ({ path: `/d/${name}`, name, app: "belge", openedAt: 0, ...extra });

  it("sabitlenenler başa alınır; iki grup da en son açılan önce", () => {
    const list = [f("a"), f("b", { pinned: true }), f("c"), f("d", { pinned: true })];
    expect(orderRecent(list).map((x) => x.name)).toEqual(["b", "d", "a", "c"]);
    const timed = [f("a", { openedAt: 1 }), f("b", { openedAt: 2, pinned: true }), f("c", { openedAt: 3 }), f("d", { openedAt: 4, pinned: true })];
    expect(orderRecent(timed).map((x) => x.name)).toEqual(["d", "b", "c", "a"]);
  });

  it("ekleme başa alır, sabitlemeyi korur ve sınırı yalnız sabitlenmemişlere uygular", () => {
    let list: RecentFile[] = [f("a", { pinned: true }), f("b", { openedAt: 2 }), f("c", { openedAt: 1 })];
    list = withRecentAdded(list, f("d", { openedAt: 3 }), 2);
    expect(list.map((x) => x.name)).toEqual(["a", "d", "b"]);
    list = withRecentAdded(list, f("a", { openedAt: 5 }), 2);
    expect(list[0]).toMatchObject({ name: "a", pinned: true, openedAt: 5 });
  });

  it("sabitleme, sabitlemeyi kaldırma ve listeden çıkarma", () => {
    let list: RecentFile[] = [f("a", { openedAt: 3 }), f("b", { openedAt: 2 }), f("c", { openedAt: 1 })];
    list = withRecentPinned(list, f("c"), true);
    expect(list.map((x) => x.name)).toEqual(["c", "a", "b"]);
    list = withRecentPinned(list, f("c"), false);
    expect(list.map((x) => x.name)).toEqual(["a", "b", "c"]);
    expect(list[2].pinned).toBeUndefined();
    list = withRecentRemoved(list, f("b"));
    expect(list.map((x) => x.name)).toEqual(["a", "c"]);
  });
});
