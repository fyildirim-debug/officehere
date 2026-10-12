import { describe, expect, it } from "vitest";
import {
  fileCommands,
  foldText,
  loadRecentCommands,
  pushRecentCommand,
  rankCommands,
  RECENT_COMMANDS_LIMIT,
  recentCommandsKey,
  resolveRecentCommands,
  saveRecentCommands,
  type CommandEntry,
} from "./commandIndex";

const e = (title: string, tabLabel = "Giriş", group = "", desc = ""): CommandEntry => ({ id: `${tabLabel}:${title}`, title, desc, tabId: "t", tabLabel, group, nth: 0 });

describe("foldText: Türkçe büyük/küçük harf", () => {
  it("kalın ↔ KALIN ↔ kalin", () => {
    expect(foldText("KALIN")).toBe(foldText("kalın"));
    expect(foldText("kalın")).toBe("kalin");
    expect(foldText("Kalın")).toBe(foldText("kalin"));
  });

  it("i/İ ve ı/I aynı harf sayılır", () => {
    expect(foldText("İtalik")).toBe("italik");
    expect(foldText("ITALIK")).toBe("italik");
    expect(foldText("IŞIK")).toBe(foldText("ışık"));
    expect(foldText("ışık")).toBe("isik");
    expect(foldText("İ")).toBe(foldText("i"));
    expect(foldText("I")).toBe(foldText("ı"));
  });

  it("şapka ve Türkçe harf işaretleri yok sayılır", () => {
    expect(foldText("Gözden Geçir")).toBe("gozden gecir");
    expect(foldText("ŞEKİL")).toBe("sekil");
    expect(foldText("Kâğıt")).toBe("kagit");
  });
});

describe("rankCommands", () => {
  const list = [
    e("Kalın", "Giriş", "Yazı Tipi"),
    e("Kalın Dış Kenarlıklar", "Giriş", "Yazı Tipi"),
    e("Alt Çizgi", "Giriş", "Yazı Tipi", "Seçili metnin altını çizer."),
    e("İtalik", "Giriş", "Yazı Tipi"),
    e("Slaytı Gizle", "Slayt Gösterisi"),
    e("Gri Tonlama", "Görünüm"),
    e("Bul", "Giriş", "Düzenleme", "Sunudaki metni bulun."),
  ];

  it("büyük harfle ya da şapkasız yazılan sorguyu bulur", () => {
    expect(rankCommands(list, "KALIN")[0].title).toBe("Kalın");
    expect(rankCommands(list, "kalin")[0].title).toBe("Kalın");
    expect(rankCommands(list, "italik")[0].title).toBe("İtalik");
    expect(rankCommands(list, "İTALİK")[0].title).toBe("İtalik");
    expect(rankCommands(list, "slayti gizle")[0].title).toBe("Slaytı Gizle");
  });

  it("tam ad önce, sonra ad başı, sözcük başı, içerik, açıklama ve sekme gelir", () => {
    expect(rankCommands(list, "kalın").map((c) => c.title)).toEqual(["Kalın", "Kalın Dış Kenarlıklar"]);
    expect(rankCommands(list, "çizgi")[0].title).toBe("Alt Çizgi");
    // Yalnız açıklamada geçen sözcük: eşit puanda kısa ad önce.
    expect(rankCommands(list, "metni").map((c) => c.title)).toEqual(["Bul", "Alt Çizgi"]);
    // Ad içinde geçen sözcük, yalnız açıklamada geçenden önce gelir.
    expect(rankCommands([e("Metin Kutusu"), e("Bul", "Giriş", "", "Belgede metin arar")], "metin").map((c) => c.title)).toEqual(["Metin Kutusu", "Bul"]);
    // Yalnız sekme adında geçen sözcük de eşleşir ama en sona kalır.
    expect(rankCommands(list, "görünüm").map((c) => c.title)).toEqual(["Gri Tonlama"]);
  });

  it("her sözcük eşleşmelidir; boş sorgu sonuç vermez", () => {
    expect(rankCommands(list, "gri ton").map((c) => c.title)).toEqual(["Gri Tonlama"]);
    expect(rankCommands(list, "kalın gizle")).toEqual([]);
    expect(rankCommands(list, "yok böyle")).toEqual([]);
    expect(rankCommands(list, "   ")).toEqual([]);
  });

  it("eşit puanda kısa ad önce gelir", () => {
    const r = rankCommands([e("Kenarlıklar ve Gölgelendirme"), e("Kenarlık")], "kenar");
    expect(r.map((c) => c.title)).toEqual(["Kenarlık", "Kenarlıklar ve Gölgelendirme"]);
  });
});

describe("son kullanılan komutlar", () => {
  const mem = () => {
    const m = new Map<string, string>();
    return { getItem: (k: string) => m.get(k) ?? null, setItem: (k: string, v: string) => void m.set(k, v), m };
  };

  it("en yeni başa gelir, tekrar edilmez, en çok 5 kayıt tutulur", () => {
    let list = pushRecentCommand([], e("Kalın"));
    list = pushRecentCommand(list, e("İtalik"));
    list = pushRecentCommand(list, e("Kalın"));
    expect(list.map((r) => r.title)).toEqual(["Kalın", "İtalik"]);
    for (const t of ["A", "B", "C", "D", "E", "F"]) list = pushRecentCommand(list, e(t));
    expect(list).toHaveLength(RECENT_COMMANDS_LIMIT);
    expect(list.map((r) => r.title)).toEqual(["F", "E", "D", "C", "B"]);
  });

  it("uygulama başına ayrı anahtarda saklanır; bozuk kayıt yok sayılır", () => {
    const s = mem();
    saveRecentCommands("belge", pushRecentCommand([], e("Kalın")), s);
    expect(recentCommandsKey("belge")).not.toBe(recentCommandsKey("tablo"));
    expect(loadRecentCommands("belge", s).map((r) => r.title)).toEqual(["Kalın"]);
    expect(loadRecentCommands("tablo", s)).toEqual([]);
    s.m.set(recentCommandsKey("sunum"), "{bozuk");
    expect(loadRecentCommands("sunum", s)).toEqual([]);
    s.m.set(recentCommandsKey("sunum"), JSON.stringify([{ id: 1 }, { id: "x", title: "X", tabLabel: "Giriş" }]));
    expect(loadRecentCommands("sunum", s).map((r) => r.id)).toEqual(["x"]);
    expect(loadRecentCommands("belge", null)).toEqual([]);
  });

  it("kayıtlar güncel dizindeki komutlara eşlenir (kimlik değişse de sekme + ad ile)", () => {
    const kalin = e("Kalın");
    const index = [kalin, e("İtalik"), { ...e("Kaydet", "Dosya"), id: "file:save" }];
    const stored = [
      { id: "file:save", title: "Kaydet", tabLabel: "Dosya" },
      { id: "eski-kimlik", title: "Kalın", tabLabel: "Giriş" },
      { id: "yok", title: "Kaldırılan", tabLabel: "Giriş" },
    ];
    expect(resolveRecentCommands(stored, index).map((c) => c.title)).toEqual(["Kaydet", "Kalın"]);
  });
});

describe("fileCommands", () => {
  it("Dosya eylemlerini aranabilir komutlar olarak verir ve çalıştırır", () => {
    const calls: string[] = [];
    const h = {
      newDoc: () => calls.push("new"),
      open: () => calls.push("open"),
      save: () => calls.push("save"),
      saveAs: () => calls.push("saveAs"),
      exportPdf: () => calls.push("pdf"),
      print: () => calls.push("print"),
      backstage: (p: string) => calls.push(`backstage:${p}`),
    };
    const list = fileCommands(h, "Belge");
    const titles = list.map((c) => c.title);
    for (const t of ["Yeni Belge", "Aç", "Kaydet", "Farklı Kaydet", "PDF olarak dışarı aktar", "Yazdır", "Seçenekler"]) expect(titles).toContain(t);
    expect(rankCommands(list, "farklı")[0].title).toBe("Farklı Kaydet");
    expect(rankCommands(list, "pdf")[0].title).toBe("PDF olarak dışarı aktar");
    expect(rankCommands(list, "dışa aktar").map((c) => c.title).slice(0, 2)).toEqual(["Dışarı Aktar", "PDF olarak dışarı aktar"]);
    expect(rankCommands(list, "SEÇENEKLER")[0].title).toBe("Seçenekler");
    rankCommands(list, "farklı kaydet")[0].run!();
    rankCommands(list, "seçenekler")[0].run!();
    expect(calls).toEqual(["saveAs", "backstage:options"]);
    expect(list.every((c) => c.id.startsWith("file:") && c.tabLabel === "Dosya")).toBe(true);
  });
});
