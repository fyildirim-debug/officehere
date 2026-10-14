// Yerel günlük: boyut sınırlı döner dosyalar, kişisel içerik (belge yolu, ev dizini) yazılmaz.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { describePath, LocalLog, sanitize, sanitizeStack } from "../src/main/log";
import { saveErrorMessage, saveErrorReason } from "../src/main/errors";

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), "oh-gunluk-"));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe("günlük", () => {
  it("boyut sınırını aşınca döner; en çok 3 dosya", () => {
    const log = new LocalLog({ dir, maxBytes: 2000, files: 3 });
    for (let i = 0; i < 200; i++) log.write("hata", "ana", `ileti ${i} ${"x".repeat(40)}`);
    const files = readdirSync(dir).sort();
    expect(files).toEqual(["officehere.1.log", "officehere.2.log", "officehere.log"]);
    for (const f of files) expect(statSync(path.join(dir, f)).size).toBeLessThanOrEqual(2000);
    const tail = log.tail(5);
    expect(tail).toHaveLength(5);
    expect(tail[4]).toContain("ileti 199");
    expect(tail[0]).toContain("ileti 195");
  });

  it("son satırlar gerekirse önceki dosyadan tamamlanır", () => {
    const log = new LocalLog({ dir, maxBytes: 400, files: 3 });
    for (let i = 0; i < 6; i++) log.write("bilgi", "test", `satır ${i} ${"y".repeat(60)}`);
    const tail = log.tail(4);
    expect(tail.map((l) => /satır (\d)/.exec(l)?.[1])).toEqual(["2", "3", "4", "5"]);
  });

  it("belge yolları özetlenir, ev dizini gizlenir, kod konumları dosya adına iner", () => {
    const home = "/home/ahmet";
    const s = sanitize("“/home/ahmet/Belgeler/Gizli Proje/maaşlar.xlsx” açılamadı; C:\\Users\\ahmet\\Desktop\\dilekce.docx da", { home });
    expect(s).not.toContain("maaşlar");
    expect(s).not.toContain("dilekce");
    expect(s).not.toContain("ahmet");
    expect(s).toContain(".xlsx");
    expect(s).toMatch(/#[0-9a-f]{8}\.docx/);
    expect(describePath("/a/b/c.DOCX")).toMatch(/^#[0-9a-f]{8}\.docx$/);
    const stack = sanitizeStack("TypeError: x\n    at f (file:///home/ahmet/app/resources/apps/belge/assets/index-abc.js:12:3)\n    at g (/home/ahmet/kod/main.js:1:1)\nnot a frame", { home });
    expect(stack).toEqual(["at f (index-abc.js:12:3)", "at g (main.js:1:1)"]);
  });

  it("satırlar tek satıra indirgenir, uzunluk sınırlanır; yığın izi girintili yazılır", () => {
    const log = new LocalLog({ dir, home: "/home/x" });
    log.write("hata", "belge", `çok\nsatırlı ${"z".repeat(2000)}`, { stack: "Error: a\n    at q (/home/x/a.js:1:2)", code: "ENOSPC" });
    const text = readFileSync(path.join(dir, "officehere.log"), "utf8");
    const [first, second] = text.split("\n");
    expect(first).toMatch(/^\d{4}-\d\d-\d\dT.* HATA \[belge\] çok satırlı z+… \(ENOSPC\)$/);
    expect(first.length).toBeLessThan(700);
    expect(second).toBe("    at q (a.js:1:2)");
  });

  it("günlük yazılamazsa uygulama etkilenmez", () => {
    // Dizin yerine düz dosya: altına günlük yazılamaz.
    const blocker = path.join(dir, "dosya");
    writeFileSync(blocker, "x");
    const log = new LocalLog({ dir: path.join(blocker, "gunluk") });
    expect(() => log.write("hata", "x", "y")).not.toThrow();
    expect(log.tail()).toEqual([]);
  });
});

describe("hata nedenleri", () => {
  it("Türkçe kısa neden ve öneri", () => {
    expect(saveErrorReason("ENOSPC")).toBe("Disk dolu");
    expect(saveErrorReason("EACCES")).toBe("Yazma izni yok");
    expect(saveErrorReason("EBUSY")).toBe("Dosya başka bir programda açık");
    expect(saveErrorReason("EROFS")).toBe("Disk salt okunur");
    expect(saveErrorReason("BILINMEYEN")).toBe("Dosya yazılamadı");
    expect(saveErrorMessage("ENOSPC")).toMatch(/^Disk dolu\. /);
  });
});
