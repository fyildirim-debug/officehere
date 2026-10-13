// @vitest-environment jsdom
// Biçim tanıma (uzantı + içerik imzası), açma uyarıları ve LibreOffice ile dış doğrulama.
import { describe, expect, it } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { COMPAT_NOTE, OPEN_ACCEPT, OPEN_FILTERS, detectFormat, openDocument, readDocx, saveFilters, writeOdt, writeRtf } from "../index";
import { all, fixtureBuffer, marksOf, richDoc, text } from "./bicim-yardimci";

describe("biçim tanıma", () => {
  it("içerik imzası uzantıdan önce gelir", async () => {
    expect(await detectFormat("yanlis.docx", fixtureBuffer("lo-ornek.odt"))).toBe("odt");
    expect(await detectFormat("yanlis.txt", fixtureBuffer("lo-ornek.doc"))).toBe("doc");
    expect(await detectFormat("yanlis.doc", fixtureBuffer("lo-ornek.rtf"))).toBe("rtf");
    expect(await detectFormat("yanlis.odt", fixtureBuffer("lo-ornek.docx"))).toBe("docx");
    expect(await detectFormat("adsiz", fixtureBuffer("lo-ornek.odt"))).toBe("odt");
    expect(await detectFormat("not.md", new TextEncoder().encode("# Başlık").buffer)).toBe("md");
    expect(await detectFormat("not.markdown", new TextEncoder().encode("x").buffer)).toBe("md");
    expect(await detectFormat("sablon.dotm", fixtureBuffer("lo-ornek.docx"))).toBe("docx");
    expect(await detectFormat("sayfa.doc", new TextEncoder().encode("<!DOCTYPE html><html><body>x</body></html>").buffer)).toBe("html");
  });

  it("açma: yanlış uzantılı dosya doğru okuyucuyla açılır; .doc uyumluluk uyarısı verir", async () => {
    const odt = await openDocument("rapor.docx", fixtureBuffer("lo-ornek.odt"));
    expect(odt.format).toBe("odt");
    expect(text(odt.doc.content)).toContain("Giriş ve Amaç");
    const doc = await openDocument("rapor.doc", fixtureBuffer("lo-ornek.doc"));
    expect(doc.format).toBe("doc");
    expect(doc.warning).toBe(COMPAT_NOTE);
    expect(COMPAT_NOTE).not.toMatch(/Word|Microsoft/);
    const rtf = await openDocument("rapor.rtf", fixtureBuffer("lo-ornek.rtf"));
    expect(rtf.format).toBe("rtf");
    expect(all(rtf.doc.content, "table")).toHaveLength(1);
  });

  it("bozuk dosyalar anlaşılır hatayla reddedilir", async () => {
    await expect(openDocument("bozuk.doc", new TextEncoder().encode("bozuk içerik").buffer)).rejects.toThrow(/\.doc/);
    await expect(openDocument("bozuk.rtf", new TextEncoder().encode("rtf değil").buffer)).rejects.toThrow(/RTF/);
    await expect(openDocument("bozuk.odt", new TextEncoder().encode("odt değil").buffer)).rejects.toThrow(/ODT/);
  });

  it("aç ve kaydet filtreleri", () => {
    expect(OPEN_FILTERS[0].name).toBe("Tüm Belgeler");
    expect(OPEN_FILTERS[0].extensions).toEqual(["docx", "docm", "dotx", "dotm", "doc", "odt", "ott", "rtf", "md", "markdown", "html", "htm", "txt"]);
    expect(OPEN_FILTERS.map((f) => f.name)).toContain("OpenDocument Metin (.odt)");
    expect(OPEN_ACCEPT).toContain(".odt");
    expect(OPEN_ACCEPT).toContain(".doc,");
    expect(saveFilters("odt").map((f) => f.extensions[0])).toEqual(["odt", "docx", "rtf", "md", "html", "txt"]);
    expect(JSON.stringify(OPEN_FILTERS)).not.toMatch(/Word|Microsoft/);
  });
});

/** LibreOffice kuruluysa yazılan dosyaların onun tarafından açılabildiği doğrulanır. */
const hasSoffice = (() => {
  try {
    execFileSync("soffice", ["--version"], { stdio: "ignore", timeout: 60_000 });
    return true;
  } catch {
    return false;
  }
})();

describe.skipIf(!hasSoffice)("LibreOffice ile doğrulama", () => {
  it("yazılan .odt ve .rtf LibreOffice ile .docx'e çevrilir; metin ve biçim korunur", async () => {
    const dir = mkdtempSync(join(tmpdir(), "oh-belge-"));
    try {
      const src = richDoc();
      writeFileSync(join(dir, "ornek-odt.odt"), await writeOdt(src));
      writeFileSync(join(dir, "ornek-rtf.rtf"), writeRtf(src));
      execFileSync("soffice", ["--headless", `-env:UserInstallation=file://${join(dir, "profil")}`, "--convert-to", "docx", "--outdir", dir, join(dir, "ornek-odt.odt"), join(dir, "ornek-rtf.rtf")], { stdio: "ignore", timeout: 150_000 });
      for (const f of ["ornek-odt.docx", "ornek-rtf.docx"]) {
        const d = await readDocx(readFileSync(join(dir, f)));
        const s = text(d.content);
        expect(s).toContain("Rapor Başlığı İĞÜŞÖÇ");
        expect(s).toContain("ığüşöç");
        expect(s).toContain("Birleşik");
        expect(all(d.content, "heading").map((h) => h.attrs?.level)).toEqual([1, 2]);
        expect(marksOf(d.content, "kalın")).toContain("bold");
        expect(marksOf(d.content, "italik")).toContain("italic");
        expect(all(d.content, "table")).toHaveLength(1);
        expect(all(d.content, "image")).toHaveLength(1);
        expect(all(d.content, "note")[0]?.attrs?.text).toContain("Dipnot metni");
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 200_000);
});
