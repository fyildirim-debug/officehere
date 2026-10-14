// Masaüstünde yazma hatası kodu (ENOSPC, EACCES…) sayfaya Türkçe nedeniyle ulaşır.
import { afterEach, describe, expect, it } from "vitest";
import { SaveError, describeSaveError, saveErrorReason, saveFile, saveFileAs, type OfficehereNative } from "../src";

function withNative(n: Partial<OfficehereNative>) {
  (globalThis as unknown as { window: { officehereNative?: unknown } }).window = { officehereNative: n };
}

afterEach(() => {
  delete (globalThis as unknown as { window?: unknown }).window;
});

describe("kaydetme hataları", () => {
  it("kodlar Türkçe kısa nedene çevrilir", () => {
    expect(saveErrorReason("ENOSPC")).toBe("Disk dolu");
    expect(saveErrorReason("EACCES")).toBe("Yazma izni yok");
    expect(saveErrorReason("EBUSY")).toBe("Dosya başka bir programda açık");
    expect(saveErrorReason("QuotaExceededError")).toBe("Disk dolu");
    expect(saveErrorReason(undefined)).toBe("Dosya yazılamadı");
    expect(describeSaveError(new SaveError("EROFS"))).toBe("Disk salt okunur");
  });

  it("Kaydet: ana süreç yazamazsa SaveError (kod + neden + ileti) fırlatılır", async () => {
    withNative({ writeFileSafe: async () => ({ ok: false, code: "ENOSPC", reason: "Disk dolu", message: "Disk dolu. Diskte yer açın." }), writeFile: async () => false });
    const err = (await saveFile({ name: "a.docx", path: "/ev/a.docx" }, "a.docx", [], new Uint8Array([1])).catch((e: unknown) => e)) as SaveError;
    expect(err).toBeInstanceOf(SaveError);
    expect(err.code).toBe("ENOSPC");
    expect(err.reason).toBe("Disk dolu");
    expect(err.message).toBe("Disk dolu. Diskte yer açın.");
  });

  it("Kaydet: başarıda dosya tanıtıcısı döner; eski kabukta (yalnız writeFile) da çalışır", async () => {
    withNative({ writeFileSafe: async () => ({ ok: true }) });
    const ref = { name: "a.docx", path: "/ev/a.docx" };
    expect(await saveFile(ref, "a.docx", [], "x")).toBe(ref);
    withNative({ writeFile: async () => false });
    await expect(saveFile(ref, "a.docx", [], "x")).rejects.toBeInstanceOf(SaveError);
  });

  it("Farklı Kaydet: { error } sonucu SaveError olur, iptal null kalır", async () => {
    withNative({ saveFileAs: async () => ({ error: { code: "EACCES", reason: "Yazma izni yok", message: "Yazma izni yok." } }) });
    await expect(saveFileAs("a.docx", [], "x")).rejects.toMatchObject({ code: "EACCES", reason: "Yazma izni yok" });
    withNative({ saveFileAs: async () => null });
    expect(await saveFileAs("a.docx", [], "x")).toBeNull();
    withNative({ saveFileAs: async () => ({ path: "/ev/b.docx", name: "b.docx" }) });
    expect(await saveFileAs("b.docx", [], "x")).toEqual({ path: "/ev/b.docx", name: "b.docx" });
  });
});
