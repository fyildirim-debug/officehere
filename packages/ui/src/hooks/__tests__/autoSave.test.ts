import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  AUTO_SAVE_KEY,
  backupPathFor,
  canAutoSavePath,
  createIdleScheduler,
  ensureBackup,
  readAutoSavePref,
  resetBackupsForTest,
  runAutoSave,
  type AutoSaveIo,
  type AutoSaveTurn,
} from "../autoSaveCore";
import { resolveSaveState, saveStatus, trackSave } from "../saveStatus";

/** Bellekte dosya sistemi: okuma/yazma çağrıları kaydedilir. */
function memoryIo(files: Record<string, string>, opts: { failWrite?: (p: string) => boolean } = {}) {
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  const reads: string[] = [];
  const writes: string[] = [];
  const io: AutoSaveIo = {
    async read(p) {
      reads.push(p);
      return p in files ? (enc.encode(files[p]).buffer as ArrayBuffer) : null;
    },
    async write(p, data) {
      writes.push(p);
      if (opts.failWrite?.(p)) return false;
      files[p] = dec.decode(data);
      return true;
    },
  };
  return { io, files, reads, writes };
}

function reporter() {
  const events: string[] = [];
  return { events, report: { saving: () => events.push("saving"), saved: () => events.push("saved"), failed: () => events.push("failed") } };
}

function turn(over: Partial<AutoSaveTurn> = {}): AutoSaveTurn & { cleaned: number } {
  const t = {
    cleaned: 0,
    path: "/belgeler/rapor.docx",
    formats: ["docx"],
    enabled: true,
    isDirty: () => true,
    prepare: async () => ({ data: new TextEncoder().encode("YENİ"), isCurrent: () => true }),
    markClean: () => {
      t.cleaned++;
    },
    ...over,
  };
  return t;
}

beforeEach(() => resetBackupsForTest());

describe("seçenek", () => {
  it("varsayılan açık; yalnız açıkça kapatılınca kapalı", () => {
    const store = new Map<string, string>();
    const storage = { getItem: (k: string) => store.get(k) ?? null, setItem: (k: string, v: string) => void store.set(k, v) };
    expect(readAutoSavePref(storage)).toBe(true);
    store.set(AUTO_SAVE_KEY, "0");
    expect(readAutoSavePref(storage)).toBe(false);
    store.set(AUTO_SAVE_KEY, "1");
    expect(readAutoSavePref(storage)).toBe(true);
    expect(readAutoSavePref({ getItem: () => { throw new Error("kapalı"); }, setItem: () => {} })).toBe(true);
  });
});

describe("hangi dosya otomatik kaydedilir", () => {
  it("yalnız kendi biçimlerimiz (docx/xlsx/pptx), büyük-küçük harf fark etmez", () => {
    for (const p of ["C:\\Belgeler\\Rapor.DOCX", "/ev/tablo.xlsx", "/ev/sunu.pptx"]) expect(canAutoSavePath(p)).toBe(true);
  });

  it("biçim kaybettiren ya da makro/şablon dosyaları ve yolu olmayan belgeler kaydedilmez", () => {
    for (const p of ["/ev/veri.csv", "/ev/sayfa.html", "/ev/not.txt", "/ev/makro.xlsm", "/ev/makro.docm", "/ev/sablon.dotx", "/ev/eski.odt", "/ev/uzantisiz"]) {
      expect(canAutoSavePath(p)).toBe(false);
    }
    expect(canAutoSavePath(undefined)).toBe(false);
    expect(canAutoSavePath("")).toBe(false);
  });

  it("uygulama kendi biçimiyle sınırlar", () => {
    expect(canAutoSavePath("/ev/tablo.xlsx", ["docx"])).toBe(false);
    expect(canAutoSavePath("/ev/rapor.docx", ["docx"])).toBe(true);
  });

  it("yedek dosyası özgün adın sonuna .bak alır", () => {
    expect(backupPathFor("/ev/rapor.docx")).toBe("/ev/rapor.docx.bak");
  });
});

describe("oturumluk yedek", () => {
  it("özgün dosyayı bir kez .bak olarak kopyalar", async () => {
    const m = memoryIo({ "/ev/a.docx": "ÖZGÜN" });
    expect(await ensureBackup("/ev/a.docx", m.io)).toBe(true);
    expect(await ensureBackup("/ev/a.docx", m.io)).toBe(true);
    expect(m.files["/ev/a.docx.bak"]).toBe("ÖZGÜN");
    expect(m.reads).toEqual(["/ev/a.docx"]);
    expect(m.writes).toEqual(["/ev/a.docx.bak"]);
  });

  it("okunamayan ya da yedeği yazılamayan dosya için false döner ve yeniden denemez", async () => {
    const missing = memoryIo({});
    expect(await ensureBackup("/ev/yok.docx", missing.io)).toBe(false);
    expect(await ensureBackup("/ev/yok.docx", missing.io)).toBe(false);
    expect(missing.reads).toHaveLength(1);
    const ro = memoryIo({ "/ev/b.docx": "x" }, { failWrite: (p) => p.endsWith(".bak") });
    expect(await ensureBackup("/ev/b.docx", ro.io)).toBe(false);
  });
});

describe("boşta kalma zamanlayıcısı", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("yazma bittikten 2 sn sonra bir kez çalışır", () => {
    const run = vi.fn();
    const s = createIdleScheduler(run, 2000, 30_000);
    s.touch();
    vi.advanceTimersByTime(1500);
    s.touch();
    vi.advanceTimersByTime(1999);
    expect(run).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(run).toHaveBeenCalledTimes(1);
    expect(s.pending).toBe(false);
  });

  it("sürekli yazmada en geç 30 sn içinde çalışır", () => {
    const run = vi.fn();
    const s = createIdleScheduler(run, 2000, 30_000);
    for (let i = 0; i < 29; i++) {
      s.touch();
      vi.advanceTimersByTime(1000);
    }
    expect(run).not.toHaveBeenCalled();
    s.touch();
    vi.advanceTimersByTime(1000);
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("iptal edilince çalışmaz", () => {
    const run = vi.fn();
    const s = createIdleScheduler(run, 2000, 30_000);
    s.touch();
    s.cancel();
    vi.advanceTimersByTime(60_000);
    expect(run).not.toHaveBeenCalled();
  });
});

describe("otomatik kaydetme turu", () => {
  it("önce yedek alır, sonra kendi biçiminde yazar ve kirli durumu temizler", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("saved");
    expect(m.writes).toEqual(["/belgeler/rapor.docx.bak", "/belgeler/rapor.docx"]);
    expect(m.files["/belgeler/rapor.docx.bak"]).toBe("ÖZGÜN");
    expect(m.files["/belgeler/rapor.docx"]).toBe("YENİ");
    expect(r.events).toEqual(["saving", "saved"]);
    expect(t.cleaned).toBe(1);
    // Aynı oturumda ikinci kayıtta yedek yeniden alınmaz.
    await runAutoSave(turn(), m.io, r.report);
    expect(m.writes.filter((p) => p.endsWith(".bak"))).toHaveLength(1);
  });

  it("yedek alınamazsa belge yazılmaz ve rozet değişmez", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" }, { failWrite: (p) => p.endsWith(".bak") });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("no-backup");
    expect(m.files["/belgeler/rapor.docx"]).toBe("ÖZGÜN");
    expect(m.writes).toEqual(["/belgeler/rapor.docx.bak"]);
    expect(r.events).toEqual([]);
    expect(t.cleaned).toBe(0);
  });

  it("seçenek kapalıysa, belge temizse ya da biçim uygun değilse hiçbir şey yapmaz", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN", "/belgeler/veri.csv": "a;b" });
    const r = reporter();
    expect(await runAutoSave(turn({ enabled: false }), m.io, r.report)).toBe("skipped");
    expect(await runAutoSave(turn({ isDirty: () => false }), m.io, r.report)).toBe("skipped");
    expect(await runAutoSave(turn({ path: "/belgeler/veri.csv", formats: ["xlsx"] }), m.io, r.report)).toBe("skipped");
    expect(await runAutoSave(turn({ path: undefined }), m.io, r.report)).toBe("skipped");
    expect(m.reads).toEqual([]);
    expect(m.writes).toEqual([]);
    expect(r.events).toEqual([]);
  });

  it("yazarken belge değiştiyse yazılır ama kirli kalır", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const r = reporter();
    const t = turn({ prepare: async () => ({ data: new TextEncoder().encode("ARA"), isCurrent: () => false }) });
    expect(await runAutoSave(t, m.io, r.report)).toBe("changed");
    expect(m.files["/belgeler/rapor.docx"]).toBe("ARA");
    expect(t.cleaned).toBe(0);
  });

  it("yazma ya da üretme başarısızsa 'Kaydedilemedi' bildirir", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" }, { failWrite: (p) => p.endsWith(".docx") });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("failed");
    expect(r.events).toEqual(["saving", "failed"]);
    expect(t.cleaned).toBe(0);
    const r2 = reporter();
    const broken = turn({ prepare: async () => Promise.reject(new Error("bozuk")) });
    expect(await runAutoSave(broken, memoryIo({ "/belgeler/rapor.docx": "x" }).io, r2.report)).toBe("failed");
    expect(r2.events).toEqual(["saving", "failed"]);
  });
});

describe("kayıt durumu rozeti", () => {
  afterEach(() => saveStatus.reset());

  it("kaydetme sürerken ve hatada kirli bayraktan önce gelir", () => {
    expect(resolveSaveState(false, { phase: "idle" })).toBeNull();
    expect(resolveSaveState(true, { phase: "idle" })).toBe("dirty");
    expect(resolveSaveState(true, { phase: "saved", savedAt: 1 })).toBe("dirty");
    expect(resolveSaveState(false, { phase: "saved", savedAt: 1 })).toBe("saved");
    expect(resolveSaveState(true, { phase: "saving" })).toBe("saving");
    expect(resolveSaveState(true, { phase: "error" })).toBe("error");
  });

  it("elle Kaydet: doğrudan yazmada başarı/başarısızlık, konum sorulup iptal edilince değişmez", async () => {
    await trackSave(true, async () => ({ ok: true }));
    expect(saveStatus.get().phase).toBe("saved");
    expect(saveStatus.get().savedAt).toBeTypeOf("number");
    await trackSave(true, async () => null);
    expect(saveStatus.get().phase).toBe("error");
    saveStatus.reset();
    await trackSave(false, async () => null);
    expect(saveStatus.get().phase).toBe("idle");
    await expect(trackSave(false, async () => Promise.reject(new Error("x")))).rejects.toThrow("x");
    expect(saveStatus.get().phase).toBe("error");
  });
});
