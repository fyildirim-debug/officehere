import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SaveError, type NativeWriteOptions, type WriteOutcome } from "@officehere/core";
import {
  AUTO_SAVE_KEY,
  AUTO_SAVE_IDLE_MS,
  AUTO_SAVE_MAX_WAIT_MS,
  autoSaveDelays,
  autoSaveDisabledReason,
  backupPathFor,
  canAutoSavePath,
  createIdleScheduler,
  disableAutoSave,
  readAutoSavePref,
  resetAutoSaveForTest,
  runAutoSave,
  runAutoSaveTurn,
  type AutoSaveIo,
  type AutoSaveTurn,
} from "../autoSaveCore";
import { resolveSaveState, saveStatus, trackSave } from "../saveStatus";

/**
 * Bellekte dosya sistemi; ana sürecin yazma kurallarını taklit eder: "sidecar" ile
 * ".bak" yalnız yoksa oluşturulur, "keepPrevious" ile üzerine yazılan içerik saklanır.
 */
function memoryIo(files: Record<string, string>, opts: { failWrite?: (p: string) => string | undefined } = {}) {
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  const reads: string[] = [];
  const writes: { path: string; options: NativeWriteOptions }[] = [];
  const previous: Record<string, string> = {};
  const io: AutoSaveIo = {
    async read(p) {
      reads.push(p);
      return p in files ? (enc.encode(files[p]).buffer as ArrayBuffer) : null;
    },
    async write(p, data, options): Promise<WriteOutcome> {
      writes.push({ path: p, options });
      const code = opts.failWrite?.(p);
      if (code) return { ok: false, code, reason: code === "ENOSPC" ? "Disk dolu" : undefined, message: code === "ENOSPC" ? "Disk dolu. Diskte yer açın." : undefined };
      if (p in files) {
        if (options.sidecar && !(backupPathFor(p) in files)) files[backupPathFor(p)] = files[p];
        if (options.keepPrevious) previous[p] = files[p];
      }
      files[p] = dec.decode(data);
      return { ok: true };
    },
    async revert(p) {
      if (!(p in previous)) return { ok: false, code: "ENOENT" };
      files[p] = previous[p];
      return { ok: true };
    },
  };
  return { io, files, reads, writes };
}

function reporter() {
  const events: string[] = [];
  const report = {
    saving: () => void events.push("saving"),
    saved: () => void events.push("saved"),
    failed: (reason?: string) => void events.push(reason ? `failed:${reason}` : "failed"),
    paused: (reason: string | null) => void events.push(reason ? `paused:${reason}` : "resumed"),
    alert: (title: string) => void events.push(`alert:${title}`),
  };
  return { events, report };
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

beforeEach(() => resetAutoSaveForTest());

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

  it("büyük dosyada aralık seyrelir: son kayıt süresinin 10 katı", () => {
    expect(autoSaveDelays(50)).toEqual({ idleMs: AUTO_SAVE_IDLE_MS, maxWaitMs: AUTO_SAVE_MAX_WAIT_MS });
    expect(autoSaveDelays(1500)).toEqual({ idleMs: 15_000, maxWaitMs: AUTO_SAVE_MAX_WAIT_MS });
    expect(autoSaveDelays(6000)).toEqual({ idleMs: 60_000, maxWaitMs: 60_000 });
    // Süreler işlevse her dokunuşta yeniden okunur.
    let cost = 0;
    const run = vi.fn();
    const s = createIdleScheduler(run, () => autoSaveDelays(cost).idleMs, () => autoSaveDelays(cost).maxWaitMs);
    cost = 3000;
    s.touch();
    vi.advanceTimersByTime(29_999);
    expect(run).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(run).toHaveBeenCalledTimes(1);
  });
});

describe("otomatik kaydetme turu", () => {
  it("yan yedek isteğiyle kendi biçiminde yazar ve kirli durumu temizler", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("saved");
    expect(m.writes).toEqual([{ path: "/belgeler/rapor.docx", options: { sidecar: true, keepPrevious: false } }]);
    expect(m.files["/belgeler/rapor.docx.bak"]).toBe("ÖZGÜN");
    expect(m.files["/belgeler/rapor.docx"]).toBe("YENİ");
    expect(r.events).toEqual(["resumed", "saving", "saved"]);
    expect(t.cleaned).toBe(1);
  });

  it("iki oturum: .bak özgün içeriği korur (ikinci oturum yedeği ezmez)", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    await runAutoSave(turn({ prepare: async () => ({ data: new TextEncoder().encode("OTURUM-1"), isCurrent: () => true }) }), m.io, reporter().report);
    // Uygulama yeniden başlatıldı: bellekteki durum yok.
    resetAutoSaveForTest();
    await runAutoSave(turn({ prepare: async () => ({ data: new TextEncoder().encode("OTURUM-2"), isCurrent: () => true }) }), m.io, reporter().report);
    expect(m.files["/belgeler/rapor.docx"]).toBe("OTURUM-2");
    expect(m.files["/belgeler/rapor.docx.bak"]).toBe("ÖZGÜN");
  });

  it("yedek alınamazsa belge yazılmaz; otomatik kaydetme o dosyada duraklar", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" }, { failWrite: () => "EBACKUP" });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("no-backup");
    expect(m.files["/belgeler/rapor.docx"]).toBe("ÖZGÜN");
    expect(r.events).toEqual(["resumed", "saving", "failed:Özgün dosyanın yedeği alınamadı", "paused:Özgün dosyanın yedeği alınamadı"]);
    expect(t.cleaned).toBe(0);
    expect(autoSaveDisabledReason("/belgeler/rapor.docx")).toBe("Özgün dosyanın yedeği alınamadı");
    expect(await runAutoSave(turn(), m.io, reporter().report)).toBe("disabled");
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

  it("dosya uyarıyla açıldıysa (bekletme) yazılmaz ve rozet nedeni söyler", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const r = reporter();
    expect(await runAutoSave(turn({ hold: "Dosya açılırken bazı öğeler okunamadı" }), m.io, r.report)).toBe("held");
    expect(m.writes).toEqual([]);
    expect(m.files["/belgeler/rapor.docx"]).toBe("ÖZGÜN");
    expect(r.events).toEqual(["paused:Dosya açılırken bazı öğeler okunamadı"]);
  });

  it("yazarken belge değiştiyse yazılır ama kirli kalır", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const r = reporter();
    const t = turn({ prepare: async () => ({ data: new TextEncoder().encode("ARA"), isCurrent: () => false }) });
    expect(await runAutoSave(t, m.io, r.report)).toBe("changed");
    expect(m.files["/belgeler/rapor.docx"]).toBe("ARA");
    expect(t.cleaned).toBe(0);
  });

  it("yazma başarısızsa nedeniyle 'Kaydedilemedi' bildirir (Disk dolu)", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" }, { failWrite: () => "ENOSPC" });
    const r = reporter();
    const t = turn();
    expect(await runAutoSave(t, m.io, r.report)).toBe("failed");
    expect(r.events).toEqual(["resumed", "saving", "failed:Disk dolu"]);
    expect(t.cleaned).toBe(0);
    const r2 = reporter();
    const broken = turn({ prepare: async () => Promise.reject(new Error("bozuk")) });
    expect(await runAutoSave(broken, memoryIo({ "/belgeler/rapor.docx": "x" }).io, r2.report)).toBe("failed");
    expect(r2.events).toEqual(["resumed", "saving", "failed"]);
  });

  it("yazılan dosya doğrulanır: başarılıysa temizlenir", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const verified: string[] = [];
    const t = turn({ verify: async (b) => (verified.push(new TextDecoder().decode(b)), true), idle: async () => {} });
    expect(await runAutoSave(t, m.io, reporter().report)).toBe("saved");
    expect(verified).toEqual(["YENİ"]);
    expect(m.writes[0].options).toEqual({ sidecar: true, keepPrevious: true });
    expect(t.cleaned).toBe(1);
  });

  it("doğrulama başarısız: önceki içerik geri yüklenir, otomatik kaydetme kapanır, kullanıcı uyarılır", async () => {
    const m = memoryIo({ "/belgeler/rapor.docx": "İYİ" });
    const r = reporter();
    const t = turn({ verify: async () => false, idle: async () => {} });
    expect(await runAutoSave(t, m.io, r.report)).toBe("verify-failed");
    expect(m.files["/belgeler/rapor.docx"]).toBe("İYİ");
    expect(t.cleaned).toBe(0);
    expect(r.events).toContain("alert:Otomatik kaydetme kapatıldı");
    expect(r.events).toContain("paused:Kaydedilen dosya doğrulanamadı");
    // Aynı dosyada artık yazılmaz.
    const before = m.writes.length;
    expect(await runAutoSave(turn({ verify: async () => true }), m.io, reporter().report)).toBe("disabled");
    expect(m.writes.length).toBe(before);
    // Okuyucunun hata fırlatması da başarısızlık sayılır.
    resetAutoSaveForTest();
    const m2 = memoryIo({ "/belgeler/rapor.docx": "İYİ" });
    expect(await runAutoSave(turn({ verify: async () => Promise.reject(new Error("açılamadı")) }), m2.io, reporter().report)).toBe("verify-failed");
    expect(m2.files["/belgeler/rapor.docx"]).toBe("İYİ");
  });

  it("başka pencerede açık dosya: otomatik kaydetme kapatılır", async () => {
    disableAutoSave("/belgeler/rapor.docx", "Bu dosya başka bir pencerede de açık");
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    expect(await runAutoSave(turn(), m.io, reporter().report)).toBe("disabled");
    expect(m.writes).toEqual([]);
  });

  it("tur maliyeti ölçülür (aralığı seyreltmek için)", async () => {
    let clock = 1000;
    const m = memoryIo({ "/belgeler/rapor.docx": "ÖZGÜN" });
    const t = turn({
      now: () => clock,
      prepare: async () => {
        clock += 700;
        return { data: new Uint8Array([1]), isCurrent: () => true };
      },
    });
    const r = await runAutoSaveTurn(t, m.io, reporter().report);
    expect(r).toEqual({ result: "saved", costMs: 700 });
  });
});

describe("kayıt durumu rozeti", () => {
  afterEach(() => saveStatus.reset());

  it("kaydetme sürerken ve hatada kirli bayraktan önce gelir; duraklatma yalnız kirliyken görünür", () => {
    expect(resolveSaveState(false, { phase: "idle" })).toBeNull();
    expect(resolveSaveState(true, { phase: "idle" })).toBe("dirty");
    expect(resolveSaveState(true, { phase: "saved", savedAt: 1 })).toBe("dirty");
    expect(resolveSaveState(false, { phase: "saved", savedAt: 1 })).toBe("saved");
    expect(resolveSaveState(true, { phase: "saving" })).toBe("saving");
    expect(resolveSaveState(true, { phase: "error" })).toBe("error");
    expect(resolveSaveState(true, { phase: "idle", paused: "uyarı" })).toBe("paused");
    expect(resolveSaveState(false, { phase: "idle", paused: "uyarı" })).toBeNull();
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

  it("yazma hatası (SaveError) nedeniyle rozete yansır, akış null ile sürer", async () => {
    const r = await trackSave(true, async () => Promise.reject(new SaveError("ENOSPC", "Disk dolu. Diskte yer açın.")));
    expect(r).toBeNull();
    expect(saveStatus.get()).toMatchObject({ phase: "error", error: "Disk dolu", errorDetail: "Disk dolu. Diskte yer açın." });
    const r2 = await trackSave(false, async () => Promise.reject(new SaveError("EBUSY")));
    expect(r2).toBeNull();
    expect(saveStatus.get().error).toBe("Dosya başka bir programda açık");
  });

  it("elle kayıt açılış bekletmesini kaldırır; yeni belge yüklenince bekletme yeniden başlar", async () => {
    saveStatus.reset();
    const s0 = saveStatus.get();
    expect(s0.manualGen).not.toBe(s0.gen);
    await trackSave(true, async () => ({ ok: true }));
    expect(saveStatus.get().manualGen).toBe(saveStatus.get().gen);
    saveStatus.reset();
    expect(saveStatus.get().manualGen).not.toBe(saveStatus.get().gen);
  });
});
