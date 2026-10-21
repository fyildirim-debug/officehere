// Yedekler: "<ad>.bak" yalnız yoksa oluşturulur (oturumlar arası ezilmez), önceki
// sürümler en çok 3 tane döner, doğrulama başarısızsa son yazmadan önceki içerik geri gelir.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, symlinkSync, utimesSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { BackupStore, sidecarPathFor } from "../src/main/backups";
import { SaveService } from "../src/main/saveService";

// Her yazma fsync + yeniden adlandırma yapar; Windows CI'ında virüs taraması dosyayı kilitleyince
// yeniden adlandırma tekrar denenir (sınırlı) ve tek test 5 sn varsayılanı aşabilir.
vi.setConfig({ testTimeout: 60_000 });

let dir: string;
let root: string;
beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), "oh-yedek-"));
  root = path.join(dir, "userData", "yedekler");
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const text = (p: string) => readFileSync(p, "utf8");
const B = (s: string) => Buffer.from(s);

describe("yan yedek (.bak)", () => {
  it("iki oturum: .bak özgün içeriği korur, ikinci oturumda ezilmez", async () => {
    const f = path.join(dir, "rapor.docx");
    writeFileSync(f, "ÖZGÜN");
    // 1. oturum: otomatik kaydetme iki kez yazar.
    const s1 = new SaveService(new BackupStore(root));
    await s1.write(f, B("OTURUM1-A"), { session: 1, sidecar: true });
    await s1.write(f, B("OTURUM1-B"), { session: 1, sidecar: true });
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
    // 2. oturum (uygulama yeniden başlatıldı: bellekteki durum yok).
    const s2 = new SaveService(new BackupStore(root));
    await s2.write(f, B("OTURUM2"), { session: 1, sidecar: true });
    expect(text(f)).toBe("OTURUM2");
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
  });

  it("elle Kaydet yan yedek oluşturmaz; yeni dosyada da oluşmaz", async () => {
    const f = path.join(dir, "elle.docx");
    writeFileSync(f, "ÖZGÜN");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("YENİ"), { session: 1 });
    expect(existsSync(sidecarPathFor(f))).toBe(false);
    const g = path.join(dir, "yeni.docx");
    await s.write(g, B("İLK"), { session: 1, sidecar: true });
    expect(existsSync(sidecarPathFor(g))).toBe(false);
  });

  it("var olan .bak (kullanıcının ya da önceki oturumun) hiç değiştirilmez", async () => {
    const f = path.join(dir, "var.docx");
    writeFileSync(f, "ŞİMDİKİ");
    writeFileSync(sidecarPathFor(f), "ESKİ YEDEK");
    const store = new BackupStore(root);
    expect(await store.ensureSidecar(f)).toBe(false);
    expect(text(sidecarPathFor(f))).toBe("ESKİ YEDEK");
  });

  it("yan yedek alınamazsa (klasör yazılamaz) belge yazılmaz: EBACKUP", async () => {
    const f = path.join(dir, "yedeksiz.docx");
    writeFileSync(f, "ÖZGÜN");
    const store = new BackupStore(root);
    store.ensureSidecar = async () => Promise.reject(Object.assign(new Error("EACCES"), { code: "EACCES" }));
    const s = new SaveService(store);
    await expect(s.write(f, B("YENİ"), { session: 1, sidecar: true })).rejects.toMatchObject({ code: "EBACKUP" });
    expect(text(f)).toBe("ÖZGÜN");
    expect(readdirSync(dir).filter((n) => n.endsWith(".tmp"))).toEqual([]);
  });
});

describe("önceki sürümler", () => {
  it("her pencere oturumunda ilk yazmadan önce bir sürüm; en çok 3 tane, en yeni önce", async () => {
    const f = path.join(dir, "surumlu.xlsx");
    writeFileSync(f, "S0");
    const store = new BackupStore(root);
    const s = new SaveService(store);
    for (let i = 1; i <= 5; i++) {
      // Her döngü ayrı bir pencere oturumu; oturum içinde ikinci yazma sürüm almaz.
      await s.write(f, B(`S${i}`), { session: i });
      await s.write(f, B(`S${i}b`), { session: i });
      await new Promise((r) => setTimeout(r, 2));
    }
    const list = await store.list(f);
    expect(list).toHaveLength(3);
    const contents = await Promise.all(list.map(async (v) => (await store.readVersion(f, v.id))!.toString()));
    expect(contents).toEqual(["S4b", "S3b", "S2b"]);
    // Kullanıcının klasörü kirlenmez: dosyanın yanında yalnız kendisi var.
    expect(readdirSync(dir).filter((n) => n !== "userData")).toEqual(["surumlu.xlsx"]);
  });

  it("Farklı Kaydet ile üzerine yazılan dosyanın da sürümü alınır; pencere kapanınca yeni oturum yeni sürüm alır", async () => {
    const f = path.join(dir, "farkli.pptx");
    writeFileSync(f, "ESKİ SUNU");
    const store = new BackupStore(root);
    const s = new SaveService(store);
    await s.write(f, B("YENİ SUNU"), { session: 7 });
    expect((await store.list(f)).length).toBe(1);
    s.forgetSession(7);
    await s.write(f, B("DAHA YENİ"), { session: 7 });
    expect((await store.list(f)).length).toBe(2);
  });

  it("sürüm geri yükleme: geçerli içerik önce sürüm olarak saklanır (geri alınabilir)", async () => {
    const f = path.join(dir, "geri.docx");
    writeFileSync(f, "BİR");
    const store = new BackupStore(root);
    const s = new SaveService(store);
    await s.write(f, B("İKİ"), { session: 1 });
    const [v] = await store.list(f);
    await s.restoreVersion(f, v.id);
    expect(text(f)).toBe("BİR");
    const all = await Promise.all((await store.list(f)).map(async (x) => (await store.readVersion(f, x.id))!.toString()));
    expect(all).toContain("İKİ");
  });

  it("geçersiz sürüm kimliğiyle dizin dışına çıkılamaz", async () => {
    const store = new BackupStore(root);
    expect(await store.versionPath("/a.docx", "../../etc/passwd")).toBeNull();
    expect(await store.readVersion("/a.docx", "surum-0000000000001-abcdef.docx")).toBeNull();
    const s = new SaveService(store);
    await expect(s.restoreVersion(path.join(dir, "a.docx"), "../x")).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("doğrulama başarısız: son otomatik kaydetmeden önceki içerik geri yazılır", async () => {
    const f = path.join(dir, "dogrula.docx");
    writeFileSync(f, "İYİ-1");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("İYİ-2"), { session: 1, sidecar: true, keepPrevious: true });
    await s.write(f, B("BOZUK"), { session: 1, sidecar: true, keepPrevious: true });
    await s.revertPrevious(f);
    expect(text(f)).toBe("İYİ-2");
    expect(text(sidecarPathFor(f))).toBe("İYİ-1");
  });

  // macOS'ta /var → /private/var, Windows'ta kısa adlar (RUNNER~1) ve bağlantılı klasörler: kaydetme gerçek
  // yolu kullanır; liste, okuma ve geri yükleme dosyanın açıldığı yolla da aynı sürümleri bulmalı.
  it.skipIf(process.platform === "win32")("sembolik bağlantılı klasör yoluyla açılan dosyanın sürümleri bulunur", async () => {
    const real = path.join(dir, "gercek");
    mkdirSync(real);
    const link = path.join(dir, "bag");
    symlinkSync(real, link, "dir");
    const viaLink = path.join(link, "rapor.docx");
    writeFileSync(viaLink, "BİR");
    const store = new BackupStore(root);
    const s = new SaveService(store);
    await s.write(viaLink, B("İKİ"), { session: 1, keepPrevious: true });
    for (const p of [viaLink, path.join(real, "rapor.docx")]) {
      const list = await store.list(p);
      expect(list).toHaveLength(1);
      expect((await store.readVersion(p, list[0].id))!.toString()).toBe("BİR");
    }
    await s.revertPrevious(viaLink);
    expect(text(viaLink)).toBe("BİR");
  });

  it("uzun süredir dokunulmayan dosyaların yedek dizinleri temizlenir", async () => {
    const f = path.join(dir, "eski.docx");
    writeFileSync(f, "x");
    const store = new BackupStore(root);
    await store.snapshot(f);
    const old = new Date(Date.now() - 200 * 24 * 3600_000);
    utimesSync(path.join(store.dirFor(realpathSync.native(f)), "kaynak.json"), old, old);
    expect(await store.prune()).toBe(1);
    expect(await store.list(f)).toEqual([]);
  });
});

describe("yan yedeğin temiz kapanışta silinmesi", () => {
  /** Önceki sürümlerin içerikleri (en yeni önce). */
  const versionTexts = async (store: BackupStore, f: string) => Promise.all((await store.list(f)).map(async (v) => (await store.readVersion(f, v.id))!.toString("utf8")));

  it("otomatik kaydedilen dosya: pencere temiz kapanınca .bak silinir, özgün içerik önceki sürümlerde kalır", async () => {
    const f = path.join(dir, "rapor.xlsx");
    writeFileSync(f, "ÖZGÜN");
    const store = new BackupStore(root);
    const s = new SaveService(store);
    await s.write(f, B("OTO-1"), { session: "7:1", sidecar: true, keepPrevious: true });
    await s.write(f, B("OTO-2"), { session: "7:1", sidecar: true, keepPrevious: true });
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
    // Yerel realpath: Windows kısa adı (RUNNER~1) uygulamadaki fs.promises.realpath gibi uzun ada çözülür.
    expect(await s.releaseSidecars(7, { clean: true, path: f })).toEqual([sidecarPathFor(realpathSync.native(f))]);
    expect(existsSync(sidecarPathFor(f))).toBe(false);
    expect(text(f)).toBe("OTO-2");
    expect(await versionTexts(store, f)).toContain("ÖZGÜN");
    // Sonraki oturum: yeni yedek o anki içerikten alınır ve yine temiz kapanışta silinir.
    await s.write(f, B("OTO-3"), { session: "8:1", sidecar: true });
    expect(text(sidecarPathFor(f))).toBe("OTO-2");
    await s.releaseSidecars("8", { clean: true, path: f });
    expect(existsSync(sidecarPathFor(f))).toBe(false);
  });

  it("özgün içerik önceki sürümlerde yoksa (sürüm alınamamıştı) silmeden önce oraya alınır", async () => {
    const f = path.join(dir, "surumsuz.docx");
    writeFileSync(f, "ÖZGÜN");
    const store = new BackupStore(root);
    const snapshot = store.snapshot.bind(store);
    store.snapshot = () => Promise.reject(new Error("EACCES"));
    const s = new SaveService(store);
    await s.write(f, B("OTO"), { session: 1, sidecar: true });
    expect(await store.list(f)).toEqual([]);
    store.snapshot = snapshot;
    await s.releaseSidecars(1, { clean: true, path: f });
    expect(existsSync(sidecarPathFor(f))).toBe(false);
    expect(await versionTexts(store, f)).toEqual(["ÖZGÜN"]);
  });

  it("kirli kapanış, çökme ya da başka belgedeki yedek: .bak kalır", async () => {
    const f = path.join(dir, "kirli.pptx");
    writeFileSync(f, "ÖZGÜN");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("OTO"), { session: 1, sidecar: true });
    expect(await s.releaseSidecars(1, { clean: false, path: f })).toEqual([]);
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
    // Pencerede artık başka bir belge açık: bu dosyanın kapanış durumu bilinmez.
    const g = path.join(dir, "onceki.pptx");
    writeFileSync(g, "ÖZGÜN-G");
    await s.write(g, B("OTO-G"), { session: 2, sidecar: true });
    expect(await s.releaseSidecars(2, { clean: true, path: path.join(dir, "baska.pptx") })).toEqual([]);
    expect(text(sidecarPathFor(g))).toBe("ÖZGÜN-G");
  });

  it("doğrulaması başarısız (geri alınan) otomatik kayıttan sonra .bak kalır", async () => {
    const f = path.join(dir, "dogrulanamadi.docx");
    writeFileSync(f, "ÖZGÜN");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("BOZUK"), { session: 1, sidecar: true, keepPrevious: true });
    await s.revertPrevious(f);
    expect(text(f)).toBe("ÖZGÜN");
    // Sonra elle kaydedilse bile bu süreçte silinmez.
    await s.write(f, B("ELLE"), { session: 1 });
    expect(await s.releaseSidecars(1, { clean: true, path: f })).toEqual([]);
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
  });

  it("son yazma başarısızsa .bak kalır; sonraki başarılı yazmadan sonra silinir", async () => {
    const f = path.join(dir, "basarisiz.xlsx");
    writeFileSync(f, "ÖZGÜN");
    const store = new BackupStore(root);
    const keep = store.keepPrevious.bind(store);
    store.keepPrevious = () => Promise.reject(Object.assign(new Error("ENOSPC"), { code: "ENOSPC" }));
    const s = new SaveService(store);
    await expect(s.write(f, B("OTO"), { session: 1, sidecar: true, keepPrevious: true })).rejects.toMatchObject({ code: "EBACKUP" });
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
    expect(await s.releaseSidecars(1, { clean: true, path: f })).toEqual([]);
    expect(existsSync(sidecarPathFor(f))).toBe(true);
    // Yeni süreç (yedek önceden var): hiç silinmez.
    store.keepPrevious = keep;
    const s2 = new SaveService(store);
    await s2.write(f, B("OTO-2"), { session: 1, sidecar: true, keepPrevious: true });
    expect(await s2.releaseSidecars(1, { clean: true, path: f })).toEqual([]);
    expect(text(sidecarPathFor(f))).toBe("ÖZGÜN");
  });

  it("önceden var olan ya da sonradan değiştirilen .bak silinmez", async () => {
    const f = path.join(dir, "kullanici.docx");
    writeFileSync(f, "ŞİMDİKİ");
    writeFileSync(sidecarPathFor(f), "KULLANICININ");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("OTO"), { session: 1, sidecar: true });
    expect(await s.releaseSidecars(1, { clean: true, path: f })).toEqual([]);
    expect(text(sidecarPathFor(f))).toBe("KULLANICININ");
    // Bu süreçte oluşturuldu ama kapanmadan önce başkası değiştirdi.
    const g = path.join(dir, "degisen.docx");
    writeFileSync(g, "ÖZGÜN");
    await s.write(g, B("OTO"), { session: 1, sidecar: true });
    writeFileSync(sidecarPathFor(g), "ELLE DEĞİŞTİRİLDİ");
    expect(await s.releaseSidecars(1, { clean: true, path: g })).toEqual([]);
    expect(text(sidecarPathFor(g))).toBe("ELLE DEĞİŞTİRİLDİ");
  });

  it("dosya başka pencerede de açıksa yedek o pencere kapanınca silinir", async () => {
    const f = path.join(dir, "iki-pencere.xlsx");
    writeFileSync(f, "ÖZGÜN");
    const s = new SaveService(new BackupStore(root));
    await s.write(f, B("OTO"), { session: "3:1", sidecar: true });
    expect(await s.releaseSidecars(3, { clean: true, path: f, stillOpen: () => true })).toEqual([]);
    expect(existsSync(sidecarPathFor(f))).toBe(true);
    // Öbür pencere (yedeği oluşturmayan) temiz kapanır.
    expect(await s.releaseSidecars(4, { clean: true, path: f })).toHaveLength(1);
    expect(existsSync(sidecarPathFor(f))).toBe(false);
  });
});
