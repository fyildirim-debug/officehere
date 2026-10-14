// Yedekler: "<ad>.bak" yalnız yoksa oluşturulur (oturumlar arası ezilmez), önceki
// sürümler en çok 3 tane döner, doğrulama başarısızsa son yazmadan önceki içerik geri gelir.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, symlinkSync, utimesSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { BackupStore, sidecarPathFor } from "../src/main/backups";
import { SaveService } from "../src/main/saveService";

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
