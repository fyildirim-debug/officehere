// Güvenli yazma yolu: eşzamanlı yazmalar, disk dolu, yazma ortasında süreç ölümü,
// Windows dosya kilidi, izinler. Ölçüm (öncesi): /tmp/olcum/araclar/kaydet_sim.mjs —
// iki eşzamanlı kayıtta 200 denemenin 199'unda dosya bozuluyordu.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { chmodSync, existsSync, mkdtempSync, promises as fsp, readdirSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { vi } from "vitest";
import { enqueue, isTempName, pendingWrites, queueKey, safeWriteFile, SafeWriteError, tempPathFor, writeAtomicUnqueued, type SafeFs } from "../src/main/safeWrite";

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), "oh-guvenli-"));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const hash = (b: Uint8Array) => createHash("sha1").update(b).digest("hex");
/** Belirlenimli sözde rastgele içerik. */
function bytes(n: number, seed: number): Buffer {
  const b = Buffer.alloc(n);
  let s = seed;
  for (let i = 0; i < n; i++) {
    s = (s * 1103515245 + 12345) & 0x7fffffff;
    b[i] = s >> 16;
  }
  return b;
}
const temps = () => readdirSync(dir).filter((f) => f.includes(".tmp"));

/** Geçici dosyaya yazarken belli bir bayttan sonra ENOSPC veren dosya sistemi (disk dolu). */
function diskFullFs(limit: number): SafeFs {
  return {
    ...fsp,
    open: (async (p: string, flags?: string, mode?: number) => {
      const fh = await fsp.open(p, flags, mode);
      if (!isTempName(path.basename(String(p)))) return fh;
      return new Proxy(fh, {
        get(t, k) {
          if (k === "writeFile")
            return async (d: Uint8Array) => {
              await t.write(d.subarray(0, Math.min(limit, d.byteLength)));
              if (d.byteLength > limit) throw Object.assign(new Error("ENOSPC: no space left on device, write"), { code: "ENOSPC" });
            };
          const v = Reflect.get(t, k) as unknown;
          return typeof v === "function" ? (v as (...a: unknown[]) => unknown).bind(t) : v;
        },
      });
    }) as SafeFs["open"],
  } as SafeFs;
}

// Eşzamanlı yazma ve gerçek süreç öldürme denemeleri Windows/macOS CI'ında yavaş: 5 sn varsayılan yetmez.
vi.setConfig({ testTimeout: 60_000 });

describe("güvenli yazma", () => {
  it("geçici ad benzersiz ve hedefle aynı dizinde", () => {
    const a = tempPathFor("/ev/rapor.docx");
    const b = tempPathFor("/ev/rapor.docx");
    expect(a).not.toBe(b);
    expect(path.dirname(a)).toBe(path.join("/ev"));
    expect(isTempName(path.basename(a), "rapor.docx")).toBe(true);
    expect(isTempName("rapor.docx.oh-12022.tmp")).toBe(false);
  });

  it("yeni dosya oluşturur, var olanın yerine geçer, geçici dosya bırakmaz", async () => {
    const f = path.join(dir, "a.docx");
    await safeWriteFile(f, bytes(1000, 1));
    expect(hash(readFileSync(f))).toBe(hash(bytes(1000, 1)));
    await safeWriteFile(f, bytes(50, 2));
    expect(hash(readFileSync(f))).toBe(hash(bytes(50, 2)));
    expect(temps()).toEqual([]);
  });

  // 200 sıralı atomik yazma (fsync + rename): Windows koşucusunda 60 sn'yi aşabiliyor (yandaki yarış testi gibi 120 sn)
  it("aynı dosyaya eşzamanlı 200 yazma: hiçbiri bozulmaz, sıra korunur, hepsi başarılı", async () => {
    const f = path.join(dir, "yaris.xlsx");
    writeFileSync(f, bytes(1000, 5));
    const datas = Array.from({ length: 200 }, (_, i) => bytes(1024 + ((i * 7919) % 200) * 1024, i + 10));
    const results = await Promise.allSettled(datas.map((d) => safeWriteFile(f, d)));
    expect(results.filter((r) => r.status === "rejected")).toEqual([]);
    // Kuyruk sırası korunur: son istenen yazma kazanır ve dosya bütündür.
    expect(hash(readFileSync(f))).toBe(hash(datas[199]));
    expect(temps()).toEqual([]);
  }, 120_000);

  it("otomatik kaydetme + Kaydet yarışı (kaydet_sim senaryosu, 200 deneme): dosya hep A ya da B, asla karışık", async () => {
    const f = path.join(dir, "yaris2.xlsx");
    const A = bytes(3 * 1024 * 1024, 3);
    const B = bytes(1024 * 1024, 4);
    const ha = hash(A);
    const hb = hash(B);
    let bozuk = 0;
    let basarisiz = 0;
    for (let i = 0; i < 200; i++) {
      writeFileSync(f, bytes(1000, 5));
      const r = await Promise.allSettled([safeWriteFile(f, A), safeWriteFile(f, B)]);
      basarisiz += r.filter((x) => x.status === "rejected").length;
      const h = hash(readFileSync(f));
      if (h !== ha && h !== hb) bozuk++;
      else expect(h).toBe(hb);
    }
    expect({ bozuk, basarisiz, gecici: temps().length }).toEqual({ bozuk: 0, basarisiz: 0, gecici: 0 });
  }, 120_000);

  it("disk dolu: özgün dosya bütün kalır, geçici dosya kalmaz, hata kodu iletilir", async () => {
    const f = path.join(dir, "belge.docx");
    const orig = bytes(60 * 1024, 1);
    writeFileSync(f, orig);
    const err = await safeWriteFile(f, bytes(120 * 1024, 2), { fs: diskFullFs(40 * 1024) }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(SafeWriteError);
    expect((err as SafeWriteError).code).toBe("ENOSPC");
    expect((err as SafeWriteError).step).toBe("write");
    expect(hash(readFileSync(f))).toBe(hash(orig));
    expect(temps()).toEqual([]);
  });

  it("Farklı Kaydet ile var olan dosyanın üzerine yazma da atomik: disk dolunca kesilmez", async () => {
    // Eski yol fs.writeFile ile doğrudan yazıyordu: 122.880 baytlık yeni içerik 106.496 baytta kesiliyordu.
    const f = path.join(dir, "uzerine.docx");
    const orig = bytes(60 * 1024, 7);
    writeFileSync(f, orig);
    await expect(safeWriteFile(f, bytes(122_880, 8), { fs: diskFullFs(106_496) })).rejects.toMatchObject({ code: "ENOSPC" });
    expect(readFileSync(f).length).toBe(orig.length);
    expect(hash(readFileSync(f))).toBe(hash(orig));
    expect(temps()).toEqual([]);
  });

  it("yazma ortasında süreç ölürse dosya ya eski ya tam yeni sürümdür (30 deneme)", async () => {
    const f = path.join(dir, "cokme.docx");
    const orig = bytes(200 * 1024, 7);
    const mod = path.resolve(__dirname, "../src/main/safeWrite.ts");
    let yarim = 0;
    let eski = 0;
    let yeni = 0;
    for (let i = 0; i < 30; i++) {
      writeFileSync(f, orig);
      const code = `import { safeWriteFile } from ${JSON.stringify(mod)};
        await safeWriteFile(${JSON.stringify(f)}, Buffer.alloc(32 * 1024 * 1024, 7));`;
      const c = spawn(process.execPath, ["--experimental-strip-types", "--no-warnings", "--input-type=module", "-e", code], { stdio: "ignore" });
      const exited = new Promise((r) => c.on("exit", r));
      await new Promise((r) => setTimeout(r, 40 + i * 6));
      c.kill("SIGKILL");
      await exited;
      const got = readFileSync(f);
      if (hash(got) === hash(orig)) eski++;
      else if (got.length === 32 * 1024 * 1024 && got.every((x) => x === 7)) yeni++;
      else yarim++;
    }
    expect(yarim).toBe(0);
    expect(eski + yeni).toBe(30);
  }, 120_000);

  it("Windows'ta kilitli dosya: taşıma birkaç kez denenir, sonra 'başka programda açık' (EBUSY)", async () => {
    const f = path.join(dir, "kilitli.docx");
    writeFileSync(f, "ÖZGÜN");
    let tries = 0;
    const lockedFs = { ...fsp, rename: async () => {
      tries++;
      throw Object.assign(new Error("EPERM: operation not permitted, rename"), { code: "EPERM" });
    } } as SafeFs;
    const err = (await safeWriteFile(f, Buffer.from("YENİ"), { fs: lockedFs, platform: "win32" }).catch((e: unknown) => e)) as SafeWriteError;
    expect(err.code).toBe("EBUSY");
    expect(err.step).toBe("replace");
    expect(tries).toBe(6);
    expect(readFileSync(f, "utf8")).toBe("ÖZGÜN");
    expect(temps()).toEqual([]);
  });

  it("geçici kilit kalkınca yeniden deneme başarılı olur", async () => {
    const f = path.join(dir, "kisa-kilit.docx");
    writeFileSync(f, "ÖZGÜN");
    let tries = 0;
    const flaky = { ...fsp, rename: async (a: string, b: string) => {
      if (++tries < 3) throw Object.assign(new Error("EBUSY"), { code: "EBUSY" });
      return fsp.rename(a, b);
    } } as SafeFs;
    await safeWriteFile(f, Buffer.from("YENİ"), { fs: flaky, platform: "win32" });
    expect(readFileSync(f, "utf8")).toBe("YENİ");
  });

  it.skipIf(process.platform === "win32" || process.getuid?.() === 0)("salt okunur dosyanın üzerine yazılmaz (EACCES)", async () => {
    const f = path.join(dir, "salt.docx");
    writeFileSync(f, "ÖZGÜN");
    chmodSync(f, 0o444);
    await expect(safeWriteFile(f, Buffer.from("YENİ"))).rejects.toMatchObject({ code: "EACCES" });
    expect(readFileSync(f, "utf8")).toBe("ÖZGÜN");
  });

  it.skipIf(process.platform === "win32")("izinler ve sembolik bağlantı korunur", async () => {
    const f = path.join(dir, "izin.docx");
    writeFileSync(f, "x");
    chmodSync(f, 0o640);
    await safeWriteFile(f, Buffer.from("yeni"));
    expect(statSync(f).mode & 0o777).toBe(0o640);
    const link = path.join(dir, "kisayol.docx");
    symlinkSync(f, link);
    await safeWriteFile(link, Buffer.from("bağlantıdan"));
    expect(readFileSync(f, "utf8")).toBe("bağlantıdan");
    expect((await fsp.lstat(link)).isSymbolicLink()).toBe(true);
  });

  it("klasör yoksa ENOENT, hedef klasörse EISDIR", async () => {
    await expect(safeWriteFile(path.join(dir, "yok", "a.docx"), Buffer.from("x"))).rejects.toMatchObject({ code: "ENOENT" });
    await expect(safeWriteFile(dir, Buffer.from("x"))).rejects.toMatchObject({ code: "EISDIR" });
  });

  it("yedek adımı başarısızsa yazma iptal edilir", async () => {
    const f = path.join(dir, "yedeksiz.docx");
    writeFileSync(f, "ÖZGÜN");
    await expect(writeAtomicUnqueued(f, Buffer.from("YENİ"), { beforeReplace: async () => Promise.reject(new Error("dolu")) })).rejects.toMatchObject({ code: "EBACKUP", step: "backup" });
    expect(readFileSync(f, "utf8")).toBe("ÖZGÜN");
    expect(temps()).toEqual([]);
  });

  it("artık geçici dosyalar temizlenir: yazanı ölmüş olanlar hemen, eskiler; çalışan sürecin taze dosyasına dokunulmaz", async () => {
    const f = path.join(dir, "artik.docx");
    writeFileSync(f, "x");
    const old = tempPathFor(f);
    const fresh = tempPathFor(f);
    // Çalışmayan bir sürecin (pid 2^22+) yarım bıraktığı dosya.
    const dead = path.join(dir, "artik.docx.oh-4194999-1-0123abcd.tmp");
    writeFileSync(old, "eski");
    writeFileSync(fresh, "taze");
    writeFileSync(dead, "ölü");
    const past = new Date(Date.now() - 3600_000);
    await fsp.utimes(old, past, past);
    await safeWriteFile(f, Buffer.from("y"));
    expect(existsSync(old)).toBe(false);
    expect(existsSync(dead)).toBe(false);
    expect(existsSync(fresh)).toBe(true);
  });

  it("kuyruk: önceki iş başarısız olsa da sonraki çalışır; yollar büyük-küçük harf duyarsız dosya sistemlerinde birleşir", async () => {
    const order: string[] = [];
    const a = enqueue("/x/a", async () => {
      await new Promise((r) => setTimeout(r, 20));
      order.push("1");
      throw new Error("bir");
    });
    const b = enqueue("/x/a", async () => void order.push("2"));
    await expect(a).rejects.toThrow("bir");
    await b;
    await pendingWrites();
    expect(order).toEqual(["1", "2"]);
    expect(queueKey("C:\\Belgeler\\Rapor.DOCX", "win32")).toBe(queueKey("c:\\belgeler\\rapor.docx", "win32"));
    expect(queueKey("/ev/A.docx", "linux")).not.toBe(queueKey("/ev/a.docx", "linux"));
  });
});
