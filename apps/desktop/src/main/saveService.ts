// Belge yazmalarının tek giriş noktası: güvenli yazma + yedekler.
// Ana süreçteki IPC işleyicileri (Kaydet, otomatik kaydetme, Farklı Kaydet,
// PDF, sürüm geri yükleme) yalnız bu sınıfı kullanır. Electron'a bağlı değildir.
//
// Yan yedeğin ("<ad>.bak") ömrü: otomatik kaydetme onu yalnız yoksa oluşturur ve bu süreçte
// oluşturulanlar hatırlanır. Pencere düzgün kapanınca (kaydedilmemiş değişiklik yok, çökmedi)
// ve dosyaya son yazma başarılıysa yedeğin içeriği önceki sürümlerde olduğu doğrulanıp yedek
// silinir. Doğrulaması başarısız olan (geri alınan) otomatik kayıttan sonra, son yazma
// başarısızsa ya da pencere kirli kapanırsa yedek kalır. Önceden var olan (kullanıcının ya da
// önceki bir oturumun) .bak dosyasına hiç dokunulmaz.
import { promises as nodeFs } from "node:fs";
import { sidecarPathFor, type BackupStore } from "./backups";
import { enqueue, queueKey, SafeWriteError, writeAtomicUnqueued, type SafeFs } from "./safeWrite";

export interface WriteRequest {
  /** Yazan pencere oturumu: "önceki sürüm" her oturumda dosyaya ilk yazmadan önce alınır. */
  session: string | number;
  /** Otomatik kaydetme: "<ad>.bak" yan yedeği yoksa önce o alınır (alınamazsa yazılmaz); temiz kapanışta silinir. */
  sidecar?: boolean;
  /** Otomatik kaydetme doğrulaması: üzerine yazılan içerik "onceki" yuvasında tutulur. */
  keepPrevious?: boolean;
  /** Önceki sürüm: "first" (oturumda ilk yazma, varsayılan), "never". */
  snapshot?: "first" | "never";
}

/** Bu süreçte oluşturulan yan yedek. */
interface CreatedSidecar {
  /** Sayfanın verdiği yol (yazma kuyruğu ve pencere eşlemesi bununla). */
  path: string;
  /** Gerçek yol (yedek ve önceki sürümler bununla). */
  real: string;
  sidecar: string;
  /** Oluşturan pencere (oturum kimliğinin pencere kısmı). */
  owner: string;
  /** "ok": son yazma başarılı; "unsafe": son yazma başarısız; "keep": doğrulama başarısız oldu, yedek kalır. */
  state: "ok" | "unsafe" | "keep";
  /** Oluşturulduğu andaki boyut ve değişiklik zamanı: başkası değiştirdiyse silinmez. */
  size: number;
  mtimeMs: number;
}

/** Pencere kapanırken yan yedeklerin bırakılma koşulları. */
export interface ReleaseOptions {
  /** Pencere düzgün kapandı: kaydedilmemiş değişiklik yok ve sayfa çökmedi. */
  clean: boolean;
  /** Pencerenin kapanırken açık olan belgesi (yalnız onun yedeği silinir). */
  path?: string;
  /** Dosya başka bir pencerede hâlâ açık mı (açıksa yedek o pencere kapanana dek kalır). */
  stillOpen?: (path: string) => boolean;
}

export interface SaveServiceOptions {
  fs?: SafeFs;
  platform?: NodeJS.Platform;
  /** Kaydı engellemeyen sorunlar (sürüm alınamadı vb.) için günlük. */
  warn?: (message: string, err?: unknown) => void;
}

export class SaveService {
  private readonly snapshotted = new Set<string>();
  /** Bu süreçte oluşturulan yan yedekler (anahtar: sayfanın yolunun kuyruk anahtarı). */
  private readonly sidecars = new Map<string, CreatedSidecar>();
  constructor(
    readonly store: BackupStore,
    private readonly opts: SaveServiceOptions = {},
  ) {}

  private get fs(): SafeFs {
    return this.opts.fs ?? nodeFs;
  }

  private sessionKey(session: string | number, target: string) {
    return `${session}\u0000${queueKey(target)}`;
  }

  /** Son yazmanın sonucunu yan yedeğin kaydına işler ("keep" kalıcıdır). */
  private mark(target: string, state: "ok" | "unsafe") {
    const rec = this.sidecars.get(queueKey(target));
    if (rec && rec.state !== "keep") rec.state = state;
  }

  /** Bir yazmayı yürütür ve sonucunu yan yedeğin kaydına işler. */
  private async tracked(target: string, job: () => Promise<void>): Promise<void> {
    try {
      await job();
    } catch (err) {
      this.mark(target, "unsafe");
      throw err;
    }
    this.mark(target, "ok");
  }

  /** Dosyayı yedekleriyle birlikte atomik yazar; başarısızlıkta SafeWriteError fırlatır. */
  write(target: string, data: Uint8Array, req: WriteRequest): Promise<void> {
    return enqueue(target, () =>
      this.tracked(target, () =>
        writeAtomicUnqueued(target, data, {
          fs: this.opts.fs,
          platform: this.opts.platform,
          beforeReplace: async (real) => {
            if (req.sidecar) {
              let created: boolean;
              try {
                created = await this.store.ensureSidecar(real);
              } catch (err) {
                throw new SafeWriteError("EBACKUP", "backup", err);
              }
              if (created) await this.remember(target, real, req.session);
            }
            if (req.keepPrevious) {
              try {
                await this.store.keepPrevious(real);
              } catch (err) {
                throw new SafeWriteError("EBACKUP", "backup", err);
              }
            }
            const key = this.sessionKey(req.session, real);
            if ((req.snapshot ?? "first") === "first" && !this.snapshotted.has(key)) {
              try {
                await this.store.snapshot(real);
                this.snapshotted.add(key);
              } catch (err) {
                // Önceki sürüm bir güvenlik ağıdır; alınamaması kaydı engellemez.
                this.opts.warn?.("Önceki sürüm saklanamadı", err);
              }
            }
          },
        }),
      ),
    );
  }

  /** Yeni oluşturulan yan yedeği hatırlar (pencere düzgün kapanınca silinmek üzere). */
  private async remember(target: string, real: string, session: string | number) {
    const sidecar = sidecarPathFor(real);
    try {
      const st = await this.fs.stat(sidecar);
      this.sidecars.set(queueKey(target), { path: target, real, sidecar, owner: String(session).split(":")[0], state: "unsafe", size: st.size, mtimeMs: st.mtimeMs });
    } catch {
      // Ölçülemeyen yedek hatırlanmaz: silinmez, olduğu gibi kalır.
    }
  }

  /** Otomatik kaydetmenin son yazmasından önceki içeriği geri yazar (doğrulama başarısızsa). */
  revertPrevious(target: string): Promise<void> {
    // Doğrulanamayan bir kayıt oldu: özgün dosyanın yan yedeği bu süreçte silinmez.
    const rec = this.sidecars.get(queueKey(target));
    if (rec) rec.state = "keep";
    return enqueue(target, async () => {
      const prev = await this.store.readPrevious(target);
      if (!prev) throw new SafeWriteError("ENOENT", "backup");
      await writeAtomicUnqueued(target, prev, { fs: this.opts.fs, platform: this.opts.platform });
    });
  }

  /** Önceki bir sürümü geri yükler; geçerli içerik de önce yeni bir sürüm olarak saklanır (geri alınabilir). */
  restoreVersion(target: string, id: string): Promise<void> {
    return enqueue(target, () =>
      this.tracked(target, async () => {
        const data = await this.store.readVersion(target, id);
        if (!data) throw new SafeWriteError("ENOENT", "backup");
        await writeAtomicUnqueued(target, data, {
          fs: this.opts.fs,
          platform: this.opts.platform,
          beforeReplace: async (real) => {
            try {
              await this.store.snapshot(real);
            } catch (err) {
              this.opts.warn?.("Geri yüklemeden önce sürüm saklanamadı", err);
            }
          },
        });
      }),
    );
  }

  /**
   * Pencere kapandı: penceredeki belgenin bu süreçte oluşturulan yan yedeği, pencere düzgün kapandıysa ve
   * dosyaya son yazma başarılıysa silinir; silmeden önce yedeğin içeriği önceki sürümlere alınır (yoksa).
   * Dosya başka pencerede hâlâ açıksa yedek o pencere kapanana dek kalır. Silinen yedeklerin yolları döner.
   */
  async releaseSidecars(owner: string | number, opts: ReleaseOptions): Promise<string[]> {
    const id = String(owner);
    const current = opts.path ? queueKey(opts.path) : null;
    const jobs: Promise<string | null>[] = [];
    for (const [key, rec] of [...this.sidecars]) {
      if (rec.owner !== id && key !== current) continue;
      if (opts.stillOpen?.(rec.path)) continue;
      // Pencerede başka bir belge açıksa bu dosyanın kapanış durumu bilinmez: yedek kalır.
      if (!opts.clean || key !== current) {
        this.sidecars.delete(key);
        continue;
      }
      // Karar, dosyadaki bekleyen yazmalar bittikten sonra son yazmanın sonucuna göre verilir.
      jobs.push(
        enqueue(rec.path, async () => {
          if (this.sidecars.get(key) !== rec || opts.stillOpen?.(rec.path)) return null;
          this.sidecars.delete(key);
          return rec.state === "ok" ? this.removeSidecar(rec) : null;
        }),
      );
    }
    return (await Promise.all(jobs)).filter((p): p is string => p !== null);
  }

  private async removeSidecar(rec: CreatedSidecar): Promise<string | null> {
    try {
      const st = await this.fs.stat(rec.sidecar);
      // Yedek oluşturulduktan sonra değiştirildiyse (kullanıcı ya da başka bir program) dokunulmaz.
      if (st.size !== rec.size || st.mtimeMs !== rec.mtimeMs) return null;
      if (!(await this.store.ensureVersionOf(rec.real, rec.sidecar))) return null;
      await this.fs.unlink(rec.sidecar);
      return rec.sidecar;
    } catch (err) {
      this.opts.warn?.("Yan yedek silinemedi", err);
      return null;
    }
  }

  /**
   * Pencere kapandı: bir sonraki açılışta yeni bir "önceki sürüm" alınsın.
   * Oturum kimliği "<pencere>:<yükleme>" biçimindeyse pencere kimliği vermek yeter.
   */
  forgetSession(session: string | number): void {
    const id = String(session);
    for (const k of [...this.snapshotted]) {
      const s = k.slice(0, k.indexOf("\u0000"));
      if (s === id || s.startsWith(`${id}:`)) this.snapshotted.delete(k);
    }
  }
}
