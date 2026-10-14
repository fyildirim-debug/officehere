// Belge yazmalarının tek giriş noktası: güvenli yazma + yedekler.
// Ana süreçteki IPC işleyicileri (Kaydet, otomatik kaydetme, Farklı Kaydet,
// PDF, sürüm geri yükleme) yalnız bu sınıfı kullanır. Electron'a bağlı değildir.
import type { BackupStore } from "./backups";
import { enqueue, queueKey, SafeWriteError, writeAtomicUnqueued, type SafeFs } from "./safeWrite";

export interface WriteRequest {
  /** Yazan pencere oturumu: "önceki sürüm" her oturumda dosyaya ilk yazmadan önce alınır. */
  session: string | number;
  /** Otomatik kaydetme: "<ad>.bak" yan yedeği yoksa önce o alınır (alınamazsa yazılmaz). */
  sidecar?: boolean;
  /** Otomatik kaydetme doğrulaması: üzerine yazılan içerik "onceki" yuvasında tutulur. */
  keepPrevious?: boolean;
  /** Önceki sürüm: "first" (oturumda ilk yazma, varsayılan), "never". */
  snapshot?: "first" | "never";
}

export interface SaveServiceOptions {
  fs?: SafeFs;
  platform?: NodeJS.Platform;
  /** Kaydı engellemeyen sorunlar (sürüm alınamadı vb.) için günlük. */
  warn?: (message: string, err?: unknown) => void;
}

export class SaveService {
  private readonly snapshotted = new Set<string>();
  constructor(
    readonly store: BackupStore,
    private readonly opts: SaveServiceOptions = {},
  ) {}

  private sessionKey(session: string | number, target: string) {
    return `${session}\u0000${queueKey(target)}`;
  }

  /** Dosyayı yedekleriyle birlikte atomik yazar; başarısızlıkta SafeWriteError fırlatır. */
  write(target: string, data: Uint8Array, req: WriteRequest): Promise<void> {
    return enqueue(target, () =>
      writeAtomicUnqueued(target, data, {
        fs: this.opts.fs,
        platform: this.opts.platform,
        beforeReplace: async (real) => {
          if (req.sidecar) {
            try {
              await this.store.ensureSidecar(real);
            } catch (err) {
              throw new SafeWriteError("EBACKUP", "backup", err);
            }
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
    );
  }

  /** Otomatik kaydetmenin son yazmasından önceki içeriği geri yazar (doğrulama başarısızsa). */
  revertPrevious(target: string): Promise<void> {
    return enqueue(target, async () => {
      const prev = await this.store.readPrevious(target);
      if (!prev) throw new SafeWriteError("ENOENT", "backup");
      await writeAtomicUnqueued(target, prev, { fs: this.opts.fs, platform: this.opts.platform });
    });
  }

  /** Önceki bir sürümü geri yükler; geçerli içerik de önce yeni bir sürüm olarak saklanır (geri alınabilir). */
  restoreVersion(target: string, id: string): Promise<void> {
    return enqueue(target, async () => {
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
    });
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
