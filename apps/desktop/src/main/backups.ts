// Kaydetmeden önce alınan yedekler. İki katman:
//
// 1. Yan yedek "<ad>.<uzantı>.bak" (dosyanın yanında): OfficeHere'in bir dosyaya
//    KENDİLİĞİNDEN (otomatik kaydetme) ilk yazmasından önceki özgün içerik.
//    Yalnız yoksa oluşturulur; sonraki oturumlarda asla ezilmez. Kullanıcının
//    görebileceği ve kolayca geri adlandırabileceği tek dosya budur.
//
// 2. Önceki sürümler (kullanıcı veri dizininde, userData/yedekler/<özet>/):
//    her pencere oturumunda dosyaya ilk yazmadan (Kaydet, otomatik kaydetme ya da
//    Farklı Kaydet ile üzerine yazma) hemen önceki içerik. En çok 3 tanesi döner
//    biçimde tutulur ve Dosya › Bilgi › "Önceki sürümler"den geri yüklenebilir.
//    Kullanıcının klasörünü kirletmemek için bunlar dosyanın yanına konmaz.
//    Aynı dizinde "onceki" adlı tek bir yuva daha vardır: otomatik kaydetmenin
//    son yazmasından önceki içerik (yazılan dosya doğrulanamazsa geri yüklenir).
//
// Bu modül Electron'a bağlı değildir.
import { createHash, randomBytes } from "node:crypto";
import { promises as nodeFs } from "node:fs";
import path from "node:path";
import { copyFileAtomic, queueKey, type SafeFs } from "./safeWrite";

export const MAX_VERSIONS = 3;

export interface VersionInfo {
  /** Sürüm kimliği (yedek dizinindeki dosya adı). */
  id: string;
  /** Yedeğin alındığı an (ms). */
  savedAt: number;
  size: number;
}

/** Yan yedek yolu: "rapor.docx" → "rapor.docx.bak". */
export function sidecarPathFor(target: string): string {
  return `${target}.bak`;
}

const VERSION_RE = /^surum-(\d{13})-[0-9a-f]{6}(\.[^.]+)?$/;

export class BackupStore {
  private readonly fs: SafeFs & Pick<typeof nodeFs, "mkdir" | "writeFile" | "rm">;
  constructor(
    /** Kök dizin (userData/yedekler). */
    readonly root: string,
    opts: { fs?: SafeFs & Pick<typeof nodeFs, "mkdir" | "writeFile" | "rm">; maxVersions?: number } = {},
  ) {
    this.fs = opts.fs ?? nodeFs;
    this.maxVersions = opts.maxVersions ?? MAX_VERSIONS;
  }
  private readonly maxVersions: number;

  /** Dosyaya özgü yedek dizini (yolun özeti; dosya adları ve yollar dizin adına yazılmaz). */
  dirFor(target: string): string {
    return path.join(this.root, createHash("sha256").update(queueKey(target)).digest("hex").slice(0, 20));
  }

  private async prepareDir(target: string): Promise<string> {
    const dir = this.dirFor(target);
    await this.fs.mkdir(dir, { recursive: true });
    // Hangi dosyanın yedekleri olduğu (temizlik ve tanılama için).
    await this.fs.writeFile(path.join(dir, "kaynak.json"), JSON.stringify({ yol: target, guncellendi: Date.now() })).catch(() => undefined);
    return dir;
  }

  /** Yan yedeği yalnız yoksa oluşturur. Oluşturulduysa true. */
  ensureSidecar(target: string): Promise<boolean> {
    return copyFileAtomic(target, sidecarPathFor(target), { noClobber: true, fs: this.fs });
  }

  /** Dosyanın şu anki içeriğini yeni bir "önceki sürüm" olarak saklar; en eskileri siler. */
  async snapshot(target: string, now = Date.now()): Promise<VersionInfo> {
    const dir = await this.prepareDir(target);
    const ext = path.extname(target).toLowerCase().replace(/[^.a-z0-9]/g, ""); // yerel-ayar: ascii
    const id = `surum-${String(now).padStart(13, "0")}-${randomBytes(3).toString("hex")}${ext}`;
    await copyFileAtomic(target, path.join(dir, id), { fs: this.fs });
    const list = await this.list(target);
    for (const old of list.slice(this.maxVersions)) await this.fs.unlink(path.join(dir, old.id)).catch(() => undefined);
    const st = await this.fs.stat(path.join(dir, id));
    return { id, savedAt: now, size: st.size };
  }

  /** Önceki sürümler, en yeni önce. */
  async list(target: string): Promise<VersionInfo[]> {
    const dir = this.dirFor(target);
    let names: string[];
    try {
      names = await this.fs.readdir(dir);
    } catch {
      return [];
    }
    const out: VersionInfo[] = [];
    for (const name of names) {
      const m = VERSION_RE.exec(name);
      if (!m) continue;
      try {
        const st = await this.fs.stat(path.join(dir, name));
        out.push({ id: name, savedAt: Number(m[1]), size: st.size });
      } catch {
        // silinmiş olabilir
      }
    }
    return out.sort((a, b) => b.savedAt - a.savedAt || (a.id < b.id ? 1 : -1));
  }

  /** Bir sürümün yolu (kimlik doğrulanır; dizin dışına çıkılamaz). */
  versionPath(target: string, id: string): string | null {
    if (typeof id !== "string" || !VERSION_RE.test(id)) return null;
    return path.join(this.dirFor(target), id);
  }

  async readVersion(target: string, id: string): Promise<Buffer | null> {
    const p = this.versionPath(target, id);
    if (!p) return null;
    try {
      return await this.fs.readFile(p);
    } catch {
      return null;
    }
  }

  private previousPath(target: string): string {
    return path.join(this.dirFor(target), `onceki${path.extname(target).toLowerCase().replace(/[^.a-z0-9]/g, "")}`); // yerel-ayar: ascii
  }

  /** Otomatik kaydetmenin üzerine yazacağı içeriği "onceki" yuvasına kopyalar. */
  async keepPrevious(target: string): Promise<void> {
    await this.prepareDir(target);
    await copyFileAtomic(target, this.previousPath(target), { fs: this.fs });
  }

  async readPrevious(target: string): Promise<Buffer | null> {
    try {
      return await this.fs.readFile(this.previousPath(target));
    } catch {
      return null;
    }
  }

  /** Uzun süredir dokunulmamış dosyaların yedek dizinlerini siler (açılışta, en iyi çabayla). */
  async prune(maxAgeMs = 90 * 24 * 3600_000, now = Date.now()): Promise<number> {
    let removed = 0;
    let dirs: string[];
    try {
      dirs = await this.fs.readdir(this.root);
    } catch {
      return 0;
    }
    for (const d of dirs) {
      if (!/^[0-9a-f]{20}$/.test(d)) continue;
      try {
        const st = await this.fs.stat(path.join(this.root, d, "kaynak.json"));
        if (now - st.mtimeMs > maxAgeMs) {
          await this.fs.rm(path.join(this.root, d), { recursive: true, force: true });
          removed++;
        }
      } catch {
        // kaynak.json yoksa dokunulmaz
      }
    }
    return removed;
  }
}
