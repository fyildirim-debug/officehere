// Tek güvenli yazma yolu: belge kaydı (Kaydet, otomatik kaydetme, Farklı Kaydet),
// PDF dışa aktarma ve yedekler hep buradan geçer.
//
//   1. Aynı dizinde benzersiz adlı geçici dosya (pid + sayaç + rastgele) "wx" ile açılır.
//   2. Veri yazılır, dosya fsync edilir, kapatılır (izinleri özgün dosyadan alır).
//   3. İstenirse (beforeReplace) özgün dosyanın yedekleri alınır.
//   4. Geçici dosya hedefin yerine taşınır (rename: ya eski ya yeni içerik, asla yarım dosya).
//   5. Dizin fsync edilir (taşıma kalıcı olsun; Windows'ta yapılamaz, yok sayılır).
// Herhangi bir adımda hata olursa geçici dosya silinir ve özgün dosyaya dokunulmaz.
//
// Aynı dosyaya yazmalar yol başına bir kuyrukta sıraya girer: otomatik kaydetme,
// elle Kaydet, Farklı Kaydet ve farklı pencereler birbirinin yazmasını bölemez.
// Bu modül Electron'a bağlı değildir (birim testleri doğrudan çalıştırır).
import { constants, promises as nodeFs } from "node:fs";
import { randomBytes } from "node:crypto";
import path from "node:path";

/** Kullanılan dosya sistemi işlemleri (testlerde hata enjekte etmek için değiştirilebilir). */
export type SafeFs = Pick<typeof nodeFs, "open" | "rename" | "unlink" | "stat" | "realpath" | "chmod" | "readdir" | "copyFile" | "link" | "readFile" | "access">;

export type WriteStep = "write" | "backup" | "replace";

/** Yazma hatası: işletim sistemi hata kodu (ENOSPC, EACCES…) ve hangi adımda olduğu. */
export class SafeWriteError extends Error {
  readonly code: string;
  readonly step: WriteStep;
  constructor(code: string, step: WriteStep, cause?: unknown) {
    super(`${code} (${step})`);
    this.name = "SafeWriteError";
    this.code = code;
    this.step = step;
    if (cause !== undefined) (this as { cause?: unknown }).cause = cause;
  }
}

export function errorCode(err: unknown): string {
  if (err instanceof SafeWriteError) return err.code;
  const code = (err as { code?: unknown } | null)?.code;
  return typeof code === "string" && code ? code : "EUNKNOWN";
}

let counter = 0;

/** Hedefle aynı dizinde benzersiz geçici dosya adı. */
export function tempPathFor(target: string): string {
  counter = (counter + 1) % 1_000_000_000;
  return path.join(path.dirname(target), `${path.basename(target)}.oh-${process.pid}-${counter}-${randomBytes(4).toString("hex")}.tmp`);
}

/** Bu modülün geçici dosyası mı (ör. yarıda kalmış bir yazmadan artakalan). */
export function isTempName(name: string, base?: string): boolean {
  const m = /^(.*)\.oh-\d+-\d+-[0-9a-f]{8}\.tmp$/.exec(name);
  return !!m && (base === undefined || m[1] === base);
}

/** Kuyruk anahtarı: büyük-küçük harf duyarsız dosya sistemlerinde (Windows, macOS) küçük harf. */
export function queueKey(p: string, platform: NodeJS.Platform = process.platform): string {
  const r = path.resolve(p);
  // Dosya sistemi büyük-küçük harf eşlemesi yerel ayardan bağımsızdır (Türkçe İ/ı kuralı uygulanmaz).
  return platform === "win32" || platform === "darwin" ? r.toLowerCase() : r; // yerel-ayar: ascii
}

const queues = new Map<string, Promise<unknown>>();

/** İşi aynı yoldaki önceki işler bitince çalıştırır (önceki iş başarısız olsa bile). */
export function enqueue<T>(target: string, job: () => Promise<T>): Promise<T> {
  const key = queueKey(target);
  const prev = queues.get(key) ?? Promise.resolve();
  const run = prev.then(job, job);
  const tail = run.then(
    () => undefined,
    () => undefined,
  );
  queues.set(key, tail);
  void tail.then(() => {
    if (queues.get(key) === tail) queues.delete(key);
  });
  return run;
}

/** Süren ya da sırada bekleyen yazma var mı. */
export function hasPendingWrites(): boolean {
  return queues.size > 0;
}

/** Bekleyen yazması olan yollar (kapanışta beklemek ve testler için). */
export function pendingWrites(): Promise<void> {
  return Promise.all([...queues.values()]).then(() => undefined);
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Windows'ta virüs tarayıcısı/dizinleyici dosyayı kısa süre kilitleyebilir: taşıma birkaç kez denenir. */
async function renameWithRetry(fs: SafeFs, from: string, to: string, platform: NodeJS.Platform): Promise<void> {
  const delays = platform === "win32" ? [30, 60, 120, 240, 480] : [];
  for (let i = 0; ; i++) {
    try {
      await fs.rename(from, to);
      return;
    } catch (err) {
      const code = errorCode(err);
      if (i >= delays.length || !["EPERM", "EACCES", "EBUSY"].includes(code)) {
        // Windows'ta geçici dosya aynı dizine yazılabildiği hâlde taşıma reddediliyorsa hedef başka programda açıktır.
        throw new SafeWriteError(platform === "win32" && (code === "EPERM" || code === "EACCES") ? "EBUSY" : code, "replace", err);
      }
      await sleep(delays[i]);
    }
  }
}

/** Dizin girdisini kalıcı kılar (POSIX). Desteklenmezse sessizce geçilir. */
async function syncDir(fs: SafeFs, dir: string, platform: NodeJS.Platform): Promise<void> {
  if (platform === "win32") return;
  let fh: Awaited<ReturnType<SafeFs["open"]>> | undefined;
  try {
    fh = await fs.open(dir, "r");
    await fh.sync();
  } catch {
    // Bazı dosya sistemleri dizin fsync'ini desteklemez.
  } finally {
    await fh?.close().catch(() => undefined);
  }
}

/** Yarıda kalmış eski geçici dosyaları temizlenmiş hedefler (oturumda bir kez bakılır). */
const swept = new Set<string>();

/** Geçici dosyayı yazan süreç hâlâ çalışıyor mu (adındaki pid'den). */
function writerAlive(name: string): boolean {
  const pid = Number(/\.oh-(\d+)-/.exec(name)?.[1]);
  if (!pid || pid === process.pid) return pid === process.pid;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return errorCode(err) === "EPERM";
  }
}

/**
 * Aynı hedefe ait artık geçici dosyaları siler (çökme/elektrik kesintisi artığı): yazan süreci
 * artık çalışmayanlar hemen, diğerleri 10 dakikadan eskiyse.
 */
async function sweepStale(fs: SafeFs, target: string): Promise<void> {
  const key = queueKey(target);
  if (swept.has(key)) return;
  swept.add(key);
  try {
    const dir = path.dirname(target);
    const base = path.basename(target);
    const now = Date.now();
    for (const name of await fs.readdir(dir)) {
      if (!isTempName(name, base)) continue;
      const p = path.join(dir, name);
      try {
        if (!writerAlive(name) || now - (await fs.stat(p)).mtimeMs > 10 * 60_000) await fs.unlink(p);
      } catch {
        // başka bir işlem silmiş olabilir
      }
    }
  } catch {
    // dizin okunamıyorsa temizlik atlanır
  }
}

export interface SafeWriteOptions {
  fs?: SafeFs;
  platform?: NodeJS.Platform;
  /**
   * Geçici dosya tamamen yazılıp diske indikten sonra, hedefin yerine geçmeden
   * hemen önce çağrılır (yalnız hedef varsa): özgün içeriğin yedekleri burada alınır.
   * Hata fırlatırsa yazma iptal edilir, hedefe dokunulmaz.
   */
  beforeReplace?: (realTarget: string) => Promise<void>;
}

/** Hedef dosyanın gerçek yolu (sembolik bağlantı korunur, gösterdiği dosya güncellenir) ve var olup olmadığı. */
async function resolveTarget(fs: SafeFs, target: string): Promise<{ real: string; mode?: number }> {
  try {
    const real = await fs.realpath(target);
    const st = await fs.stat(real);
    if (st.isDirectory()) throw new SafeWriteError("EISDIR", "write");
    // Salt okunur dosyanın üzerine (taşıma yoluyla da olsa) yazılmaz.
    try {
      await fs.access(real, constants.W_OK);
    } catch (err) {
      throw new SafeWriteError(errorCode(err) === "ENOENT" ? "ENOENT" : "EACCES", "write", err);
    }
    return { real, mode: st.mode & 0o7777 };
  } catch (err) {
    if (err instanceof SafeWriteError) throw err;
    return { real: path.resolve(target) };
  }
}

/**
 * Kuyruğu beklemeden tek bir atomik yazma. Genellikle `safeWriteFile` kullanılır;
 * bu işlev kuyruğun içinden (ör. sürüm geri yüklerken) çağrılmak içindir.
 */
export async function writeAtomicUnqueued(target: string, data: Uint8Array, opts: SafeWriteOptions = {}): Promise<void> {
  const fs = opts.fs ?? nodeFs;
  const platform = opts.platform ?? process.platform;
  const { real, mode } = await resolveTarget(fs, target);
  await sweepStale(fs, real);
  const tmp = tempPathFor(real);
  let fh: Awaited<ReturnType<SafeFs["open"]>> | undefined;
  let created = false;
  try {
    try {
      fh = await fs.open(tmp, "wx", mode ?? 0o666);
      created = true;
      await fh.writeFile(data);
      await fh.sync();
      await fh.close();
      fh = undefined;
      // umask'tan bağımsız olarak özgün dosyanın izinleri korunur.
      if (mode !== undefined && platform !== "win32") await fs.chmod(tmp, mode).catch(() => undefined);
    } catch (err) {
      throw err instanceof SafeWriteError ? err : new SafeWriteError(errorCode(err), "write", err);
    }
    if (opts.beforeReplace && mode !== undefined) {
      try {
        await opts.beforeReplace(real);
      } catch (err) {
        throw err instanceof SafeWriteError ? err : new SafeWriteError("EBACKUP", "backup", err);
      }
    }
    await renameWithRetry(fs, tmp, real, platform);
    created = false;
    await syncDir(fs, path.dirname(real), platform);
  } finally {
    if (fh) await fh.close().catch(() => undefined);
    if (created) await fs.unlink(tmp).catch(() => undefined);
  }
}

/**
 * Dosyayı atomik olarak yazar (yol başına kuyrukta). Başarısızlıkta
 * SafeWriteError fırlatır; özgün dosya bütün kalır, geçici dosya kalmaz.
 */
export function safeWriteFile(target: string, data: Uint8Array, opts: SafeWriteOptions = {}): Promise<void> {
  return enqueue(target, () => writeAtomicUnqueued(target, data, opts));
}

/**
 * `source` dosyasının içeriğini `dest`'e atomik kopyalar. `noClobber` ise `dest`
 * zaten varsa dokunulmaz (oturumlar arası özgün yedeğin ezilmemesi için).
 * Kopya yapıldıysa true döner.
 */
export async function copyFileAtomic(source: string, dest: string, opts: { noClobber?: boolean; fs?: SafeFs; platform?: NodeJS.Platform } = {}): Promise<boolean> {
  const fs = opts.fs ?? nodeFs;
  const platform = opts.platform ?? process.platform;
  if (opts.noClobber && (await exists(fs, dest))) return false;
  const tmp = tempPathFor(dest);
  let created = false;
  try {
    await fs.copyFile(source, tmp, constants.COPYFILE_EXCL | constants.COPYFILE_FICLONE);
    created = true;
    let fh: Awaited<ReturnType<SafeFs["open"]>> | undefined;
    try {
      fh = await fs.open(tmp, "r+");
      await fh.sync();
    } finally {
      await fh?.close().catch(() => undefined);
    }
    if (opts.noClobber) {
      try {
        // link(): hedef varsa EEXIST ile başarısız olur, yani başka bir yazma araya girse bile ezmez.
        await fs.link(tmp, dest);
        await syncDir(fs, path.dirname(dest), platform);
        return true;
      } catch (err) {
        const code = errorCode(err);
        if (code === "EEXIST") return false;
        // Sabit bağlantı desteklemeyen dosya sistemleri (FAT32/exFAT, bazı ağ paylaşımları).
        if (await exists(fs, dest)) return false;
      }
    }
    await renameWithRetry(fs, tmp, dest, platform);
    created = false;
    await syncDir(fs, path.dirname(dest), platform);
    return true;
  } finally {
    if (created) await fs.unlink(tmp).catch(() => undefined);
  }
}

async function exists(fs: SafeFs, p: string): Promise<boolean> {
  try {
    await fs.stat(p);
    return true;
  } catch {
    return false;
  }
}
