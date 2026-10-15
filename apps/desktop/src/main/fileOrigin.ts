// Dosya internetten mi geldi (tarayıcı indirmesi ya da e-posta eki)? Görüntüleyici kipi (Seçenekler ›
// "İndirilenler ve e-posta eklerinden gelen dosyaları salt okunur aç") bunu sorar.
// - Windows: tarayıcılar ve e-posta istemcileri dosyaya "Zone.Identifier" alternatif veri akışı
//   (Web İşareti) ekler; ZoneId 3 (İnternet) ve 4 (Kısıtlanmış siteler) internetten gelmiş sayılır.
// - macOS: indirilen ve e-postadan kaydedilen dosyada "com.apple.quarantine" özniteliği bulunur.
// - Linux: dosyaya işaret konmaz; XDG İndirilenler dizini (user-dirs.dirs, ayrıca ~/İndirilenler ve
//   ~/Downloads) ile Thunderbird'ün ekleri açtığı geçici dizin (/tmp/mozilla_<kullanıcı>0) esas alınır.
// Disk erişimi bağımlılık olarak verilir (birim testinde sahte dosya sistemi ve öznitelik).
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

/** Neden internetten gelmiş sayıldı (günlük ve testler için). */
export type InternetSource = "web-isareti" | "karantina" | "indirilenler" | "eposta-eki";

export interface OriginDeps {
  platform: NodeJS.Platform;
  home: string;
  tmpdir: string;
  env: Record<string, string | undefined>;
  /** Dosyayı okur; yoksa hata fırlatır (Windows'ta "yol:Zone.Identifier" akışı da). */
  readFile(p: string): Promise<Uint8Array>;
  /** macOS genişletilmiş özniteliği; yoksa null. */
  readXattr(p: string, name: string): Promise<string | null>;
}

/** Zone.Identifier akışındaki ZoneId; metin UTF-8 ya da (BOM'lu) UTF-16LE olabilir. Yoksa null. */
export function parseZoneId(data: Uint8Array): number | null {
  const text = data[0] === 0xff && data[1] === 0xfe ? new TextDecoder("utf-16le").decode(data.subarray(2)) : new TextDecoder().decode(data);
  const m = /^\s*ZoneId\s*=\s*(\d+)\s*$/im.exec(text);
  return m ? Number(m[1]) : null;
}

/** İnternet (3) ve Kısıtlanmış siteler (4); yerel, intranet ve güvenilen siteler değil. */
export function isInternetZone(zone: number | null): boolean {
  return zone === 3 || zone === 4;
}

/**
 * user-dirs.dirs içindeki XDG_DOWNLOAD_DIR: "$HOME/İndirilenler" ya da mutlak yol (xdg-user-dirs
 * biçimi). Satır yoksa, yol göreliyse ya da ev dizininin kendisiyse (dizin kapatılmış) null.
 */
export function parseUserDirs(text: string, home: string): string | null {
  for (const line of text.split(/\r?\n/)) {
    const m = /^\s*XDG_DOWNLOAD_DIR\s*=\s*"((?:[^"\\]|\\.)*)"\s*$/.exec(line);
    if (!m) continue;
    const value = m[1].replace(/\\(.)/g, "$1").replace(/^\$(?:HOME|\{HOME\})(?=\/|$)/, home);
    if (!path.posix.isAbsolute(value)) return null;
    const dir = path.posix.resolve(value);
    return dir === path.posix.resolve(home) ? null : dir;
  }
  return null;
}

/** Linux'ta İndirilenler sayılan dizinler: XDG ayarı (varsa) ile ~/İndirilenler ve ~/Downloads. */
export async function linuxDownloadDirs(deps: OriginDeps): Promise<string[]> {
  const xdgConfig = deps.env.XDG_CONFIG_HOME;
  const config = xdgConfig && path.posix.isAbsolute(xdgConfig) ? xdgConfig : path.posix.join(deps.home, ".config");
  const dirs: string[] = [];
  try {
    const dir = parseUserDirs(new TextDecoder().decode(await deps.readFile(path.posix.join(config, "user-dirs.dirs"))), deps.home);
    if (dir) dirs.push(dir);
  } catch {
    // Dosya yoksa yalnız alışılmış adlar.
  }
  for (const name of ["İndirilenler", "Downloads"]) dirs.push(path.posix.join(deps.home, name));
  return [...new Set(dirs)];
}

/** Dosya dizinin içinde mi (alt dizinler dahil; yazım biçimi farkı NFC ile giderilir). */
export function isInside(file: string, dir: string): boolean {
  const rel = path.posix.relative(path.posix.resolve(dir.normalize("NFC")), path.posix.resolve(file.normalize("NFC")));
  return rel !== "" && rel !== ".." && !rel.startsWith("../") && !path.posix.isAbsolute(rel);
}

/** Dosya internetten geldiyse nedeni, gelmediyse (ya da anlaşılamadıysa) null. */
export async function internetSource(filePath: string, deps: OriginDeps): Promise<InternetSource | null> {
  if (deps.platform === "win32") {
    try {
      return isInternetZone(parseZoneId(await deps.readFile(`${filePath}:Zone.Identifier`))) ? "web-isareti" : null;
    } catch {
      return null;
    }
  }
  if (deps.platform === "darwin") {
    try {
      return (await deps.readXattr(filePath, "com.apple.quarantine")) ? "karantina" : null;
    } catch {
      return null;
    }
  }
  // Thunderbird ekleri açmadan önce /tmp/mozilla_<kullanıcı>0 altına yazar.
  const inTmp = path.posix.relative(path.posix.resolve(deps.tmpdir), path.posix.resolve(filePath));
  if (/^mozilla_[^/]*\/./.test(inTmp)) return "eposta-eki";
  for (const dir of await linuxDownloadDirs(deps)) if (isInside(filePath, dir)) return "indirilenler";
  return null;
}

/** Gerçek dosya sistemi ve macOS'un xattr aracı. */
export function nodeOriginDeps(): OriginDeps {
  return {
    platform: process.platform,
    home: os.homedir(),
    tmpdir: os.tmpdir(),
    env: process.env,
    readFile: (p) => fs.readFile(p),
    readXattr: (p, name) =>
      new Promise((resolve) => {
        execFile("/usr/bin/xattr", ["-p", name, p], { timeout: 3000, encoding: "utf8" }, (err, stdout) => resolve(err ? null : stdout.trim() || null));
      }),
  };
}
