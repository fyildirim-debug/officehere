import { describe, expect, it } from "vitest";
import { internetSource, isInside, isInternetZone, linuxDownloadDirs, parseUserDirs, parseZoneId, type OriginDeps } from "../src/main/fileOrigin";

// Kaynak tespiti üç platformda, sahte dosya sistemi ve öznitelikle sınanır.

const enc = new TextEncoder();

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(2 + text.length * 2);
  out[0] = 0xff;
  out[1] = 0xfe;
  for (let i = 0; i < text.length; i++) {
    out[2 + i * 2] = text.charCodeAt(i) & 0xff;
    out[3 + i * 2] = text.charCodeAt(i) >> 8;
  }
  return out;
}

/** Bellekte dosyalar (yol → içerik) ve öznitelikler (yol → ad → değer). */
function deps(over: Partial<OriginDeps> & { files?: Record<string, string | Uint8Array>; xattrs?: Record<string, Record<string, string>> } = {}): OriginDeps & { reads: string[] } {
  const reads: string[] = [];
  const { files = {}, xattrs = {}, ...rest } = over;
  return {
    reads,
    platform: "linux",
    home: "/home/ayse",
    tmpdir: "/tmp",
    env: {},
    async readFile(p) {
      reads.push(p);
      const v = files[p];
      if (v === undefined) throw Object.assign(new Error(`ENOENT: ${p}`), { code: "ENOENT" });
      return typeof v === "string" ? enc.encode(v) : v;
    },
    async readXattr(p, name) {
      return xattrs[p]?.[name] ?? null;
    },
    ...rest,
  };
}

describe("Windows: Zone.Identifier (Web İşareti)", () => {
  const file = "C:\\Users\\ayse\\Downloads\\rapor.docx";
  const ads = `${file}:Zone.Identifier`;

  it("ZoneId okunur: UTF-8 ve BOM'lu UTF-16LE, büyük-küçük harf ve boşluk farkı önemsiz", () => {
    expect(parseZoneId(enc.encode("[ZoneTransfer]\r\nZoneId=3\r\nReferrerUrl=https://ornek.com/\r\nHostUrl=https://ornek.com/rapor.docx\r\n"))).toBe(3);
    expect(parseZoneId(utf16le("[ZoneTransfer]\r\nZoneId=3\r\n"))).toBe(3);
    expect(parseZoneId(enc.encode("[ZoneTransfer]\nzoneid = 4\n"))).toBe(4);
    expect(parseZoneId(enc.encode("[ZoneTransfer]\r\nHostUrl=about:internet\r\n"))).toBeNull();
    expect(parseZoneId(enc.encode(""))).toBeNull();
  });

  it("yalnız İnternet (3) ve Kısıtlanmış siteler (4) internetten sayılır", () => {
    expect([0, 1, 2, 3, 4, null].map(isInternetZone)).toEqual([false, false, false, true, true, false]);
  });

  it("indirilen dosya (ZoneId=3) internetten; yerel ağdan (ZoneId=1) ya da işaretsiz dosya değil", async () => {
    expect(await internetSource(file, deps({ platform: "win32", files: { [ads]: "[ZoneTransfer]\r\nZoneId=3\r\n" } }))).toBe("web-isareti");
    expect(await internetSource(file, deps({ platform: "win32", files: { [ads]: utf16le("[ZoneTransfer]\r\nZoneId=4\r\n") } }))).toBe("web-isareti");
    expect(await internetSource(file, deps({ platform: "win32", files: { [ads]: "[ZoneTransfer]\r\nZoneId=1\r\n" } }))).toBeNull();
    const d = deps({ platform: "win32" });
    expect(await internetSource(file, d)).toBeNull();
    expect(d.reads).toEqual([ads]);
  });

  it("Windows'ta dizin adı tek başına yetmez (işaretsiz dosya İndirilenler'de olsa da)", async () => {
    expect(await internetSource("C:\\Users\\ayse\\Downloads\\kendi.docx", deps({ platform: "win32", home: "C:\\Users\\ayse" }))).toBeNull();
  });
});

describe("macOS: com.apple.quarantine", () => {
  const file = "/Users/ayse/Downloads/sunu.pptx";

  it("karantina özniteliği olan dosya internetten (tarayıcı ya da Mail eki)", async () => {
    expect(await internetSource(file, deps({ platform: "darwin", xattrs: { [file]: { "com.apple.quarantine": "0083;66f1a2b3;Safari;9C2E" } } }))).toBe("karantina");
    expect(await internetSource(file, deps({ platform: "darwin", xattrs: { [file]: { "com.apple.quarantine": "0081;66f1a2b3;Mail;" } } }))).toBe("karantina");
  });

  it("özniteliği olmayan ya da okunamayan dosya internetten sayılmaz", async () => {
    expect(await internetSource(file, deps({ platform: "darwin" }))).toBeNull();
    expect(await internetSource(file, deps({ platform: "darwin", xattrs: { [file]: { "com.apple.metadata:kMDItemWhereFroms": "x" } } }))).toBeNull();
    const failing = deps({
      platform: "darwin",
      readXattr: async () => {
        throw new Error("xattr yok");
      },
    });
    expect(await internetSource(file, failing)).toBeNull();
  });
});

describe("Linux: XDG İndirilenler dizini", () => {
  it("user-dirs.dirs: $HOME ve mutlak yol; ev dizininin kendisi ve göreli yol geçersiz", () => {
    const home = "/home/ayse";
    expect(parseUserDirs('# xdg-user-dirs\nXDG_DESKTOP_DIR="$HOME/Masaüstü"\nXDG_DOWNLOAD_DIR="$HOME/İndirilenler"\n', home)).toBe("/home/ayse/İndirilenler");
    expect(parseUserDirs('XDG_DOWNLOAD_DIR="${HOME}/Downloads"', home)).toBe("/home/ayse/Downloads");
    expect(parseUserDirs('XDG_DOWNLOAD_DIR="/veri/indirilen"', home)).toBe("/veri/indirilen");
    expect(parseUserDirs('XDG_DOWNLOAD_DIR="$HOME/Benim \\"İndirilenlerim\\""', home)).toBe('/home/ayse/Benim "İndirilenlerim"');
    expect(parseUserDirs('XDG_DOWNLOAD_DIR="$HOME/"', home)).toBeNull();
    expect(parseUserDirs('XDG_DOWNLOAD_DIR="indirilen"', home)).toBeNull();
    expect(parseUserDirs("XDG_MUSIC_DIR=\"$HOME/Müzik\"\n", home)).toBeNull();
  });

  it("dizinler: XDG_CONFIG_HOME'daki ayar, yoksa ~/.config; her durumda ~/İndirilenler ve ~/Downloads", async () => {
    const withXdg = deps({ env: { XDG_CONFIG_HOME: "/home/ayse/ayar" }, files: { "/home/ayse/ayar/user-dirs.dirs": 'XDG_DOWNLOAD_DIR="/veri/indirilen"' } });
    expect(await linuxDownloadDirs(withXdg)).toEqual(["/veri/indirilen", "/home/ayse/İndirilenler", "/home/ayse/Downloads"]);
    expect(withXdg.reads).toEqual(["/home/ayse/ayar/user-dirs.dirs"]);
    const plain = deps({ files: { "/home/ayse/.config/user-dirs.dirs": 'XDG_DOWNLOAD_DIR="$HOME/İndirilenler"' } });
    expect(await linuxDownloadDirs(plain)).toEqual(["/home/ayse/İndirilenler", "/home/ayse/Downloads"]);
    // Göreli XDG_CONFIG_HOME belirtime göre yok sayılır.
    const relative = deps({ env: { XDG_CONFIG_HOME: "ayar" } });
    expect(await linuxDownloadDirs(relative)).toEqual(["/home/ayse/İndirilenler", "/home/ayse/Downloads"]);
    expect(relative.reads).toEqual(["/home/ayse/.config/user-dirs.dirs"]);
  });

  it("İndirilenler'deki (alt dizinler dahil) dosya internetten; benzer adlı kardeş dizin değil", async () => {
    const d = deps({ files: { "/home/ayse/.config/user-dirs.dirs": 'XDG_DOWNLOAD_DIR="/veri/indirilen"' } });
    expect(await internetSource("/veri/indirilen/teklif.xlsx", d)).toBe("indirilenler");
    expect(await internetSource("/veri/indirilen/eski/teklif.xlsx", d)).toBe("indirilenler");
    expect(await internetSource("/home/ayse/İndirilenler/rapor.docx", d)).toBe("indirilenler");
    expect(await internetSource("/home/ayse/Downloads/rapor.docx", d)).toBe("indirilenler");
    expect(await internetSource("/home/ayse/İndirilenler2/rapor.docx", d)).toBeNull();
    expect(await internetSource("/home/ayse/Belgeler/rapor.docx", d)).toBeNull();
    expect(await internetSource("/veri/indirilen/../belgeler/rapor.docx", d)).toBeNull();
  });

  it("dizin adındaki ayrık yazım (NFD) aynı dizin sayılır", () => {
    const nfd = "/home/ayse/İndirilenler".normalize("NFD");
    expect(nfd).not.toBe("/home/ayse/İndirilenler");
    expect(isInside(`${nfd}/a.docx`, "/home/ayse/İndirilenler")).toBe(true);
    expect(isInside("/home/ayse/İndirilenler", "/home/ayse/İndirilenler")).toBe(false);
  });

  it("Thunderbird eki (/tmp/mozilla_<kullanıcı>0) e-posta eki sayılır; diğer geçici dosyalar değil", async () => {
    const d = deps({ tmpdir: "/tmp" });
    expect(await internetSource("/tmp/mozilla_ayse0/sozlesme.docx", d)).toBe("eposta-eki");
    expect(await internetSource("/tmp/oh-deneme/sozlesme.docx", d)).toBeNull();
    expect(await internetSource("/tmp/mozilla_ayse0", d)).toBeNull();
  });
});
