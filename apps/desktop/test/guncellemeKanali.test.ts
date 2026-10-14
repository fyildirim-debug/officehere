// Gerçek electron-updater (6.8, GitHub sağlayıcısı) yerel bir sahte GitHub sunucusuna karşı çalışır.
// Sunucu sürüm akışını (releases.atom), "latest" sürümü ve kanal dosyalarını release.yml'nin
// yüklediği biçimde verir: ön sürümde yalnız beta*.yml, kararlıda latest*.yml + beta*.yml.
// Kanal kuralı updates.ts ile aynı işlevle (bindChannel) bağlanır.
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createServer, request, type IncomingMessage, type RequestOptions } from "node:http";
import type { AddressInfo } from "node:net";
import { createRequire } from "node:module";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { AppUpdater } from "electron-updater";
import { bindChannel, type UpdateChannel } from "../src/main/updateCore";

// electron-updater'ın kendi HTTP katmanı (Electron "net" yerine Node http ile).
const fromUpdater = createRequire(createRequire(import.meta.url).resolve("electron-updater"));
const { HttpExecutor } = fromUpdater("builder-util-runtime") as { HttpExecutor: new () => object };

class NodeExecutor extends HttpExecutor {
  createRequest(options: RequestOptions, callback: (res: IncomingMessage) => void) {
    // Her istek kendi bağlantısıyla: test sunucusu kısa ömürlü, bekleyen bağlantı kalmasın.
    return request({ ...options, agent: false }, callback);
  }
}

class Denetleyici extends AppUpdater {
  constructor(version: string, dir: string) {
    super(null, {
      version,
      name: "OfficeHere",
      isPackaged: true,
      appUpdateConfigPath: path.join(dir, "app-update.yml"),
      userDataPath: dir,
      baseCachePath: dir,
      whenReady: () => Promise.resolve(),
      relaunch() {},
      quit() {},
      onQuit() {},
    });
  }
  protected doDownloadUpdate(): never {
    throw new Error("testte indirme yok");
  }
  quitAndInstall() {}
}

interface Release {
  tag: string;
  prerelease: boolean;
  /** Sürümdeki kanal dosyaları: kanal adı → dosyadaki sürüm. */
  files: Record<string, string>;
}

// release.yml'nin yayımladığı biçim.
const stable = (v: string): Release => ({ tag: `v${v}`, prerelease: false, files: { latest: v, beta: v } });
const beta = (v: string): Release => ({ tag: `v${v}`, prerelease: true, files: { beta: v } });
// 0.1.6'daki release.yml: ön sürüm işareti yok, electron-builder beta sürümde de latest*.yml üretir.
const oldBeta = (v: string): Release => ({ tag: `v${v}`, prerelease: false, files: { latest: v } });

let releases: Release[] = [];
let requested: string[] = [];
let host = "";
const dir = mkdtempSync(path.join(tmpdir(), "oh-guncelleme-"));

const yml = (v: string) => `version: ${v}\nfiles:\n  - url: OfficeHere-${v}.exe\n    sha512: AA==\n    size: 1\npath: OfficeHere-${v}.exe\nsha512: AA==\nreleaseDate: '2026-10-14T00:00:00.000Z'\n`;

const atom = () =>
  `<?xml version="1.0" encoding="UTF-8"?>\n<feed xmlns="http://www.w3.org/2005/Atom">` +
  releases
    .map((r) => `<entry><id>${r.tag}</id><link rel="alternate" type="text/html" href="https://github.com/fy/oh/releases/tag/${r.tag}"/><title>${r.tag}</title><content type="html">No content.</content></entry>`)
    .join("") +
  `</feed>`;

const server = createServer((req, res) => {
  const p = new URL(req.url ?? "/", "http://yerel").pathname;
  requested.push(p);
  if (p === "/fy/oh/releases.atom") {
    res.setHeader("Content-Type", "application/atom+xml");
    res.end(atom());
    return;
  }
  if (p === "/api/v3/repos/fy/oh/releases/latest") {
    // GitHub: "latest" en yeni ön sürüm olmayan sürümdür.
    const r = releases.find((x) => !x.prerelease);
    res.statusCode = r ? 200 : 404;
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify(r ? { tag_name: r.tag } : { message: "Not Found" }));
    return;
  }
  const m = /^\/fy\/oh\/releases\/download\/([^/]+)\/(latest|beta|alpha|rc)(?:-[a-z0-9-]+)?\.yml$/.exec(p);
  const version = m ? releases.find((r) => r.tag === m[1])?.files[m[2]] : undefined;
  if (version) {
    res.end(yml(version));
    return;
  }
  res.statusCode = 404;
  res.end("Not Found");
});

beforeAll(async () => {
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  host = `127.0.0.1:${(server.address() as AddressInfo).port}`;
});

afterAll(() => {
  server.close();
});

function updater(current: string) {
  const u = new Denetleyici(current, dir);
  u.logger = null;
  u.autoDownload = false;
  (u as unknown as { httpExecutor: object }).httpExecutor = new NodeExecutor();
  u.setFeedURL({ provider: "github", owner: "fy", repo: "oh", host, protocol: "http" });
  return u;
}

/** Kullanıcıya sunulan sürüm (yoksa null) ve bu denetimde okunan dosyalar. */
async function offered(u: AppUpdater, list: Release[]): Promise<string | null> {
  releases = list;
  requested = [];
  const r = await u.checkForUpdates();
  return r?.isUpdateAvailable ? r.updateInfo.version : null;
}

async function withChannel(current: string, channel: UpdateChannel, list: Release[]) {
  const u = updater(current);
  bindChannel(u, current, () => channel);
  return offered(u, list);
}

describe("güncelleme kanalı (gerçek electron-updater, sahte GitHub)", () => {
  it("kontrol: 0.1.6'nın yayın biçiminde beta etiketi kararlı kullanıcıya otomatik kurulurdu", async () => {
    // Kanal ayarı olmadan (0.1.6 davranışı) electron-updater betayı "latest" sürüm olarak alır.
    expect(await offered(updater("0.1.6"), [oldBeta("0.2.0-beta.1"), stable("0.1.6")])).toBe("0.2.0-beta.1");
  });

  it("kararlı kullanıcı, işareti unutulmuş beta sürümünü bile almaz", async () => {
    expect(await withChannel("0.1.6", "kararli", [oldBeta("0.2.0-beta.1"), stable("0.1.6")])).toBeNull();
  });

  it("kararlı kullanıcı ön sürüm olarak yayımlanan betaya hiç dokunmaz", async () => {
    expect(await withChannel("0.1.6", "kararli", [beta("0.2.0-beta.1"), stable("0.1.6")])).toBeNull();
    expect(requested.some((p) => p.includes("/v0.2.0-beta.1/"))).toBe(false);
    expect(requested.some((p) => /\/v0\.1\.6\/latest[^/]*\.yml$/.test(p))).toBe(true);
  });

  it("beta kullanıcısı ön sürümü görür", async () => {
    expect(await withChannel("0.1.6", "beta", [beta("0.2.0-beta.1"), stable("0.1.6")])).toBe("0.2.0-beta.1");
    expect(requested.some((p) => /\/v0\.2\.0-beta\.1\/beta[^/]*\.yml$/.test(p))).toBe(true);
  });

  it("beta kullanıcısı sonraki kararlı sürümü de görür", async () => {
    const list = [stable("0.2.0"), beta("0.2.0-beta.1"), stable("0.1.6")];
    expect(await withChannel("0.2.0-beta.1", "beta", list)).toBe("0.2.0");
    expect(await withChannel("0.2.0-beta.1", "kararli", list)).toBe("0.2.0");
  });

  it("betadan kararlıya dönüşte geri sürüme inilmez", async () => {
    const list = [beta("0.2.0-beta.1"), stable("0.1.6")];
    expect(await withChannel("0.2.0-beta.1", "kararli", list)).toBeNull();
    // Doğrulama: electron-updater 6.8'de channel ataması allowDowngrade'i açar; açıkça kapatılmasa
    // kararlıya dönen kullanıcıya 0.1.6 sunulurdu.
    const u = updater("0.2.0-beta.1");
    u.allowPrerelease = false;
    u.channel = "latest";
    expect(u.allowDowngrade).toBe(true);
    expect(await offered(u, list)).toBe("0.1.6");
  });

  it("betadaki kullanıcıya daha sonra çıkan eski numaralı kararlı düzeltme sunulmaz", async () => {
    expect(await withChannel("0.2.0-beta.1", "beta", [stable("0.1.7"), beta("0.2.0-beta.1"), stable("0.1.6")])).toBeNull();
    expect(await withChannel("0.1.6", "kararli", [stable("0.1.7"), beta("0.2.0-beta.1"), stable("0.1.6")])).toBe("0.1.7");
  });

  it("rc ön sürümü hiçbir kanala otomatik gelmez", async () => {
    const rc: Release = { tag: "v0.2.0-rc.1", prerelease: true, files: { beta: "0.2.0-rc.1" } };
    expect(await withChannel("0.2.0-beta.2", "beta", [rc, beta("0.2.0-beta.2"), stable("0.1.6")])).toBeNull();
    expect(await withChannel("0.1.6", "kararli", [rc, beta("0.2.0-beta.2"), stable("0.1.6")])).toBeNull();
  });
});
