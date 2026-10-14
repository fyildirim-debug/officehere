// Güncelleme kanalı (Kararlı / Beta) uygulamada: hem elle kipte (GitHub API) hem otomatik kipte
// (electron-updater'ın GitHub sağlayıcısı) yerel sahte bir GitHub sunucusuyla. Uygulama sürümü
// package.json'dakidir (0.x); sunucudaki ön sürüm 9.0.0-beta.1, kararlı sürüm 0.0.1'dir.
import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { mkdtempSync, readFileSync } from "node:fs";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");
const REPO = "fyildirim-debug/officehere";

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(env: Record<string, string>): Promise<{ app: ElectronApplication; win: Page; userData: string }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-kanal-"));
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox"],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", OH_UPDATE_DELAY_MS: "600000", ...env },
  });
  const win = await app.firstWindow();
  await win.waitForSelector(".oh-titlebar");
  return { app, win, userData };
}

/** Sürüm listesi, "latest", atom akışı ve kanal dosyaları (release.yml'nin yayımladığı biçimde). */
async function fakeGitHub() {
  const releases = [
    { tag: "v9.0.0-beta.1", prerelease: true, files: { beta: "9.0.0-beta.1" } as Record<string, string> },
    { tag: "v0.0.1", prerelease: false, files: { latest: "0.0.1", beta: "0.0.1" } as Record<string, string> },
  ];
  const json = (r: (typeof releases)[number]) => ({ tag_name: r.tag, html_url: `https://github.com/${REPO}/releases/tag/${r.tag}`, draft: false, prerelease: r.prerelease });
  const yml = (v: string) => `version: ${v}\nfiles:\n  - url: OfficeHere-${v}.AppImage\n    sha512: AA==\n    size: 1\npath: OfficeHere-${v}.AppImage\nsha512: AA==\nreleaseDate: '2026-10-14T00:00:00.000Z'\n`;
  const server = createServer((req, res) => {
    const p = new URL(req.url ?? "/", "http://yerel").pathname;
    if (p === "/api/latest" || p === `/api/v3/repos/${REPO}/releases/latest`) return res.end(JSON.stringify(json(releases.find((r) => !r.prerelease)!)));
    if (p === "/api/releases") return res.end(JSON.stringify(releases.map(json)));
    if (p === `/${REPO}/releases.atom`) {
      const entries = releases.map((r) => `<entry><id>${r.tag}</id><link rel="alternate" type="text/html" href="https://github.com/${REPO}/releases/tag/${r.tag}"/><title>${r.tag}</title></entry>`);
      return res.end(`<?xml version="1.0" encoding="UTF-8"?><feed xmlns="http://www.w3.org/2005/Atom">${entries.join("")}</feed>`);
    }
    const m = /\/releases\/download\/([^/]+)\/(latest|beta)(?:-[a-z0-9-]+)?\.yml$/.exec(p);
    const v = m ? releases.find((r) => r.tag === m[1])?.files[m[2]] : undefined;
    if (v) return res.end(yml(v));
    res.statusCode = 404;
    res.end();
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  return { env: { OH_UPDATE_FEED: `${base}/api/latest`, OH_UPDATE_RELEASES: `${base}/api/releases`, OH_UPDATE_GITHUB: base }, close: () => server.close() };
}

async function openOptions(win: Page) {
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
  return win.locator(".oh-update-section");
}

for (const mode of ["manual", "auto"] as const) {
  test(`${mode === "manual" ? "elle" : "otomatik"} kipte beta yalnız beta kanalında görünür; kararlıya dönünce unutulur`, async () => {
    const gh = await fakeGitHub();
    const { app, win, userData } = await launch({ OH_UPDATE_MODE: mode, ...gh.env });
    const section = await openOptions(win);
    const channel = section.getByRole("combobox", { name: "Güncelleme kanalı" });
    await expect(channel).toHaveValue("kararli");

    await section.getByRole("button", { name: "Güncelleştirmeleri denetle" }).click();
    await expect(section).toContainText("OfficeHere güncel.");
    await expect(win.locator(".oh-update-pill")).toHaveCount(0);

    // Beta kanalı: kanal değişince denetlenir, ön sürüm bulunur (otomatik kipte indirme başlar; testte
    // güncelleyici yapılandırması olmadığından indirme düşer ve sürüm elle indirme için bildirilir).
    await channel.selectOption("beta");
    await expect(section).toContainText("9.0.0-beta.1");
    await expect(win.locator(".oh-update-pill")).toContainText("9.0.0-beta.1");
    expect(JSON.parse(readFileSync(path.join(userData, "ayarlar.json"), "utf8")).guncellemeKanali).toBe("beta");

    await channel.selectOption("kararli");
    await expect(section).toContainText("OfficeHere güncel.");
    await expect(win.locator(".oh-update-pill")).toHaveCount(0);
    await app.close();
    gh.close();
  });
}
