// Kullanıcı ayarları: userData/ayarlar.json. Dosyayı güncelleştirme denetimi (otomatikGuncellemeDenetimi,
// guncellemeKanali: "kararli" | "beta") da kullanır; yazarken mevcut anahtarlar korunur.
import { app } from "electron";
import fs from "node:fs";
import path from "node:path";

function settingsPath() {
  return path.join(app.getPath("userData"), "ayarlar.json");
}

export function readSettings(): Record<string, unknown> {
  try {
    const data = JSON.parse(fs.readFileSync(settingsPath(), "utf8")) as unknown;
    return data && typeof data === "object" && !Array.isArray(data) ? (data as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

/** Verilen anahtarları mevcut ayarlarla birleştirerek yazar. */
export function writeSettings(patch: Record<string, unknown>) {
  try {
    fs.mkdirSync(path.dirname(settingsPath()), { recursive: true });
    fs.writeFileSync(settingsPath(), JSON.stringify({ ...readSettings(), ...patch }, null, 2));
  } catch {
    // Ayar yazılamazsa yalnız bu oturum için geçerli olur.
  }
}

/** Sayfaların köprü üzerinden okuyup yazabildiği ayarlar (güncelleştirme ayarının kendi IPC'si var). */
export interface DesktopSettings {
  oturumuGeriYukle: boolean;
}

export function desktopSettings(): DesktopSettings {
  return { oturumuGeriYukle: readSettings().oturumuGeriYukle === true };
}

export function setDesktopSetting(key: unknown, value: unknown): DesktopSettings {
  if (key === "oturumuGeriYukle" && typeof value === "boolean") writeSettings({ oturumuGeriYukle: value });
  return desktopSettings();
}
