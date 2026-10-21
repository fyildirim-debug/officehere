// Kullanıcı ayarları: userData/ayarlar.json. Dosyayı güncelleştirme denetimi (otomatikGuncellemeDenetimi,
// guncellemeKanali: "kararli" | "beta") ve yazım sözlüğü onayı (yazimSozlugu: "izinli" | "reddedildi";
// yoksa hiç sorulmamıştır) da kullanır; yazarken mevcut anahtarlar korunur.
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

/** Sayfaların köprü üzerinden okuyup yazabildiği ayarlar (güncelleştirme ve yazım sözlüğünün kendi IPC'si var). */
export interface DesktopSettings {
  oturumuGeriYukle: boolean;
  /**
   * Görüntüleyici kipi: İndirilenler ve e-posta eklerinden gelen dosyalar salt okunur açılır. Varsayılan
   * kapalı: OfficeHere makro çalıştırmaz, otomatik kaydetme yedekler (.bak, önceki sürümler) ve doğrulamayla korunur; Linux'ta
   * tespit yalnız dizine dayandığından varsayılan açık olsa İndirilenler'deki her dosya kilitli açılırdı.
   */
  indirilenleriSaltOkunurAc: boolean;
}

const KEYS = ["oturumuGeriYukle", "indirilenleriSaltOkunurAc"] as const;

export function desktopSettings(): DesktopSettings {
  const s = readSettings();
  return { oturumuGeriYukle: s.oturumuGeriYukle === true, indirilenleriSaltOkunurAc: s.indirilenleriSaltOkunurAc === true };
}

export function setDesktopSetting(key: unknown, value: unknown): DesktopSettings {
  if (KEYS.includes(key as (typeof KEYS)[number]) && typeof value === "boolean") writeSettings({ [key as string]: value });
  return desktopSettings();
}
