// Yazım denetimi sözlüklerinin yan etkisiz kuralları (birim testli). Windows ve Linux'ta Chromium
// Hunspell sözlüklerini Google sunucusundan indirir; OfficeHere bunu yalnız kullanıcı onayıyla yapar.

/** Yazım denetimi dilleri. */
export const SPELL_LANGUAGES = ["tr-TR", "en-US"];

/**
 * Chromium'un bu Electron sürümünde (44) beklediği sözlük dosyaları (userData/Dictionaries).
 * Electron yükseltmesi sözlük sürümünü değiştirirse dosyalar "yerelde yok" sayılır ve onay yeniden sorulur.
 */
export const DICTIONARY_FILES = ["tr-TR-4-0.bdic", "en-US-10-1.bdic"];

/** İndirilecek toplam boyut (ölçülen: 9.644.083 + 451.968 bayt). */
export const DICTIONARY_MB = 10;

/** Chromium'un varsayılan sözlük adresi. */
export const GOOGLE_DICTIONARY_URL = "https://redirector.gvt1.com/edgedl/chrome/dict/";

/** Onay yokken indirme adresi: dosya şeması ağa çıkmaz, Chromium'un olası denemesi hemen düşer. */
export const NO_DOWNLOAD_URL = "file:///officehere-sozluk-indirilmez/";

/** ayarlar.json "yazimSozlugu": kullanıcının yanıtı; hiç sorulmadıysa undefined. */
export type DictionaryConsent = "izinli" | "reddedildi" | undefined;

export function consentFrom(value: unknown): DictionaryConsent {
  return value === "izinli" || value === "reddedildi" ? value : undefined;
}

export interface SpellPlan {
  /** Diller etkin mi; değilse dil listesi boştur ve Chromium hiçbir sözlüğü yüklemez ya da indirmez. */
  enable: boolean;
  /** Eksik sözlük indirilebilir mi; değilse indirme adresi ağa çıkmayan adrestir. */
  download: boolean;
  /** İlk açılış sorusu gösterilsin mi. */
  ask: boolean;
}

/**
 * Açılışta ve kullanıcının yanıtından sonra yazım denetiminin durumu.
 * - macOS sistemin denetleyicisini kullanır: sözlük indirilmez, soru sorulmaz.
 * - Onay varsa diller etkin, indirme serbest.
 * - Onay yoksa ama iki sözlük de yerelde (önceki sürümlerin kullanıcıları) ağa çıkmadan kullanılır.
 * - Hiç sorulmadıysa ve sözlük yoksa: kapalı, ilk açılış sorusu gösterilir.
 */
export function spellPlan(o: { platform: string; consent: DictionaryConsent; files: readonly string[] }): SpellPlan {
  if (o.platform === "darwin") return { enable: true, download: false, ask: false };
  if (o.consent === "izinli") return { enable: true, download: true, ask: false };
  if (o.consent === "reddedildi") return { enable: false, download: false, ask: false };
  const local = DICTIONARY_FILES.every((f) => o.files.includes(f));
  return local ? { enable: true, download: false, ask: false } : { enable: false, download: false, ask: true };
}

export type SpellStatus = "sistem" | "yok" | "indiriliyor" | "hazir" | "hata";

/** Seçenekler'de gösterilen durum. */
export function spellStatus(o: { platform: string; enabled: boolean; files: readonly string[]; downloading: boolean; failed: boolean }): SpellStatus {
  if (o.platform === "darwin") return "sistem";
  if (!o.enabled) return "yok";
  if (o.downloading) return "indiriliyor";
  if (o.failed) return "hata";
  return DICTIONARY_FILES.every((f) => o.files.includes(f)) ? "hazir" : "indiriliyor";
}

/** "Kaldır" ile silinen dosyalar: yalnız bu iki dilin sözlükleri (kullanıcının sözcük listesine dokunulmaz). */
export function removableDictionaries(files: readonly string[]): string[] {
  return files.filter((f) => /^(tr-TR|en-US)-[\d-]+\.bdic$/i.test(f));
}
