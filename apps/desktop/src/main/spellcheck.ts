// Yazım denetimi sözlükleri: Windows ve Linux'ta Chromium Türkçe ve İngilizce sözlükleri Google
// sunucusundan indirir. Kullanıcı onay vermeden bu istek gitmez (kurallar spellcheckCore.ts'te):
// onay yokken dil listesi boştur ve indirme adresi ağa çıkmayan bir adrestir. macOS sistemin
// denetleyicisini kullanır, hiçbir şey indirilmez.
import { app, BrowserWindow, ipcMain, session, type Session } from "electron";
import fs from "node:fs";
import path from "node:path";
import { readSettings, writeSettings } from "./settings";
import {
  consentFrom,
  DICTIONARY_MB,
  GOOGLE_DICTIONARY_URL,
  NO_DOWNLOAD_URL,
  removableDictionaries,
  SPELL_LANGUAGES,
  spellPlan,
  spellStatus,
  type SpellPlan,
  type SpellStatus,
} from "./spellcheckCore";

// Testler için: onaydan sonraki indirme Google yerine bu adrese (yerel sunucu) gider.
const DOWNLOAD_URL = process.env.OH_SPELLCHECK_DICT_URL ?? GOOGLE_DICTIONARY_URL;

interface SpellcheckState {
  status: SpellStatus;
  /** İlk açılış sorusu gösterilsin mi. */
  ask: boolean;
  /** İndirilecek yaklaşık boyut (MB). */
  sizeMb: number;
}

let enabled = false;
let ask = false;
let failed = false;
const downloading = new Set<string>();

function dictionaryDir() {
  return path.join(app.getPath("userData"), "Dictionaries");
}

function localFiles(): string[] {
  try {
    return fs.readdirSync(dictionaryDir());
  } catch {
    return [];
  }
}

function currentState(): SpellcheckState {
  return { status: spellStatus({ platform: process.platform, enabled, files: localFiles(), downloading: downloading.size > 0, failed }), ask, sizeMb: DICTIONARY_MB };
}

function broadcast(): SpellcheckState {
  const s = currentState();
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send("oh:spellcheck-state", s);
  return s;
}

/** Bir adım hata verse de sonrakiler uygulanır (açılış durmaz; dil listesi her durumda boşaltılır). */
function attempt(step: () => void) {
  try {
    step();
  } catch {
    // Sonraki adım (boş dil listesi) indirmeyi tek başına da engeller.
  }
}

function apply(ses: Session, plan: SpellPlan) {
  enabled = plan.enable;
  ask = plan.ask;
  if (process.platform === "darwin") return;
  attempt(() => ses.setSpellCheckerDictionaryDownloadURL(plan.download ? DOWNLOAD_URL : NO_DOWNLOAD_URL));
  // Boş liste bekleyen sözlük yüklemesini iptal eder; doluysa önce boşaltmak, aynı listeyle
  // yeniden denemeyi (ör. indirme hatasından sonra) de tetikler.
  attempt(() => ses.setSpellCheckerLanguages([]));
  if (plan.enable) {
    attempt(() => ses.setSpellCheckerLanguages(SPELL_LANGUAGES));
    attempt(() => ses.setSpellCheckerEnabled(true));
  } else {
    attempt(() => ses.setSpellCheckerEnabled(false));
  }
}

/** Bu iki dilin sözlük dosyalarını siler. Windows'ta kullanımdaysa silinemez; bir sonraki açılışta silinir. */
function removeDictionaries() {
  for (const f of removableDictionaries(localFiles())) {
    try {
      fs.rmSync(path.join(dictionaryDir(), f), { force: true });
    } catch {
      // Açılışta yeniden denenir (onay "reddedildi" iken).
    }
  }
}

/** Kullanıcının yanıtı: İndir (true) ya da Şimdi değil / Kaldır (false). */
function setDictionaries(on: boolean): SpellcheckState {
  writeSettings({ yazimSozlugu: on ? "izinli" : "reddedildi" });
  failed = false;
  downloading.clear();
  apply(session.defaultSession, spellPlan({ platform: process.platform, consent: on ? "izinli" : "reddedildi", files: localFiles() }));
  if (!on) removeDictionaries();
  return broadcast();
}

/**
 * Uygulama hazır olunca, herhangi bir pencere açılmadan önce çağrılır. Oturuma ilk erişimde Chromium
 * kayıtlı (ya da sistem dilinin) sözlüğünü yüklemeye başlar; dil listesi aynı adımda ayarlandığı için
 * onay yoksa bu yükleme iptal olur ve indirme isteği hiç çıkmaz.
 */
export function initSpellcheck() {
  // Karar oturuma dokunmadan önceki dosyalara göre verilir: Chromium oturum açılırken bozuk bulduğu
  // sözlüğü arka planda silebilir.
  const consent = consentFrom(readSettings().yazimSozlugu);
  const plan = spellPlan({ platform: process.platform, consent, files: localFiles() });
  const ses = session.defaultSession;
  apply(ses, plan);
  if (consent === "reddedildi") removeDictionaries();

  ses.on("spellcheck-dictionary-download-begin", (_e, lang) => {
    downloading.add(lang);
    failed = false;
    broadcast();
  });
  ses.on("spellcheck-dictionary-download-success", (_e, lang) => {
    downloading.delete(lang);
    broadcast();
  });
  ses.on("spellcheck-dictionary-download-failure", (_e, lang) => {
    downloading.delete(lang);
    failed = true;
    broadcast();
  });
  ses.on("spellcheck-dictionary-initialized", () => broadcast());

  ipcMain.handle("oh:spellcheck-get", () => currentState());
  ipcMain.handle("oh:spellcheck-set", (_e, on: unknown) => setDictionaries(on === true));
}
