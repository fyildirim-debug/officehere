import { describe, expect, it } from "vitest";
import { consentFrom, DICTIONARY_FILES, NO_DOWNLOAD_URL, removableDictionaries, spellPlan, spellStatus } from "../src/main/spellcheckCore";

const both = [...DICTIONARY_FILES];

describe("spellPlan", () => {
  it("hiç sorulmamış temiz profilde yazım denetimi kapalı kalır ve soru sorulur", () => {
    expect(spellPlan({ platform: "linux", consent: undefined, files: [] })).toEqual({ enable: false, download: false, ask: true });
    expect(spellPlan({ platform: "win32", consent: undefined, files: ["Custom Dictionary.txt"] })).toEqual({ enable: false, download: false, ask: true });
  });

  it("sözlükler zaten yereldeyse ağa çıkmadan etkinleşir, soru sorulmaz", () => {
    expect(spellPlan({ platform: "win32", consent: undefined, files: [...both, "Custom Dictionary.txt"] })).toEqual({ enable: true, download: false, ask: false });
  });

  it("yalnız biri yereldeyse eksik olan onaysız indirilmez", () => {
    expect(spellPlan({ platform: "linux", consent: undefined, files: ["tr-TR-4-0.bdic"] })).toEqual({ enable: false, download: false, ask: true });
    // Electron yükseltmesi sözlük sürümünü değiştirdiyse eski dosya yetmez.
    expect(spellPlan({ platform: "linux", consent: undefined, files: ["tr-TR-3-0.bdic", "en-US-9-0.bdic"] }).ask).toBe(true);
  });

  it("onay verildiyse indirme serbest; reddedildiyse kapalı ve bir daha sorulmaz", () => {
    expect(spellPlan({ platform: "linux", consent: "izinli", files: [] })).toEqual({ enable: true, download: true, ask: false });
    expect(spellPlan({ platform: "linux", consent: "reddedildi", files: [] })).toEqual({ enable: false, download: false, ask: false });
    expect(spellPlan({ platform: "linux", consent: "reddedildi", files: both })).toEqual({ enable: false, download: false, ask: false });
  });

  it("macOS sistemin denetleyicisini kullanır: indirme yok, soru yok", () => {
    expect(spellPlan({ platform: "darwin", consent: undefined, files: [] })).toEqual({ enable: true, download: false, ask: false });
    expect(spellPlan({ platform: "darwin", consent: "reddedildi", files: [] }).download).toBe(false);
  });
});

describe("spellStatus", () => {
  const st = (o: Partial<Parameters<typeof spellStatus>[0]>) => spellStatus({ platform: "linux", enabled: true, files: [], downloading: false, failed: false, ...o });

  it("durumları ayırır", () => {
    expect(st({ platform: "darwin" })).toBe("sistem");
    expect(st({ enabled: false, files: both })).toBe("yok");
    expect(st({ downloading: true })).toBe("indiriliyor");
    expect(st({ failed: true })).toBe("hata");
    expect(st({ files: both })).toBe("hazir");
    // Onay verildi, indirme olayı henüz gelmedi.
    expect(st({})).toBe("indiriliyor");
  });
});

describe("consentFrom / removableDictionaries / NO_DOWNLOAD_URL", () => {
  it("tanınmayan onay değeri sorulmamış sayılır", () => {
    expect(consentFrom("izinli")).toBe("izinli");
    expect(consentFrom("reddedildi")).toBe("reddedildi");
    expect(consentFrom(true)).toBeUndefined();
    expect(consentFrom(undefined)).toBeUndefined();
  });

  it("yalnız Türkçe ve İngilizce sözlük dosyaları silinir", () => {
    expect(removableDictionaries(["tr-TR-4-0.bdic", "en-US-10-1.bdic", "tr-TR-3-0.bdic", "de-DE-3-0.bdic", "Custom Dictionary.txt"])).toEqual([
      "tr-TR-4-0.bdic",
      "en-US-10-1.bdic",
      "tr-TR-3-0.bdic",
    ]);
  });

  it("onaysız indirme adresi ağa çıkmaz", () => {
    expect(new URL(NO_DOWNLOAD_URL).protocol).toBe("file:");
  });
});
