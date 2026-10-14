import { describe, expect, it } from "vitest";
import { acceptsVersion, bindChannel, channelFrom, compareVersions, pickRelease, updateMode, type ChannelUpdater, type UpdateChannel } from "../src/main/updateCore";

describe("compareVersions", () => {
  it("sürüm numaralarını sayısal karşılaştırır", () => {
    expect(compareVersions("0.1.2", "0.1.1")).toBeGreaterThan(0);
    expect(compareVersions("0.1.10", "0.1.9")).toBeGreaterThan(0);
    expect(compareVersions("1.0.0", "0.99.99")).toBeGreaterThan(0);
    expect(compareVersions("v0.1.1", "0.1.1")).toBe(0);
    expect(compareVersions("0.1.1", "0.2.0")).toBeLessThan(0);
  });

  it("ön sürüm aynı numaralı kararlı sürümden eskidir", () => {
    expect(compareVersions("1.0.0-beta.1", "1.0.0")).toBeLessThan(0);
    expect(compareVersions("1.0.0", "1.0.0-rc.1")).toBeGreaterThan(0);
    expect(compareVersions("1.0.0-beta.10", "1.0.0-beta.2")).toBeGreaterThan(0);
  });
});

describe("updateMode", () => {
  const m = (platform: string, env: Record<string, string> = {}, packaged = true) => updateMode({ platform, packaged, env });

  it("Windows kurulumu ve AppImage kendini günceller", () => {
    expect(m("win32")).toBe("auto");
    expect(m("linux", { APPIMAGE: "/opt/OfficeHere.AppImage" })).toBe("auto");
  });

  it("taşınabilir exe, deb ve macOS yalnız bildirir", () => {
    expect(m("win32", { PORTABLE_EXECUTABLE_DIR: "C:\\Araclar" })).toBe("manual");
    expect(m("linux")).toBe("manual");
    expect(m("darwin")).toBe("manual");
  });

  it("geliştirme ortamında kapalıdır; ortam değişkeniyle zorlanabilir", () => {
    expect(m("win32", {}, false)).toBe("off");
    expect(m("linux", { OH_UPDATE_MODE: "manual" }, false)).toBe("manual");
  });
});

const rel = (tag: string, extra: Record<string, unknown> = {}) => ({
  tag_name: tag,
  html_url: `https://github.com/fyildirim-debug/officehere/releases/tag/${tag}`,
  draft: false,
  prerelease: tag.includes("-"),
  ...extra,
});

describe("pickRelease", () => {
  it("son sürümün etiketini ve sayfasını okur", () => {
    expect(pickRelease(rel("v0.1.2"), "kararli")).toEqual({ version: "0.1.2", url: "https://github.com/fyildirim-debug/officehere/releases/tag/v0.1.2" });
  });

  it("taslak, ön sürüm ve bozuk yanıtları kararlı kanalda yok sayar", () => {
    expect(pickRelease({ tag_name: "v0.2.0", draft: true }, "kararli")).toBeNull();
    expect(pickRelease({ tag_name: "v0.2.0", prerelease: true }, "kararli")).toBeNull();
    expect(pickRelease({ tag_name: "son" }, "kararli")).toBeNull();
    expect(pickRelease(null, "kararli")).toBeNull();
    expect(pickRelease({ message: "API rate limit exceeded" }, "kararli")).toBeNull();
  });

  it("https olmayan bağlantıyı kullanmaz", () => {
    expect(pickRelease({ tag_name: "v0.1.2", html_url: "javascript:alert(1)" }, "kararli")?.url).toBe("");
  });

  it("kararlı kanal, ön sürüm işareti unutulmuş beta etiketini de almaz", () => {
    expect(pickRelease({ tag_name: "v0.2.0-beta.1", prerelease: false }, "kararli")).toBeNull();
    expect(pickRelease([rel("v0.2.0-beta.1", { prerelease: false }), rel("v0.1.6")], "kararli")?.version).toBe("0.1.6");
  });

  it("beta kanalı listedeki en yeni kararlı ya da beta sürümünü seçer", () => {
    const list = [rel("v0.1.7"), rel("v0.2.0-beta.2"), rel("v0.2.0-beta.10"), rel("v0.1.6"), rel("v0.3.0", { draft: true })];
    expect(pickRelease(list, "beta")?.version).toBe("0.2.0-beta.10");
    expect(pickRelease([...list, rel("v0.2.0")], "beta")?.version).toBe("0.2.0");
    expect(pickRelease(list, "kararli")?.version).toBe("0.1.7");
  });

  it("rc ve alpha ön sürümleri hiçbir kanala bildirilmez", () => {
    const list = [rel("v0.2.0-rc.1"), rel("v0.2.0-alpha.3"), rel("v0.2.0-beta.1")];
    expect(pickRelease(list, "beta")?.version).toBe("0.2.0-beta.1");
    expect(pickRelease(list, "kararli")).toBeNull();
  });
});

describe("acceptsVersion", () => {
  it("kararlı kanal yalnız daha yeni kararlı sürümü alır", () => {
    expect(acceptsVersion("0.1.7", "0.1.6", "kararli")).toBe(true);
    expect(acceptsVersion("0.2.0-beta.1", "0.1.6", "kararli")).toBe(false);
    expect(acceptsVersion("v0.2.0-rc.1", "0.1.6", "kararli")).toBe(false);
    expect(acceptsVersion("0.1.6", "0.1.6", "kararli")).toBe(false);
  });

  it("beta kanalı beta ve sonraki kararlıyı alır, rc/alpha almaz", () => {
    expect(acceptsVersion("0.2.0-beta.1", "0.1.6", "beta")).toBe(true);
    expect(acceptsVersion("0.2.0", "0.2.0-beta.3", "beta")).toBe(true);
    expect(acceptsVersion("0.2.0-rc.1", "0.1.6", "beta")).toBe(false);
    expect(acceptsVersion("0.2.0-alpha.1", "0.1.6", "beta")).toBe(false);
  });

  it("kanal dönüşünde geri sürüme inilmez", () => {
    // Beta kurulu kullanıcı kararlıya döndü: eski kararlı sürüm sunulmaz.
    expect(acceptsVersion("0.1.6", "0.2.0-beta.1", "kararli")).toBe(false);
    expect(acceptsVersion("0.1.9", "0.2.0-beta.1", "beta")).toBe(false);
    // Sonraki kararlı geldiğinde alınır.
    expect(acceptsVersion("0.2.0", "0.2.0-beta.1", "kararli")).toBe(true);
  });

  it("bozuk sürüm adını reddeder", () => {
    expect(acceptsVersion("son", "0.1.6", "beta")).toBe(false);
    expect(acceptsVersion("", "0.1.6", "kararli")).toBe(false);
  });
});

describe("channelFrom", () => {
  it("tanınmayan değer kararlıdır", () => {
    expect(channelFrom("beta")).toBe("beta");
    expect(channelFrom("kararli")).toBe("kararli");
    expect(channelFrom(undefined)).toBe("kararli");
    expect(channelFrom("BETA")).toBe("kararli");
  });
});

describe("bindChannel", () => {
  /** electron-updater 6.8 AppUpdater'ının kanal alanları: channel ataması allowDowngrade'i açar. */
  function fakeUpdater(current: string) {
    const u = {
      // AppUpdater: çalışan sürüm ön sürümse varsayılan true.
      allowPrerelease: current.includes("-"),
      allowDowngrade: false,
      _channel: null as string | null,
      get channel() {
        return this._channel;
      },
      set channel(v: string | null) {
        this._channel = v;
        this.allowDowngrade = true;
      },
      isUpdateSupported: (_info: { version: string }): boolean | Promise<boolean> => true,
    };
    return u satisfies ChannelUpdater;
  }

  it("kanala göre ön sürüm iznini ve kanal adını açıkça atar, geri sürüme inmeyi kapatır", () => {
    let channel: UpdateChannel = "kararli";
    const u = fakeUpdater("0.2.0-beta.1");
    const apply = bindChannel(u, "0.2.0-beta.1", () => channel);
    expect(u).toMatchObject({ allowPrerelease: false, channel: "latest", allowDowngrade: false });
    channel = "beta";
    apply();
    expect(u).toMatchObject({ allowPrerelease: true, channel: "beta", allowDowngrade: false });
  });

  it("sağlayıcının bulduğu sürümü kanal kuralından geçirir, özgün denetimi korur", async () => {
    let channel: UpdateChannel = "kararli";
    const u = fakeUpdater("0.1.6");
    let baseCalls = 0;
    u.isUpdateSupported = () => {
      baseCalls++;
      return true;
    };
    bindChannel(u, "0.1.6", () => channel);
    expect(await u.isUpdateSupported({ version: "0.2.0-beta.1" })).toBe(false);
    expect(await u.isUpdateSupported({ version: "0.1.7" })).toBe(true);
    channel = "beta";
    expect(await u.isUpdateSupported({ version: "0.2.0-beta.1" })).toBe(true);
    // Özgün denetim (en düşük işletim sistemi sürümü) yalnız kanalın kabul ettiği sürümde çağrılır.
    expect(baseCalls).toBe(2);
  });
});
