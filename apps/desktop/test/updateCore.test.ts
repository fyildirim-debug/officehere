import { describe, expect, it } from "vitest";
import { compareVersions, releaseFromGitHub, updateMode } from "../src/main/updateCore";

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

describe("releaseFromGitHub", () => {
  it("son sürümün etiketini ve sayfasını okur", () => {
    expect(releaseFromGitHub({ tag_name: "v0.1.2", html_url: "https://github.com/fyildirim-debug/officehere/releases/tag/v0.1.2" })).toEqual({
      version: "0.1.2",
      url: "https://github.com/fyildirim-debug/officehere/releases/tag/v0.1.2",
    });
  });

  it("taslak, ön sürüm ve bozuk yanıtları yok sayar", () => {
    expect(releaseFromGitHub({ tag_name: "v0.2.0", draft: true })).toBeNull();
    expect(releaseFromGitHub({ tag_name: "v0.2.0", prerelease: true })).toBeNull();
    expect(releaseFromGitHub({ tag_name: "son" })).toBeNull();
    expect(releaseFromGitHub(null)).toBeNull();
    expect(releaseFromGitHub({ message: "API rate limit exceeded" })).toBeNull();
  });

  it("https olmayan bağlantıyı kullanmaz", () => {
    expect(releaseFromGitHub({ tag_name: "v0.1.2", html_url: "javascript:alert(1)" })?.url).toBe("");
  });
});
