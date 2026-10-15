import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Ana süreç modülü Electron'suz sınanır: atlama listesi ve Dock çağrıları kaydedilir.
const electron = vi.hoisted(() => ({
  app: {
    isPackaged: true,
    getAppPath: vi.fn(() => "/depo/apps/desktop"),
    setUserTasks: vi.fn(() => true),
    dock: { setMenu: vi.fn() },
  },
  Menu: { buildFromTemplate: vi.fn((t: unknown) => ({ template: t })) },
}));
vi.mock("electron", () => electron);

import { dockMenuTemplate, initTasks, jumpListTasks, newKindFromArgv, NEW_FLAG, startWindow } from "../src/main/tasks";

describe("newKindFromArgv / startWindow", () => {
  it("görevin bayrağını argv'nin herhangi bir yerinde tanır", () => {
    expect(newKindFromArgv(["C:\\OfficeHere\\OfficeHere.exe", "--yeni=belge"])).toBe("belge");
    // Chromium ikinci örneğin argv'sine kendi anahtarlarını ekleyebilir, sıra değişebilir.
    expect(newKindFromArgv(["electron", "--allow-file-access-from-files", "--yeni=tablo", "/depo/apps/desktop"])).toBe("tablo");
    expect(newKindFromArgv(["/opt/OfficeHere/officehere", "--yeni=sunum"])).toBe("sunum");
  });

  it("bilinmeyen ya da eksik değer yok sayılır; dosya yolları bayrak sayılmaz", () => {
    expect(newKindFromArgv(["app", "--yeni=baslangic"])).toBeNull();
    expect(newKindFromArgv(["app", "--yeni="])).toBeNull();
    expect(newKindFromArgv(["app", "--yeni", "belge"])).toBeNull();
    expect(newKindFromArgv(["app", "C:\\Belgeler\\--yeni=belge.docx"])).toBeNull();
    expect(newKindFromArgv([])).toBeNull();
  });

  it("dosyasız açılışta görev yoksa başlangıç ekranı açılır", () => {
    expect(startWindow(["OfficeHere.exe"])).toBe("baslangic");
    expect(startWindow(["OfficeHere.exe", "--yeni=belge"])).toBe("belge");
  });
});

describe("jumpListTasks", () => {
  it("üç görev: Yeni Belge, Yeni Tablo, Yeni Sunu; programın kendisini bayrakla başlatır", () => {
    const tasks = jumpListTasks({ execPath: "C:\\Program Files\\OfficeHere\\OfficeHere.exe" });
    expect(tasks.map((t) => t.title)).toEqual(["Yeni Belge", "Yeni Tablo", "Yeni Sunu"]);
    expect(tasks.map((t) => t.arguments)).toEqual([`${NEW_FLAG}=belge`, `${NEW_FLAG}=tablo`, `${NEW_FLAG}=sunum`]);
    for (const t of tasks) {
      expect(t.program).toBe("C:\\Program Files\\OfficeHere\\OfficeHere.exe");
      expect(t.description.length).toBeGreaterThan(0);
      // Simge verilmezse programın kendi simgesi.
      expect(t.iconPath).toBe(t.program);
      expect(t.iconIndex).toBe(0);
      // Görevin argümanı ikinci örnekte yeniden tanınmalı.
      expect(newKindFromArgv([t.program, ...t.arguments.split(" ")])).not.toBeNull();
    }
  });

  it("dosya türü simgesi varsa o kullanılır", () => {
    const tasks = jumpListTasks({ execPath: "C:\\OH\\OfficeHere.exe", iconFor: (k) => (k === "tablo" ? undefined : `C:\\OH\\resources\\filetypes\\${k}.ico`) });
    expect(tasks.map((t) => t.iconPath)).toEqual(["C:\\OH\\resources\\filetypes\\belge.ico", "C:\\OH\\OfficeHere.exe", "C:\\OH\\resources\\filetypes\\sunum.ico"]);
  });

  it("paketlenmemiş çalışmada uygulama dizini (boşluklu yol tırnakla) önce verilir", () => {
    const [t] = jumpListTasks({ execPath: "C:\\depo\\electron.exe", appPath: "C:\\Kod Deposu\\apps\\desktop" });
    expect(t.arguments).toBe(`"C:\\Kod Deposu\\apps\\desktop" ${NEW_FLAG}=belge`);
    const [u] = jumpListTasks({ execPath: "/usr/bin/electron", appPath: "/depo/apps/desktop" });
    expect(u.arguments).toBe(`/depo/apps/desktop ${NEW_FLAG}=belge`);
  });
});

describe("dockMenuTemplate", () => {
  it("her öğe ilgili uygulamanın yeni penceresini açar", () => {
    const open = vi.fn();
    const items = dockMenuTemplate(open);
    expect(items.map((i) => i.label)).toEqual(["Yeni Belge", "Yeni Tablo", "Yeni Sunu"]);
    for (const i of items) (i.click as () => void)();
    expect(open.mock.calls).toEqual([["belge"], ["tablo"], ["sunum"]]);
  });
});

describe("initTasks", () => {
  const platform = process.platform;
  const setPlatform = (p: string) => Object.defineProperty(process, "platform", { value: p, configurable: true });
  beforeEach(() => vi.clearAllMocks());
  afterEach(() => {
    setPlatform(platform);
    delete (process as { resourcesPath?: string }).resourcesPath;
  });

  it("Windows'ta atlama listesine üç görev yazılır (kurulu dosya türü simgeleriyle), Dock'a dokunulmaz", () => {
    setPlatform("win32");
    // Kurulumdaki resources/filetypes: yalnız belge ve sunum simgesi varsa tablo programın simgesini alır.
    const resources = mkdtempSync(path.join(tmpdir(), "oh-gorev-"));
    mkdirSync(path.join(resources, "filetypes"));
    for (const k of ["belge", "sunum"]) writeFileSync(path.join(resources, "filetypes", `${k}.ico`), "");
    Object.defineProperty(process, "resourcesPath", { value: resources, configurable: true });
    initTasks(vi.fn());
    expect(electron.app.setUserTasks).toHaveBeenCalledTimes(1);
    const [tasks] = electron.app.setUserTasks.mock.calls[0] as unknown as [Array<{ title: string; program: string; arguments: string; iconPath: string }>];
    expect(tasks.map((t) => t.title)).toEqual(["Yeni Belge", "Yeni Tablo", "Yeni Sunu"]);
    expect(tasks.every((t) => t.program === process.execPath)).toBe(true);
    expect(tasks.map((t) => t.arguments)).toEqual(["--yeni=belge", "--yeni=tablo", "--yeni=sunum"]);
    expect(tasks.map((t) => t.iconPath)).toEqual([path.join(resources, "filetypes", "belge.ico"), process.execPath, path.join(resources, "filetypes", "sunum.ico")]);
    expect(electron.app.dock.setMenu).not.toHaveBeenCalled();
  });

  it("macOS'ta Dock menüsü kurulur ve tıklama pencere açar", () => {
    setPlatform("darwin");
    const open = vi.fn();
    initTasks(open);
    expect(electron.app.setUserTasks).not.toHaveBeenCalled();
    expect(electron.app.dock.setMenu).toHaveBeenCalledTimes(1);
    const [template] = electron.Menu.buildFromTemplate.mock.calls[0] as unknown as [Array<{ label: string; click: () => void }>];
    template[2].click();
    expect(open).toHaveBeenCalledWith("sunum");
  });

  it("Linux'ta (masaüstü eylemleri .deb girdisinde) hiçbir şey yapılmaz", () => {
    setPlatform("linux");
    initTasks(vi.fn());
    expect(electron.app.setUserTasks).not.toHaveBeenCalled();
    expect(electron.app.dock.setMenu).not.toHaveBeenCalled();
  });
});
