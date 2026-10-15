// "Yeni Belge / Yeni Tablo / Yeni Sunu" kısayolları: Windows atlama listesi (görev çubuğu ve Başlat
// menüsünde sağ tık) ve macOS Dock menüsü. Atlama listesi görevi uygulamayı "--yeni=<uygulama>" ile
// başlatır; uygulama zaten açıksa tek örnek kilidi bu isteği çalışan örneğe (second-instance) iletir.
// Linux'ta aynı bayrağı .deb masaüstü girdisinin eylemleri kullanır (electron-builder.yml).
import { app, Menu, type MenuItemConstructorOptions, type Task } from "electron";
import { existsSync } from "node:fs";
import path from "node:path";
import type { AppKind, WindowKind } from "./apps";

export const NEW_FLAG = "--yeni";

export const NEW_ITEMS: ReadonlyArray<{ kind: AppKind; title: string; description: string }> = [
  { kind: "belge", title: "Yeni Belge", description: "Boş bir belge açar" },
  { kind: "tablo", title: "Yeni Tablo", description: "Boş bir çalışma kitabı açar" },
  { kind: "sunum", title: "Yeni Sunu", description: "Boş bir sunu açar" },
];

/** Komut satırındaki "--yeni=belge" isteği; yoksa ya da tanınmıyorsa null. */
export function newKindFromArgv(argv: readonly string[]): AppKind | null {
  const prefix = `${NEW_FLAG}=`;
  for (const a of argv) {
    if (!a.startsWith(prefix)) continue;
    const kind = NEW_ITEMS.find((i) => i.kind === a.slice(prefix.length));
    if (kind) return kind.kind;
  }
  return null;
}

/** Dosyasız açılışta açılacak pencere: görevden gelindiyse o uygulama, değilse başlangıç ekranı. */
export function startWindow(argv: readonly string[]): WindowKind {
  return newKindFromArgv(argv) ?? "baslangic";
}

const quote = (s: string) => (/\s/.test(s) ? `"${s}"` : s);

export interface TaskEnv {
  /** Çalışan programın yolu (process.execPath). */
  execPath: string;
  /** Paketlenmemiş çalışmada uygulama dizini (electron .); paketlenmişte boş. */
  appPath?: string;
  /** Uygulamanın simgesi; yoksa programın kendi simgesi kullanılır. */
  iconFor?: (kind: AppKind) => string | undefined;
}

/** Atlama listesinin "Görevler" bölümü (yan etkisiz). */
export function jumpListTasks(env: TaskEnv): Task[] {
  const base = env.appPath ? `${quote(env.appPath)} ` : "";
  return NEW_ITEMS.map((item) => {
    const icon = env.iconFor?.(item.kind);
    return {
      program: env.execPath,
      arguments: `${base}${NEW_FLAG}=${item.kind}`,
      title: item.title,
      description: item.description,
      iconPath: icon ?? env.execPath,
      iconIndex: 0,
    };
  });
}

/** macOS Dock menüsü (yan etkisiz). */
export function dockMenuTemplate(open: (kind: AppKind) => void): MenuItemConstructorOptions[] {
  return NEW_ITEMS.map((item) => ({ label: item.title, click: () => open(item.kind) }));
}

/** Kurulumda dosya türü simgeleri resources/filetypes altındadır (electron-builder.yml, win). */
function fileTypeIcon(kind: AppKind): string | undefined {
  if (!app.isPackaged) return undefined;
  const p = path.join(process.resourcesPath, "filetypes", `${kind}.ico`);
  return existsSync(p) ? p : undefined;
}

/** Atlama listesi görevlerini (Windows) ve Dock menüsünü (macOS) kurar; app hazır olduktan sonra çağrılır. */
export function initTasks(open: (kind: AppKind) => void): void {
  if (process.platform === "win32") {
    // Taşınabilir .exe her açılışta geçici bir dizine açılır: görev ve simge özgün .exe'yi göstermeli.
    const portable = process.env.PORTABLE_EXECUTABLE_FILE;
    const env: TaskEnv = portable
      ? { execPath: portable }
      : { execPath: process.execPath, appPath: app.isPackaged ? undefined : app.getAppPath(), iconFor: fileTypeIcon };
    // Kabuk görevleri reddederse (false; ör. kısıtlı oturum) uygulama yine de açılır.
    app.setUserTasks(jumpListTasks(env));
  } else if (process.platform === "darwin") {
    app.dock?.setMenu(Menu.buildFromTemplate(dockMenuTemplate(open)));
  }
}
