import { app } from "electron";
import fs from "node:fs";
import path from "node:path";
import type { AppKind } from "./apps";

export interface RecentFile {
  path: string;
  name: string;
  app: AppKind;
  openedAt: number;
}

const MAX = 40;

function storePath() {
  return path.join(app.getPath("userData"), "son-kullanilanlar.json");
}

export function readRecent(): RecentFile[] {
  try {
    const list = JSON.parse(fs.readFileSync(storePath(), "utf8")) as RecentFile[];
    // Silinmiş/taşınmış dosyaları listeden düşür.
    return list.filter((f) => f && typeof f.path === "string" && fs.existsSync(f.path));
  } catch {
    return [];
  }
}

export function addRecent(file: RecentFile) {
  if (!path.isAbsolute(file.path)) return;
  const list = readRecent().filter((f) => f.path !== file.path);
  list.unshift(file);
  try {
    fs.mkdirSync(path.dirname(storePath()), { recursive: true });
    fs.writeFileSync(storePath(), JSON.stringify(list.slice(0, MAX), null, 2));
  } catch {
    // Kullanıcı verisi dizini yazılamıyorsa son kullanılanlar tutulmaz.
  }
  // Windows atlama listesi / macOS Son Kullanılan Öğeler.
  app.addRecentDocument(file.path);
}
