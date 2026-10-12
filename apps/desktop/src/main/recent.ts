import { app } from "electron";
import fs from "node:fs";
import path from "node:path";
import { addRecentEntry, normalizeRecent, pinRecentEntry, removeRecentEntry, type RecentEntry } from "./recentList";

export type RecentFile = RecentEntry;

const MAX = 40;

function storePath() {
  return path.join(app.getPath("userData"), "son-kullanilanlar.json");
}

/** Taşınan/silinen dosyalar da listede kalır; arayüz onları recentStatus ile "Bulunamadı" gösterir. */
export function readRecent(): RecentFile[] {
  try {
    return normalizeRecent(JSON.parse(fs.readFileSync(storePath(), "utf8")));
  } catch {
    return [];
  }
}

function writeRecent(list: RecentFile[]) {
  try {
    fs.mkdirSync(path.dirname(storePath()), { recursive: true });
    fs.writeFileSync(storePath(), JSON.stringify(list, null, 2));
  } catch {
    // Kullanıcı verisi dizini yazılamıyorsa son kullanılanlar tutulmaz.
  }
}

export function addRecent(file: RecentFile) {
  if (!file || typeof file.path !== "string" || !path.isAbsolute(file.path)) return;
  const [entry] = normalizeRecent([file]);
  if (!entry) return;
  writeRecent(addRecentEntry(readRecent(), entry, MAX));
  // Windows atlama listesi / macOS Son Kullanılan Öğeler.
  app.addRecentDocument(file.path);
}

export function pinRecent(filePath: string, pinned: boolean) {
  if (typeof filePath !== "string") return;
  writeRecent(pinRecentEntry(readRecent(), filePath, pinned === true));
}

export function removeRecent(filePath: string) {
  if (typeof filePath !== "string") return;
  writeRecent(removeRecentEntry(readRecent(), filePath));
}

/** Yollardaki dosyalar hâlâ var mı. */
export async function recentStatus(paths: unknown): Promise<Record<string, boolean>> {
  if (!Array.isArray(paths)) return {};
  const list = paths.filter((p): p is string => typeof p === "string").slice(0, 500);
  const result: Record<string, boolean> = {};
  await Promise.all(
    list.map(async (p) => {
      try {
        result[p] = path.isAbsolute(p) && (await fs.promises.stat(p)).isFile();
      } catch {
        result[p] = false;
      }
    }),
  );
  return result;
}
