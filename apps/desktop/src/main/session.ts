// Oturum: uygulama kapanırken açık belgelerin yolları userData/oturum.json'a yazılır;
// "Açılışta son oturumu geri yükle" açıksa sonraki dosyasız açılışta bu belgeler yeniden açılır.
import { app } from "electron";
import fs from "node:fs";
import path from "node:path";
import { appForPath, type AppKind } from "./apps";

export interface SessionDoc {
  path: string;
  app: AppKind;
}

function sessionPath() {
  return path.join(app.getPath("userData"), "oturum.json");
}

export function writeSession(docs: SessionDoc[]) {
  try {
    fs.mkdirSync(path.dirname(sessionPath()), { recursive: true });
    fs.writeFileSync(sessionPath(), JSON.stringify({ kaydedildi: Date.now(), belgeler: docs }, null, 2));
  } catch {
    // Oturum yazılamazsa geri yükleme başlangıç ekranına düşer.
  }
}

export function readSession(): SessionDoc[] {
  try {
    const data = JSON.parse(fs.readFileSync(sessionPath(), "utf8")) as { belgeler?: unknown };
    if (!Array.isArray(data?.belgeler)) return [];
    return (data.belgeler as Partial<SessionDoc>[]).filter(
      (d): d is SessionDoc => !!d && typeof d.path === "string" && (d.app === "belge" || d.app === "tablo" || d.app === "sunum"),
    );
  } catch {
    return [];
  }
}

/** Son oturumdaki, hâlâ diskte duran belgeler (yinelenenler atılır). */
export function restorableSession(): SessionDoc[] {
  const seen = new Set<string>();
  const out: SessionDoc[] = [];
  for (const d of readSession()) {
    if (seen.has(d.path) || !path.isAbsolute(d.path)) continue;
    seen.add(d.path);
    try {
      if (!fs.statSync(d.path).isFile()) continue;
    } catch {
      continue;
    }
    out.push({ path: d.path, app: appForPath(d.path) ?? d.app });
  }
  return out;
}
