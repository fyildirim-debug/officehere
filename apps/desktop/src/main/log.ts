// Yerel hata günlüğü: userData/gunluk/officehere.log (+ .1, .2 döner dosyalar).
// Yalnız olay türü, hata iletisi ve kod konumu yazılır; belge metni ve dosya adları
// yazılmaz (dosyalar "#özet.uzantı" biçiminde anılır, ev dizini "~" olur).
// Hiçbir şey dışarı gönderilmez; "Hata bildir" penceresi son satırları kullanıcıya
// gösterir, kullanıcı isterse kendisi gönderir. Electron'a bağlı değildir.
import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export type LogLevel = "bilgi" | "uyari" | "hata";

const LEVEL_TEXT: Record<LogLevel, string> = { bilgi: "BİLGİ", uyari: "UYARI", hata: "HATA" };

export interface LogOptions {
  dir: string;
  /** Bir günlük dosyasının en büyük boyu (bayt). */
  maxBytes?: number;
  /** Tutulan dosya sayısı (etkin dosya dahil). */
  files?: number;
  home?: string;
}

/** Dosya yolunu kişisel bilgi taşımayan kısa bir tanıtıcıya çevirir: "#3fa2c1d0.docx". */
export function describePath(p: string): string {
  const ext = path.extname(p).toLowerCase().replace(/[^.a-z0-9]/g, "").slice(0, 8); // yerel-ayar: ascii
  return `#${createHash("sha256").update(p).digest("hex").slice(0, 8)}${ext}`;
}

// Yol parçası: boşluk, tırnak ve ayırıcı içermeyen sözcükler; tek boşlukla ayrılmış sözcükler
// ("Gizli Proje") bir parça sayılır. İç içe niceleyiciler ayırıcıyla sınırlı olduğundan doğrusal çalışır.
const SEG = `[^\\s"'\u201c\u201d\u2018\u2019<>|:*?\\\\/]+`;
const PATH_RE = new RegExp(`(?:[A-Za-z]:[\\\\/]|\\\\\\\\|/)(?:${SEG}(?: ${SEG})*[\\\\/])*${SEG}(?: ${SEG})*\\.[A-Za-z0-9]{1,8}\\b`, "g");

/** Günlüğe girecek metni temizler: belge yolları özetlenir, ev dizini gizlenir, uzunluk sınırlanır. */
export function sanitize(text: string, opts: { home?: string; max?: number } = {}): string {
  // Düzenli ifadelerin maliyeti sınırlı kalsın diye önce kısaltılır.
  let t = String(text ?? "").slice(0, 4000);
  // Kod konumları (file:///…/assets/index-abc.js:12:3) yalnız dosya adına iner.
  t = t.replace(/(?:file|https?):\/\/[^\s)]*\/([^/\s)]+)/g, "$1");
  t = t.replace(PATH_RE, (m) => (/\.(?:js|mjs|cjs|ts|tsx|html|node)$/i.test(m) ? path.basename(m) : describePath(m)));
  const home = opts.home ?? os.homedir();
  if (home && home.length > 1) t = t.split(home).join("~");
  t = t.replace(/[\u0000-\u001f\u007f]+/g, " ");
  const max = opts.max ?? 600;
  return t.length > max ? `${t.slice(0, max)}…` : t;
}

/** Yığın izinin ilk satırları (yalnız kod konumları). */
export function sanitizeStack(stack: string | undefined, opts: { home?: string } = {}): string[] {
  if (!stack) return [];
  return stack
    .split("\n")
    .slice(1)
    .map((l) => l.trim())
    .filter((l) => l.startsWith("at "))
    .slice(0, 8)
    .map((l) => sanitize(l, { home: opts.home, max: 200 }));
}

export class LocalLog {
  readonly file: string;
  private readonly maxBytes: number;
  private readonly files: number;
  private readonly home?: string;
  private size = -1;

  constructor(opts: LogOptions) {
    this.file = path.join(opts.dir, "officehere.log");
    this.maxBytes = opts.maxBytes ?? 256 * 1024;
    this.files = Math.max(1, opts.files ?? 3);
    this.home = opts.home;
  }

  private rotated(i: number) {
    return i === 0 ? this.file : this.file.replace(/\.log$/, `.${i}.log`);
  }

  private rotate() {
    try {
      fs.rmSync(this.rotated(this.files - 1), { force: true });
      for (let i = this.files - 2; i >= 0; i--) {
        if (fs.existsSync(this.rotated(i))) fs.renameSync(this.rotated(i), this.rotated(i + 1));
      }
    } catch {
      // döndürülemezse dosya kesilir
      try {
        fs.writeFileSync(this.file, "");
      } catch {
        /* yok say */
      }
    }
    this.size = 0;
  }

  /** Bir satır (ve varsa yığın izi) yazar. Günlük yazılamazsa sessizce geçer. */
  write(level: LogLevel, source: string, message: string, detail?: { stack?: string; code?: string }) {
    try {
      const lines = [`${new Date().toISOString()} ${LEVEL_TEXT[level]} [${sanitize(source, { max: 40, home: this.home })}] ${sanitize(message, { home: this.home })}${detail?.code ? ` (${sanitize(detail.code, { max: 40 })})` : ""}`];
      for (const s of sanitizeStack(detail?.stack, { home: this.home })) lines.push(`    ${s}`);
      const text = `${lines.join("\n")}\n`;
      fs.mkdirSync(path.dirname(this.file), { recursive: true });
      if (this.size < 0) {
        try {
          this.size = fs.statSync(this.file).size;
        } catch {
          this.size = 0;
        }
      }
      const bytes = Buffer.byteLength(text);
      if (this.size > 0 && this.size + bytes > this.maxBytes) this.rotate();
      fs.appendFileSync(this.file, text);
      this.size += bytes;
    } catch {
      // Günlük yazılamaz (salt okunur profil): uygulama etkilenmez.
    }
  }

  /** Son satırlar (gerekirse bir önceki dosyadan da). */
  tail(count = 60): string[] {
    const out: string[] = [];
    for (let i = 0; i < this.files && out.length < count; i++) {
      try {
        const lines = fs.readFileSync(this.rotated(i), "utf8").split("\n").filter(Boolean);
        out.unshift(...lines.slice(-(count - out.length)));
      } catch {
        // dosya yok
      }
    }
    return out.slice(-count);
  }
}
