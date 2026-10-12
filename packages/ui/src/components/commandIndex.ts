/** Komut araması: dizin girdisi, Türkçe duyarlı sıralama, son kullanılanlar ve Dosya komutları. */

export interface CommandEntry {
  id: string;
  title: string;
  desc?: string;
  shortcut?: string;
  tabId: string;
  tabLabel: string;
  group: string;
  /** Aynı grup + başlıktaki kaçıncı öğe. */
  nth: number;
  /** Şeritte düğmesi olmayan komutlar (Dosya menüsü, galeri öğeleri) doğrudan çalışır. */
  run?: () => void;
  /** Şerit düğmesinin simgesi (SVG işaretlemesi; dizin çıkarılırken alınır). */
  icon?: string;
}

/** Karşılaştırma için: Türkçe küçük harf + aksan/şapka farkı yok sayılır ("KALIN" = "kalın" = "kalin"). */
export function foldText(s: string): string {
  return s
    .toLocaleLowerCase("tr-TR")
    .replace(/ı/g, "i")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

/** Sorgudaki her sözcüğün geçtiği komutları puanla sıralar. */
export function rankCommands(list: CommandEntry[], query: string): CommandEntry[] {
  const words = foldText(query).split(/\s+/).filter(Boolean);
  if (!words.length) return [];
  const scored: { c: CommandEntry; score: number }[] = [];
  for (const c of list) {
    const title = foldText(c.title);
    const desc = foldText(c.desc ?? "");
    const ctx = foldText(`${c.tabLabel} ${c.group}`);
    let score = 0;
    let ok = true;
    for (const w of words) {
      if (title.startsWith(w)) score += 10;
      else if (title.split(/[\s/-]+/).some((t) => t.startsWith(w))) score += 7;
      else if (title.includes(w)) score += 5;
      else if (desc.includes(w)) score += 2;
      else if (ctx.includes(w)) score += 1;
      else {
        ok = false;
        break;
      }
    }
    if (!ok) continue;
    if (title === foldText(query)) score += 20;
    scored.push({ c, score: score - title.length / 100 });
  }
  return scored.sort((a, b) => b.score - a.score).map((x) => x.c);
}

/* ---------------- Son kullanılanlar ---------------- */

export const RECENT_COMMANDS_LIMIT = 5;

/** Saklanan son kullanılan komut (çalıştırma işlevi saklanmaz; dizinden yeniden bulunur). */
export interface RecentCommand {
  id: string;
  title: string;
  tabLabel: string;
}

export const recentCommandsKey = (app: string) => `oh-cmd-recent:${app}`;

/** Komutu listenin başına alır; aynı komut tekrarlanmaz, en çok RECENT_COMMANDS_LIMIT kayıt kalır. */
export function pushRecentCommand(list: RecentCommand[], c: Pick<CommandEntry, "id" | "title" | "tabLabel">): RecentCommand[] {
  const item: RecentCommand = { id: c.id, title: c.title, tabLabel: c.tabLabel };
  return [item, ...list.filter((r) => r.id !== c.id && !(r.title === c.title && r.tabLabel === c.tabLabel))].slice(0, RECENT_COMMANDS_LIMIT);
}

/** Saklanan kayıtları güncel dizindeki komutlara eşler (önce kimlik, sonra sekme + ad). */
export function resolveRecentCommands(list: RecentCommand[], index: CommandEntry[]): CommandEntry[] {
  const byId = new Map(index.map((c) => [c.id, c]));
  const byName = new Map(index.map((c) => [`${c.tabLabel}|${c.title}`, c]));
  const out: CommandEntry[] = [];
  for (const r of list) {
    const c = byId.get(r.id) ?? byName.get(`${r.tabLabel}|${r.title}`);
    if (c && !out.includes(c)) out.push(c);
  }
  return out.slice(0, RECENT_COMMANDS_LIMIT);
}

function isRecentCommand(x: unknown): x is RecentCommand {
  const r = x as RecentCommand;
  return !!r && typeof r.id === "string" && typeof r.title === "string" && typeof r.tabLabel === "string";
}

export function loadRecentCommands(app: string, storage: Pick<Storage, "getItem"> | null = safeStorage()): RecentCommand[] {
  try {
    const raw = storage?.getItem(recentCommandsKey(app));
    const v: unknown = raw ? JSON.parse(raw) : [];
    return Array.isArray(v) ? v.filter(isRecentCommand).slice(0, RECENT_COMMANDS_LIMIT) : [];
  } catch {
    return [];
  }
}

export function saveRecentCommands(app: string, list: RecentCommand[], storage: Pick<Storage, "setItem"> | null = safeStorage()): void {
  try {
    storage?.setItem(recentCommandsKey(app), JSON.stringify(list.slice(0, RECENT_COMMANDS_LIMIT)));
  } catch {
    /* depolama kapalı ya da dolu: son kullanılanlar bu oturumda kalır */
  }
}

function safeStorage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

/* ---------------- Dosya komutları ---------------- */

export interface FileCommandActions {
  newDoc: () => void;
  open: () => void;
  save: () => void;
  saveAs: () => void;
  exportPdf: () => void;
  print: () => void;
  /** Dosya görünümünü ilgili sayfada açar. */
  backstage: (page: "info" | "export" | "options") => void;
}

const fileEntry = (id: string, title: string, run: () => void, extra: Partial<CommandEntry> = {}): CommandEntry => ({ id: `file:${id}`, title, tabId: "", tabLabel: "Dosya", group: "", nth: 0, run, ...extra });

/**
 * Şeritte düğmesi olmayan Dosya menüsü komutları (Yeni, Aç, Kaydet, Farklı Kaydet,
 * PDF, Yazdır, Seçenekler…). `noun` yeni öğenin adıdır (ör. "Belge", "Çalışma Kitabı").
 */
export function fileCommands(h: FileCommandActions, noun: string, newShortcut = "Ctrl+N"): CommandEntry[] {
  return [
    fileEntry("new", `Yeni ${noun}`, h.newDoc, { shortcut: newShortcut, desc: `Yeni boş ${noun.toLocaleLowerCase("tr-TR")} oluştur` }),
    fileEntry("open", "Aç", h.open, { shortcut: "Ctrl+O", desc: "Dosya aç" }),
    fileEntry("save", "Kaydet", h.save, { shortcut: "Ctrl+S" }),
    fileEntry("saveas", "Farklı Kaydet", h.saveAs, { shortcut: "F12", desc: "Kopyasını başka ad ya da konumla kaydet" }),
    fileEntry("pdf", "PDF olarak dışarı aktar", h.exportPdf, { group: "Dışarı Aktar", desc: "PDF belgesi oluştur, dışa aktar" }),
    fileEntry("print", "Yazdır", h.print, { shortcut: "Ctrl+P", desc: "Baskı önizleme ve yazıcı ayarları" }),
    fileEntry("export", "Dışarı Aktar", () => h.backstage("export"), { desc: "Başka biçimde kaydet, dışa aktar" }),
    fileEntry("info", "Bilgi", () => h.backstage("info"), { desc: "Dosya özellikleri" }),
    fileEntry("options", "Seçenekler", () => h.backstage("options"), { desc: "Arayüz teması, yazım denetimi, ayarlar" }),
  ];
}
