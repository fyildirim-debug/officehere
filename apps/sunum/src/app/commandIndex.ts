/** Komut araması: dizin girdisi ve Türkçe duyarlı sıralama. */

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
}

/** Karşılaştırma için: Türkçe küçük harf + aksan/şapka farkı yok sayılır. */
export function fold(s: string): string {
  return s
    .toLocaleLowerCase("tr-TR")
    .replace(/ı/g, "i")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

/** Sorgudaki her sözcüğün geçtiği komutları puanla sıralar. */
export function rankCommands(list: CommandEntry[], query: string): CommandEntry[] {
  const words = fold(query).split(/\s+/).filter(Boolean);
  if (!words.length) return [];
  const scored: { c: CommandEntry; score: number }[] = [];
  for (const c of list) {
    const title = fold(c.title);
    const desc = fold(c.desc ?? "");
    const ctx = fold(`${c.tabLabel} ${c.group}`);
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
    if (title === fold(query)) score += 20;
    scored.push({ c, score: score - title.length / 100 });
  }
  return scored.sort((a, b) => b.score - a.score).map((x) => x.c);
}
