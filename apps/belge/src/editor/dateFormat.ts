// Tarih/saat biçim resimleri (OOXML/RTF alan anahtarı \@ "…" ile aynı belirteçler): d dd ddd dddd,
// M MM MMM MMMM, yy yyyy, H HH h hh, m mm, s ss, AM/PM; tek tırnak içindeki metin olduğu gibi yazılır.
// Ay ve gün adları Türkçedir. Ekle › Tarih ve Saat'in biçim listesi ve otomatik güncellenen alanlar kullanır.

export const MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
export const DAYS = ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"];
const MONTHS_SHORT = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];
const DAYS_SHORT = ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"];

export type PicToken =
  | { kind: "lit"; text: string }
  | { kind: "day" | "month" | "year" | "hour24" | "hour12" | "minute" | "second" | "ampm"; len: number; upper?: boolean };

/** Biçim resmini belirteçlere ayırır. */
export function parsePicture(pic: string): PicToken[] {
  const out: PicToken[] = [];
  const lit = (text: string) => {
    const last = out[out.length - 1];
    if (last?.kind === "lit") last.text += text;
    else out.push({ kind: "lit", text });
  };
  let i = 0;
  while (i < pic.length) {
    const c = pic[i];
    if (c === "'") {
      const end = pic.indexOf("'", i + 1);
      lit(end < 0 ? pic.slice(i + 1) : pic.slice(i + 1, end));
      i = end < 0 ? pic.length : end + 1;
      continue;
    }
    if (/^am\/pm/i.test(pic.slice(i, i + 5))) {
      out.push({ kind: "ampm", len: 5, upper: pic[i] === "A" });
      i += 5;
      continue;
    }
    const kind = c === "d" || c === "D" ? "day" : c === "M" ? "month" : c === "y" || c === "Y" ? "year" : c === "H" ? "hour24" : c === "h" ? "hour12" : c === "m" ? "minute" : c === "s" || c === "S" ? "second" : null;
    if (!kind) {
      lit(c);
      i++;
      continue;
    }
    let len = 1;
    while (pic[i + len] === c) len++;
    out.push({ kind, len });
    i += len;
  }
  return out;
}

const pad = (n: number, w = 2) => String(n).padStart(w, "0");

/** Tarihi biçim resmine göre yazar. */
export function formatPicture(pic: string, d: Date): string {
  return parsePicture(pic)
    .map((t) => {
      switch (t.kind) {
        case "lit":
          return t.text;
        case "day":
          return t.len >= 4 ? DAYS[d.getDay()] : t.len === 3 ? DAYS_SHORT[d.getDay()] : t.len === 2 ? pad(d.getDate()) : String(d.getDate());
        case "month":
          return t.len >= 4 ? MONTHS[d.getMonth()] : t.len === 3 ? MONTHS_SHORT[d.getMonth()] : t.len === 2 ? pad(d.getMonth() + 1) : String(d.getMonth() + 1);
        case "year":
          return t.len <= 2 ? pad(d.getFullYear() % 100) : String(d.getFullYear());
        case "hour24":
          return t.len >= 2 ? pad(d.getHours()) : String(d.getHours());
        case "hour12": {
          const h = d.getHours() % 12 || 12;
          return t.len >= 2 ? pad(h) : String(h);
        }
        case "minute":
          return t.len >= 2 ? pad(d.getMinutes()) : String(d.getMinutes());
        case "second":
          return t.len >= 2 ? pad(d.getSeconds()) : String(d.getSeconds());
        case "ampm":
          return d.getHours() < 12 ? "ÖÖ" : "ÖS";
      }
    })
    .join("");
}

/** Yalnız saat belirteçleri içeren resim (TIME alanı olarak yazılır). */
export function isTimePicture(pic: string): boolean {
  const toks = parsePicture(pic);
  return toks.some((t) => t.kind !== "lit") && toks.every((t) => t.kind === "lit" || t.kind === "hour24" || t.kind === "hour12" || t.kind === "minute" || t.kind === "second" || t.kind === "ampm");
}

/** Alan yönergesindeki biçim resmi: DATE \@ "d MMMM yyyy" → d MMMM yyyy. */
export function pictureOfInstr(instr: string): string | null {
  const m = /\\@\s*(?:"([^"]*)"|(\S+))/.exec(instr);
  const pic = m ? (m[1] ?? m[2] ?? null) : null;
  // LibreOffice boşlukları ters eğik çizgiyle kaçışlar: d\ MMMM\ yyyy → d MMMM yyyy.
  return pic === null ? null : pic.replace(/\\(.)/g, "$1");
}

/** Ekle › Tarih ve Saat'teki biçimler (Türkçe yerleşik biçimlerle aynı sıra). */
export const DATE_PICTURES: string[] = [
  "dd.MM.yyyy",
  "d MMMM yyyy dddd",
  "d MMMM yyyy",
  "d.M.yyyy",
  "dd.MM.yy",
  "yyyy-MM-dd",
  "d MMM yy",
  "MMMM yyyy",
  "dd.MM.yyyy HH:mm",
  "dd.MM.yyyy HH:mm:ss",
  "HH:mm",
  "HH:mm:ss",
];
