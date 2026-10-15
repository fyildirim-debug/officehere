// Yönetmeliğe uygunluk denetimi: belgeyi "Resmî Yazışmalarda Uygulanacak Usul ve Esaslar Hakkında Yönetmelik"
// (RG 10.06.2020/31151) ölçülerine göre denetler. Saf işlev: belge modeli (içerik JSON'u + ayarlar) alır, kural
// sonuçlarını döndürür; arayüz (DenetimPaneli) paragraf sırasını düzenleyicideki konuma çevirir.
import type { JSONContent } from "@tiptap/core";
import { twToCm, type BelgeDocument, type DocSettings, type PageSetup } from "../model/document";
import { effectiveProps, resolveStyles, styleIdOf, type StyleDef } from "../model/styles";
import type { Yon } from "../templates/resmi/alanlar";
import { A4, EN_KUCUK_PUNTO, GIRINTI_PT, ILETISIM_PUNTO, KENAR_CM, LOGO_UST_CM, YAZI_TIPLERI } from "../templates/resmi/olcu";

export type KuralKimligi = "kenar" | "yaziTipi" | "konu" | "sayiBicimi" | "tarihBicimi" | "kapanis" | "paragraf";
export type Durum = "uygun" | "bilgi" | "uyari" | "hata";
export type Bolge = "govde" | "altBilgi" | "ustBilgi";
export type DenetimYonu = Yon | "otomatik";

export interface Bulgu {
  durum: Exclude<Durum, "uygun">;
  ileti: string;
  /** Bölgedeki paragraf sırası (paragraf ve başlık düğümleri, belge sırasıyla). */
  paragraf?: number;
  bolge?: Bolge;
}

export interface KuralSonucu {
  kural: KuralKimligi;
  ad: string;
  madde: string;
  durum: Durum;
  /** Kural uygunsa kısa açıklama. */
  ozet: string;
  bulgular: Bulgu[];
}

export interface Kosu {
  metin: string;
  yaziTipi?: string;
  punto?: number;
}

export interface Paragraf {
  indeks: number;
  bolge: Bolge;
  tur: string;
  metin: string;
  attrs: Record<string, unknown>;
  kosular: Kosu[];
  tabloda: boolean;
  kutuda: boolean;
}

export const KURAL_ADLARI: Record<KuralKimligi, { ad: string; madde: string }> = {
  kenar: { ad: "Kâğıt ve kenar boşlukları", madde: "m.6, m.8" },
  yaziTipi: { ad: "Yazı tipi ve punto", madde: "m.7" },
  konu: { ad: "Konu satırı", madde: "m.13" },
  sayiBicimi: { ad: "Sayıların yazımı", madde: "m.16/7" },
  tarihBicimi: { ad: "Tarih biçimi", madde: "m.12" },
  kapanis: { ad: "Arz / rica uyumu", madde: "m.16/12" },
  paragraf: { ad: "Paragraf düzeni", madde: "m.16/4" },
};

const sayi = (n: number, basamak = 2) => n.toLocaleString("tr-TR", { maximumFractionDigits: basamak });
const kucukHarf = (s: string) => s.toLocaleLowerCase("tr-TR");

/* ------------------------------------------------------------------ */
/* Paragraf çıkarma                                                     */
/* ------------------------------------------------------------------ */

function puntoOku(v: unknown): number | undefined {
  if (typeof v === "number") return v;
  if (typeof v !== "string") return undefined;
  const m = /^([\d.]+)\s*(pt|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  return m[2] === "px" ? (n * 3) / 4 : n;
}

function yaziTipiOku(v: unknown): string | undefined {
  if (typeof v !== "string" || !v.trim()) return undefined;
  return v
    .split(",")[0]
    .trim()
    .replace(/^["']|["']$/g, "");
}

/** Paragraf ve başlık düğümleri, belge sırasıyla (ProseMirror descendants sırasıyla aynı). */
export function paragraflar(json: JSONContent | null | undefined, bolge: Bolge = "govde"): Paragraf[] {
  const out: Paragraf[] = [];
  const gez = (n: JSONContent, tabloda: boolean, kutuda: boolean) => {
    if (n.type === "paragraph" || n.type === "heading") {
      const kosular: Kosu[] = [];
      let metin = "";
      for (const c of n.content ?? []) {
        const ts = c.marks?.find((m) => m.type === "textStyle")?.attrs;
        if (c.type === "text") {
          kosular.push({ metin: c.text ?? "", yaziTipi: yaziTipiOku(ts?.fontFamily), punto: puntoOku(ts?.fontSize) });
          metin += c.text ?? "";
        } else if (c.type === "hardBreak") metin += "\n";
        else if (c.type === "pageField") kosular.push({ metin: "#", yaziTipi: yaziTipiOku(ts?.fontFamily), punto: puntoOku(ts?.fontSize) });
      }
      out.push({ indeks: out.length, bolge, tur: n.type, metin, attrs: (n.attrs ?? {}) as Record<string, unknown>, kosular, tabloda, kutuda });
      return;
    }
    for (const c of n.content ?? []) gez(c, tabloda || n.type === "table", kutuda || n.type === "textBox");
  };
  if (json) gez(json, false, false);
  return out;
}

const SAYI_RE = /^Sayı[\t ]*:/u;
const KONU_RE = /^Konu[\t ]*:[\t ]*/u;
const dolu = (p: Paragraf) => p.metin.trim().length > 0;

/* ------------------------------------------------------------------ */
/* Kurallar                                                             */
/* ------------------------------------------------------------------ */

function sonuc(kural: KuralKimligi, ozet: string, bulgular: Bulgu[]): KuralSonucu {
  const sira: Durum[] = ["hata", "uyari", "bilgi"];
  const durum = sira.find((d) => bulgular.some((b) => b.durum === d)) ?? "uygun";
  return { kural, ...KURAL_ADLARI[kural], durum, ozet, bulgular };
}

/** m.6/1 A4; m.8/1 üst, sol, sağ 1,5 cm (logo varsa üst 0,5 cm). Alt kenar Yönetmelik'te yok. */
export function kenarKurali(settings: DocSettings, govde: JSONContent): KuralSonucu {
  const sayfalar: PageSetup[] = [settings.page];
  for (const n of govde.content ?? []) if (n.type === "section" && n.attrs?.page) sayfalar.push(n.attrs.page as PageSetup);
  // Logo: belgenin ilk birkaç bloğunda resim ya da metin kutusu.
  const ilk = (govde.content?.[0]?.type === "section" ? govde.content[0].content : govde.content) ?? [];
  const logoVar = JSON.stringify(ilk.slice(0, 4)).match(/"type":"(image|textBox)"/) !== null;
  const tol = 0.05;
  const bulgular: Bulgu[] = [];
  const goruldu = new Set<string>();
  const ekle = (b: Bulgu) => {
    if (goruldu.has(b.ileti)) return;
    goruldu.add(b.ileti);
    bulgular.push(b);
  };
  for (const pg of sayfalar) {
    const kisa = Math.min(pg.width, pg.height);
    const uzun = Math.max(pg.width, pg.height);
    if (Math.abs(kisa - A4.w) > 30 || Math.abs(uzun - A4.h) > 30)
      ekle({ durum: "uyari", ileti: `Kâğıt A4 değil (${sayi(twToCm(pg.width), 1)} × ${sayi(twToCm(pg.height), 1)} cm). Üst yazı A4 hazırlanır; ekler farklı boyutta olabilir (m.6).` });
    const m = pg.margin;
    for (const [ad, tw] of [
      ["Sol", m.left],
      ["Sağ", m.right],
    ] as const) {
      const cm = twToCm(tw);
      if (Math.abs(cm - KENAR_CM) > tol) ekle({ durum: "hata", ileti: `${ad} kenar boşluğu ${sayi(cm)} cm; 1,5 cm olmalı (m.8/1).` });
    }
    const ust = twToCm(m.top);
    if (Math.abs(ust - KENAR_CM) > tol) {
      if (Math.abs(ust - LOGO_UST_CM) <= tol && logoVar) continue;
      ekle({
        durum: "hata",
        ileti:
          Math.abs(ust - LOGO_UST_CM) <= tol
            ? "Üst kenar boşluğu 0,5 cm yalnız logo kullanıldığında olur; logo yoksa 1,5 cm (m.8/1)."
            : `Üst kenar boşluğu ${sayi(ust)} cm; 1,5 cm (logo varsa 0,5 cm) olmalı (m.8/1).`,
      });
    }
  }
  const ust = twToCm(settings.page.margin.top);
  return sonuc("kenar", Math.abs(ust - LOGO_UST_CM) <= tol ? "A4; üst kenar 0,5 cm (logo), sol ve sağ 1,5 cm." : "A4; üst, sol ve sağ kenar boşlukları 1,5 cm.", bulgular);
}

const YAZI_ADLARI = new Set<string>(Object.values(YAZI_TIPLERI).map((y) => y.ad));

/** m.7/1: Times New Roman 12 ya da Arial 11; yazı alanında en az 9, iletişim bilgilerinde en az 8 punto. */
export function yaziTipiKurali(settings: DocSettings, bolgeler: Paragraf[][]): KuralSonucu {
  const styles = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles);
  const normal = styles.Normal.props;
  const temelAd = normal.font ?? "";
  const temelPunto = normal.size ?? 11;
  const bulgular: Bulgu[] = [];
  const temel = Object.values(YAZI_TIPLERI).find((y) => y.ad === temelAd);
  if (!temel) bulgular.push({ durum: "hata", ileti: `Belgenin yazı tipi ${temelAd || "belirsiz"}; Times New Roman ya da Arial olmalı (m.7/1).` });
  else if (temel.punto !== temelPunto) bulgular.push({ durum: "hata", ileti: `${temelAd} ${sayi(temelPunto)} punto; ${temelAd} için ${temel.punto} punto esastır (m.7/1).` });
  const yanlisTip = new Map<string, Paragraf>();
  const kullanilan = new Set<string>();
  let kucuk: Paragraf | null = null;
  let farkli: Paragraf | null = null;
  let farkliSay = 0;
  for (const liste of bolgeler)
    for (const p of liste) {
      if (!p.metin.trim() && !p.kosular.some((k) => k.metin === "#")) continue;
      const stil = effectiveProps(styles as Record<string, StyleDef>, styleIdOf(p.tur, p.attrs));
      for (const k of p.kosular) {
        if (!k.metin.trim()) continue;
        const ad = k.yaziTipi ?? stil.font ?? temelAd;
        const pt = k.punto ?? stil.size ?? temelPunto;
        kullanilan.add(ad);
        if (!YAZI_ADLARI.has(ad) && !yanlisTip.has(ad)) yanlisTip.set(ad, p);
        if (p.kutuda) continue;
        const alt = p.bolge === "govde" ? EN_KUCUK_PUNTO : ILETISIM_PUNTO;
        if (pt < alt - 0.01) kucuk ??= p;
        else if (p.bolge === "govde" && Math.abs(pt - temelPunto) > 0.01) {
          farkliSay++;
          farkli ??= p;
        }
      }
    }
  for (const [ad, p] of yanlisTip) bulgular.push({ durum: "hata", ileti: `“${ad}” yazı tipi kullanılmış; yalnız Times New Roman ya da Arial (m.7/1).`, paragraf: p.indeks, bolge: p.bolge });
  if (kucuk)
    bulgular.push({
      durum: "hata",
      ileti: kucuk.bolge === "govde" ? "Yazı alanında 9 puntodan küçük metin var (m.7/1)." : "İletişim bilgilerinde 8 puntodan küçük metin var (m.7/1).",
      paragraf: kucuk.indeks,
      bolge: kucuk.bolge,
    });
  if (farkli)
    bulgular.push({
      durum: "uyari",
      ileti: `${farkliSay} metin parçası farklı puntoda; başlık, sayı, konu, muhatap ve imza metinle aynı boyutta yazılır (Kılavuz).`,
      paragraf: farkli.indeks,
      bolge: "govde",
    });
  if (kullanilan.has("Times New Roman") && kullanilan.has("Arial")) bulgular.push({ durum: "uyari", ileti: "Belgede hem Times New Roman hem Arial var; tek yazı tipi kullanın." });
  return sonuc("yaziTipi", `${temelAd} ${sayi(temelPunto)} punto.`, bulgular);
}

/** TDK: başlıklarda küçük yazılan bağlaç ve edatlar. */
const BAGLACLAR = new Set(["ve", "ile", "veya", "yahut", "ya", "da", "de", "ki", "mı", "mi", "mu", "mü"]);
const NOKTALAMA_SONU = /[.,;:!?…\-–—()[\]"'’“”/]$/u;

/** Konu metninin yazım sorunları (m.13/1): boş, sonda noktalama, küçük harfle başlayan kelime. Şablon formu da kullanır. */
export function konuSorunlari(konuHam: string): string[] {
  const konu = konuHam.replace(/\s+/g, " ").trim();
  if (!konu) return ["Konu boş."];
  const out: string[] = [];
  if (NOKTALAMA_SONU.test(konu)) out.push(`Konunun sonunda noktalama işareti (“${konu.slice(-1)}”) var; konu noktalama işareti konmadan yazılır (m.13/1).`);
  const kucukler = konu.split(" ").filter((w, k) => {
    if (/^e-/u.test(w)) return false;
    const harf = w.match(/\p{L}/u)?.[0];
    if (!harf || !/^[\p{P}]*\p{L}/u.test(w)) return false;
    const kucuk = harf === kucukHarf(harf) && harf !== harf.toLocaleUpperCase("tr-TR");
    return kucuk && !(k > 0 && BAGLACLAR.has(kucukHarf(w.replace(/[^\p{L}]/gu, ""))));
  });
  if (kucukler.length)
    out.push(
      `Konudaki kelimelerin baş harfleri büyük yazılır: ${kucukler
        .slice(0, 4)
        .map((w) => `“${w}”`)
        .join(", ")} (m.13/1).`,
    );
  return out;
}

/** m.13/1: Konu, Sayı'nın bir alt satırında; kelimelerin baş harfleri büyük, sonunda noktalama yok. */
export function konuKurali(govde: Paragraf[]): KuralSonucu {
  const i = govde.findIndex((p) => !p.tabloda && KONU_RE.test(p.metin.trim()));
  if (i < 0) return sonuc("konu", "Konu satırı yok.", [{ durum: "bilgi", ileti: "“Konu:” satırı bulunamadı (dilekçe, tutanak gibi belgelerde gerekmeyebilir)." }]);
  const p = govde[i];
  const konu = p.metin.trim().replace(KONU_RE, "").replace(/\s+/g, " ").trim();
  const yer = { paragraf: p.indeks, bolge: "govde" as const };
  const bulgular: Bulgu[] = konuSorunlari(konu).map((ileti) => ({ durum: "hata", ileti, ...yer }));
  const once = govde.slice(0, i).filter(dolu).pop();
  if (!once || !SAYI_RE.test(once.metin.trim())) bulgular.push({ durum: "uyari", ileti: "Konu, “Sayı:” satırının bir alt satırına yazılır (m.13/1).", ...yer });
  return sonuc("konu", `Konu: ${konu}`, bulgular);
}

const grupla = (rakam: string) => rakam.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const BIRIM_RE = /^\s*(TL|₺|lira|kuruş|adet|kişi|km|kg|ton|metre|m²|m2|öğrenci|personel|sayfa|dönüm|hektar|litre)(?![\p{L}])/iu;

/** Sayı belirtecini denetler; yanlışsa doğru yazımı döndürür, doğru ya da kapsam dışıysa null. */
export function sayiDuzelt(belirtec: string, sonrasi = ""): string | null {
  if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(belirtec)) return null;
  if (/^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(belirtec)) return null; // tarih (ayrı kural)
  if (/^\d{1,2}\.\d{2}$/.test(belirtec)) return null; // saat (14.30) ya da belirsiz
  if (/^\d{3}(\.\d{2})+$/.test(belirtec)) return null; // standart dosya planı kodu
  // İngilizce yazım: 10,545.72 → 10.545,72
  let m = /^(\d{1,3}(?:,\d{3})+)(?:\.(\d+))?$/.exec(belirtec);
  if (m) return `${grupla(m[1].replace(/,/g, ""))}${m[2] ? `,${m[2]}` : ""}`;
  // Ondalık nokta: 10545.72 → 10.545,72
  m = /^(\d{3,})\.(\d{1,2})$/.exec(belirtec);
  if (m) return `${grupla(m[1])},${m[2]}`;
  m = /^(\d{4,})(?:,(\d+))?$/.exec(belirtec);
  if (!m) return null;
  const tam = m[1];
  if (tam.startsWith("0") || tam.length >= 10) return null; // posta kodu, telefon, kimlik numarası
  if (/^\s*sayılı/u.test(sonrasi)) return null; // kanun ve karar numaraları
  if (tam.length === 4 && !m[2] && !BIRIM_RE.test(sonrasi)) return null; // yıl ya da kod olabilir
  return `${grupla(tam)}${m[2] ? `,${m[2]}` : ""}`;
}

/** m.16/7: dört ve daha çok haneli sayılar üçlü gruplanıp nokta ile ayrılır; kesirler virgülle (1.452; 10.545,72). */
export function sayiBicimiKurali(govde: Paragraf[]): KuralSonucu {
  const bulgular: Bulgu[] = [];
  for (const p of govde) {
    const metin = p.metin;
    if (SAYI_RE.test(metin.trim())) continue;
    for (const m of metin.matchAll(/\d[\d.,]*/g)) {
      const ham = m[0];
      const belirtec = ham.replace(/[.,]+$/, "");
      const bas = m.index ?? 0;
      const once = metin[bas - 1] ?? "";
      const sonra = metin.slice(bas + belirtec.length);
      if (/[\p{L}\p{N}\-/]/u.test(once) || /^[\p{L}\-/]/u.test(sonra)) continue;
      const dogru = sayiDuzelt(belirtec, sonra);
      if (dogru && bulgular.length < 12) bulgular.push({ durum: "hata", ileti: `“${belirtec}” → “${dogru}” biçiminde yazılır (m.16/7).`, paragraf: p.indeks, bolge: "govde" });
    }
  }
  return sonuc("sayiBicimi", "Sayılar 1.452 ve 10.545,72 biçiminde.", bulgular);
}

const AYLAR = ["ocak", "şubat", "mart", "nisan", "mayıs", "haziran", "temmuz", "ağustos", "eylül", "ekim", "kasım", "aralık"];
const SAYISAL_TARIH = /(?<![\d.\-/])(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})(?![\d\-/]|\.\d)/g;
const YAZILI_TARIH = new RegExp(`(?<![\\p{L}\\d])(\\d{1,2})\\s+(${AYLAR.join("|")})\\s+(\\d{4})(?!\\d)`, "gu");

/** Tarih belirtecini denetler: GG.AA.YYYY ise null, değilse { ileti, öneri }. */
export function tarihDuzelt(ham: string): { oneri: string | null } | null {
  const m = /^(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})$/.exec(ham);
  if (!m) return null;
  const [, a, ayr, b, c] = m;
  const iki = (s: string) => s.padStart(2, "0");
  if (ayr === "." && a.length === 2 && b.length === 2 && c.length === 4) return null;
  if (a.length === 4 && c.length <= 2) return { oneri: `${iki(c)}.${iki(b)}.${a}` };
  if (a.length <= 2 && c.length === 4) return { oneri: `${iki(a)}.${iki(b)}.${c}` };
  return { oneri: null };
}

function tarihBul(metin: string): { ham: string; dogru: boolean }[] {
  const out: { ham: string; dogru: boolean }[] = [];
  for (const m of metin.matchAll(SAYISAL_TARIH)) {
    const gun = m[1].length === 4 ? +m[4] : +m[1];
    const ay = +m[3];
    if (ay < 1 || ay > 12 || gun < 1 || gun > 31) continue;
    out.push({ ham: m[0], dogru: tarihDuzelt(m[0]) === null });
  }
  const kucuk = kucukHarf(metin);
  for (const m of kucuk.matchAll(YAZILI_TARIH)) {
    const bas = m.index ?? 0;
    const ozgun = metin.slice(bas, bas + m[0].length);
    const ayIlk = ozgun.charAt(m[0].indexOf(m[2], m[1].length));
    out.push({ ham: ozgun, dogru: ayIlk !== kucukHarf(ayIlk) });
  }
  return out;
}

/** m.12/1: tarih sayı ile aynı satırda, yazı alanının en sağında; GG.AA.YYYY (01.09.2019) ya da 10 Ekim 2019. */
export function tarihBicimiKurali(govde: Paragraf[]): KuralSonucu {
  const bulgular: Bulgu[] = [];
  const si = govde.findIndex((p) => !p.tabloda && SAYI_RE.test(p.metin.trim()));
  if (si >= 0) {
    const sp = govde[si];
    const tarihler = tarihBul(sp.metin);
    if (!tarihler.length) {
      const yalniz = govde.slice(Math.max(0, si - 6), si + 8).find((p) => p !== sp && dolu(p) && tarihBul(p.metin).length === 1 && p.metin.trim() === tarihBul(p.metin)[0].ham);
      if (yalniz) bulgular.push({ durum: "hata", ileti: "Tarih, sayı ile aynı satırda ve yazı alanının en sağında yazılır (m.12/1).", paragraf: yalniz.indeks, bolge: "govde" });
      else bulgular.push({ durum: "bilgi", ileti: "Sayı satırında tarih yok. Elektronik imzalı belgede tarih imzanın zaman damgasından alınır (m.12/2).", paragraf: sp.indeks, bolge: "govde" });
    } else {
      const son = tarihler[tarihler.length - 1].ham;
      const once = sp.metin.slice(0, sp.metin.lastIndexOf(son));
      if (!sp.metin.trimEnd().endsWith(son) || !/\t\s*$/.test(once))
        bulgular.push({ durum: "uyari", ileti: "Tarih, sayı satırında yazı alanının en sağında olmalı (sağ sekme durağıyla) (m.12/1).", paragraf: sp.indeks, bolge: "govde" });
    }
  }
  for (const p of govde)
    for (const tr of tarihBul(p.metin)) {
      if (tr.dogru || bulgular.length >= 12) continue;
      const d = tarihDuzelt(tr.ham);
      const oneri = d?.oneri ?? (/\p{L}/u.test(tr.ham) ? tr.ham.replace(/\s(\p{Ll})/u, (_, h: string) => ` ${h.toLocaleUpperCase("tr-TR")}`) : null);
      bulgular.push({
        durum: "hata",
        ileti: d ? `“${tr.ham}” tarihi gün ve ay iki, yıl dört haneli, aralarında nokta ile yazılır${oneri ? `: “${oneri}”` : ""} (m.12/1).` : `Ay adı büyük harfle başlar: “${oneri}” (m.12/1).`,
        paragraf: p.indeks,
        bolge: "govde",
      });
    }
  return sonuc("tarihBicimi", "Tarihler GG.AA.YYYY biçiminde; tarih sayı ile aynı satırda.", bulgular);
}

const KAPANIS_RE = /(arz ve rica ederim|arz\/rica ederim|arz ederim|rica ederim|saygılarımla|iyi dileklerimle|bilgilerinize sunulur)\.?\s*$/iu;
type KapanisTuru = "arz" | "rica" | "karma" | "kisi";

function kapanisTuru(s: string): KapanisTuru | null {
  const k = kucukHarf(s.trim());
  const m = KAPANIS_RE.exec(k);
  if (!m) return null;
  if (/arz ve rica|arz\/rica/.test(m[1])) return "karma";
  if (m[1].startsWith("arz")) return "arz";
  if (m[1].startsWith("rica")) return "rica";
  return "kisi";
}

/** Muhatap: konudan sonraki (konu yoksa belgedeki) ilk ortalı satır. */
export function muhatapBul(govde: Paragraf[]): Paragraf | null {
  const ki = govde.findIndex((p) => !p.tabloda && KONU_RE.test(p.metin.trim()));
  const baslik = new Set(["T.C."]);
  return govde.slice(ki + 1).find((p) => !p.tabloda && !p.kutuda && dolu(p) && p.attrs.textAlign === "center" && !baslik.has(p.metin.trim())) ?? null;
}

/** Otomatik kipte muhataptan düzey çıkarımı. */
export function yonCikar(muhatap: string | null): Yon | "dagitim" | null {
  if (!muhatap) return null;
  const m = muhatap.trim();
  if (/^Sayın\s/u.test(m)) return "kisi";
  const b = m.toLocaleUpperCase("tr-TR");
  if (/DAĞITIM YERLERİNE$/u.test(b)) return "dagitim";
  if (/MAKAMINA$/u.test(b)) return "ust";
  return null;
}

const BEKLENEN: Record<Yon, string> = { ust: "“… arz ederim.”", alt: "“… rica ederim.”", karma: "“… arz ve rica ederim.”", kisi: "“Saygılarımla.” ya da “Bilgilerinize sunulur.”" };
const YON_ADI: Record<Yon, string> = { ust: "üst ya da aynı düzey makam", alt: "alt makam", karma: "üst ve alt makamlar birlikte", kisi: "gerçek kişi" };

/** m.16/12: alt makama "rica", üst ve aynı düzey makama "arz", karma dağıtımda "arz ve rica", gerçek kişiye "Saygılarımla.". */
export function kapanisKurali(govde: Paragraf[], yon: DenetimYonu = "otomatik"): KuralSonucu {
  const muh = muhatapBul(govde);
  const konuVar = govde.some((p) => KONU_RE.test(p.metin.trim()));
  let kapanis: Paragraf | null = null;
  for (let i = govde.length - 1; i >= 0; i--) {
    const p = govde[i];
    if (p.tabloda || !dolu(p) || /^Uygun görüşle arz ederim\.?$/u.test(p.metin.trim())) continue;
    if (kapanisTuru(p.metin)) {
      kapanis = p;
      break;
    }
  }
  if (!kapanis) {
    return sonuc("kapanis", "Kapanış ibaresi yok.", [{ durum: konuVar ? "uyari" : "bilgi", ileti: "Metin “… arz ederim.” / “… rica ederim.” gibi bir kapanış ibaresiyle bitmiyor (m.16/12)." }]);
  }
  const tur = kapanisTuru(kapanis.metin)!;
  const etkin = yon === "otomatik" ? yonCikar(muh?.metin ?? null) : yon;
  const bulgular: Bulgu[] = [];
  const yer = { paragraf: kapanis.indeks, bolge: "govde" as const };
  if (tur !== "kisi" && !kapanis.metin.trim().endsWith(".")) bulgular.push({ durum: "uyari", ileti: "Kapanış ibaresi nokta ile biter (“… arz ederim.”).", ...yer });
  if (etkin === null || etkin === "dagitim") {
    if (tur === "kisi") bulgular.push({ durum: "uyari", ileti: "Makama yazılan yazı “arz ederim” ya da “rica ederim” ile bitirilir (m.16/12).", ...yer });
    else
      bulgular.push({
        durum: "bilgi",
        ileti:
          etkin === "dagitim"
            ? "Dağıtımlı yazı: muhataplar üst ve alt makamlardan oluşuyorsa “arz ve rica ederim.” kullanılır (m.16/12-b). Düzeyi yukarıdan seçerek denetleyin."
            : "Muhatabın düzeyi anlaşılamadı; arz/rica uyumu için muhatap düzeyini seçin.",
        ...yer,
      });
  } else {
    const uygun = (etkin === "ust" && tur === "arz") || (etkin === "alt" && tur === "rica") || (etkin === "karma" && tur === "karma") || (etkin === "kisi" && tur === "kisi");
    if (!uygun)
      bulgular.push({ durum: etkin === "kisi" ? "uyari" : "hata", ileti: `Muhatap ${YON_ADI[etkin]}: kapanış ${BEKLENEN[etkin]} olmalı; belgede “${kapanis.metin.trim()}” (m.16/12).`, ...yer });
  }
  return sonuc("kapanis", `Kapanış: “${kapanis.metin.trim()}”`, bulgular);
}

const METIN_DISI = /^((Sayı|Konu|İlgi|Ek|Dağıtım|Gereği|Bilgi|Adres|Telefon|e-Posta)[\t ]*:|T\.C\.$)/u;

/** m.16/4: paragraf 1,25 cm içeriden başlar, metin iki yana yaslanır, paragraflar arasında boşluk bırakılmaz. */
export function paragrafKurali(settings: DocSettings, govde: Paragraf[]): KuralSonucu {
  const styles = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles) as Record<string, StyleDef>;
  const sorun = { girinti: [] as Paragraf[], yasla: [] as Paragraf[], aralik: [] as Paragraf[] };
  for (const p of govde) {
    const s = p.metin.trim();
    if (p.tabloda || p.kutuda || p.tur !== "paragraph" || METIN_DISI.test(s)) continue;
    const a = p.attrs;
    if (a.textAlign === "center" || a.textAlign === "right" || (a.indentLeft as number | null)) continue;
    if (s.length < 60 && !kapanisTuru(s)) continue;
    const stil = effectiveProps(styles, styleIdOf(p.tur, a));
    if (Math.abs(((a.indentFirst as number | null) ?? 0) - GIRINTI_PT) > 1) sorun.girinti.push(p);
    if ((a.textAlign ?? stil.align) !== "justify" && s.length >= 60) sorun.yasla.push(p);
    if (((a.spaceBefore as number | null) ?? stil.spaceBefore ?? 0) > 0.5 || ((a.spaceAfter as number | null) ?? stil.spaceAfter ?? 0) > 0.5) sorun.aralik.push(p);
  }
  const bulgular: Bulgu[] = [];
  const ekle = (liste: Paragraf[], ileti: string) => liste.length && bulgular.push({ durum: "uyari", ileti: `${liste.length} paragraf: ${ileti}`, paragraf: liste[0].indeks, bolge: "govde" });
  ekle(sorun.girinti, "ilk satır 1,25 cm içeriden başlamalı (m.16/4).");
  ekle(sorun.yasla, "metin iki yana yaslanmalı (m.16/4).");
  ekle(sorun.aralik, "paragraflar arasında satır boşluğu bırakılmaz (m.16/4).");
  return sonuc("paragraf", "Paragraflar 1,25 cm girintili, iki yana yaslı, aralıksız.", bulgular);
}

/* ------------------------------------------------------------------ */
/* Tümü                                                                 */
/* ------------------------------------------------------------------ */

export interface DenetimGirdisi {
  content: JSONContent;
  settings: DocSettings;
}

export function resmiYaziDenetle(doc: DenetimGirdisi | BelgeDocument, secenek: { yon?: DenetimYonu } = {}): KuralSonucu[] {
  const govde = paragraflar(doc.content, "govde");
  const alt = paragraflar(doc.settings.footer, "altBilgi");
  const ust = paragraflar(doc.settings.header, "ustBilgi");
  return [
    kenarKurali(doc.settings, doc.content),
    yaziTipiKurali(doc.settings, [govde, ust, alt]),
    konuKurali(govde),
    sayiBicimiKurali(govde),
    tarihBicimiKurali(govde),
    kapanisKurali(govde, secenek.yon ?? "otomatik"),
    paragrafKurali(doc.settings, govde),
  ];
}
