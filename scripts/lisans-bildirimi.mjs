// Üçüncü taraf lisans bildirimi (Vite eklentisi).
// Pakete giren her npm bileşeninin lisans metnini toplayıp çıktıya
// THIRD_PARTY_NOTICES.txt olarak ekler. MIT/BSD/Apache gibi lisanslar
// dağıtılan her kopyada telif ve lisans metninin bulunmasını şart koşar.
//
// Önceden derlenmiş paketlerin (ör. exceljs.min.js) içine gömülü bileşenler
// kaynak haritasından (.map) bulunur ve lisans metinleri her ortamda aynı
// sonucu vermesi için yalnız scripts/lisanslar/ önbelleğinden okunur. Lisans
// dosyası olmayan doğrudan bağımlılıklar da oradan okunur. Eksik varsa derleme
// durur ve `node scripts/lisans-onbellek.mjs <paket>` ile önbellek doldurulur.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const ONBELLEK = path.join(path.dirname(fileURLToPath(import.meta.url)), "lisanslar");
const KOK = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const NM = "/node_modules/";
const LISANS_DOSYASI = /^(licen[cs]e|copying|notice)([.\-_].*)?$/i;
// Kaynak haritası olmayan hazır paketlerin içine gömülü bileşenler (dist dosyası incelenerek belirlendi).
const GOMULU_ELLE = { jszip: ["pako", "lie", "immediate", "setimmediate", "browser-pack"] };
// Derleyicinin pakete eklediği sanal yardımcı modüller (ör. \0vite/preload-helper.js).
const SANAL = /^\0(vite|rolldown)\//;

/** Modül kimliğinden paket adını ve kök dizinini çıkarır (son node_modules bölümü). */
export function paketBul(id) {
  const temiz = id.replace(/^\0/, "").split("?")[0].replace(/\\/g, "/");
  const i = temiz.lastIndexOf(NM);
  if (i < 0) return null;
  const parca = temiz.slice(i + NM.length).split("/");
  const ad = parca[0].startsWith("@") ? `${parca[0]}/${parca[1]}` : parca[0];
  if (!ad || ad === ".pnpm") return null;
  return { ad, dizin: temiz.slice(0, i + NM.length + ad.length), dosya: temiz };
}

const onbellekAdi = (ad) => ad.replace("/", "+");

function pkgOku(dizin) {
  try {
    return JSON.parse(fs.readFileSync(path.join(dizin, "package.json"), "utf8"));
  } catch {
    return null;
  }
}

function lisansAdi(pkg) {
  const l = pkg?.license ?? pkg?.licenses;
  if (typeof l === "string") return l;
  if (Array.isArray(l)) return l.map((x) => x.type ?? x).join(" OR ");
  if (l && typeof l === "object") return l.type ?? "";
  return "";
}

function lisansMetni(dizin) {
  let ad;
  try {
    ad = fs.readdirSync(dizin).filter((f) => LISANS_DOSYASI.test(f)).sort();
  } catch {
    return "";
  }
  // Vite'nin LICENSE.md'si kendi araç bağımlılıklarını da listeler; pakete girmeyen o bölüm alınmaz.
  return ad
    .map((f) => fs.readFileSync(path.join(dizin, f), "utf8").split(/\n#+ Licenses of bundled dependencies/)[0].trim())
    .filter(Boolean)
    .join("\n\n");
}

/** Yerel pnpm deposunda paketin herhangi bir sürümünü arar. */
function yereldeBul(ad) {
  const depo = path.join(KOK, "node_modules/.pnpm");
  let girdiler;
  try {
    girdiler = fs.readdirSync(depo);
  } catch {
    return null;
  }
  const on = `${onbellekAdi(ad)}@`;
  for (const g of girdiler.filter((x) => x.startsWith(on)).sort().reverse()) {
    const dizin = path.join(depo, g, "node_modules", ad);
    if (fs.existsSync(path.join(dizin, "package.json"))) return dizin;
  }
  return null;
}

/** Önbellek dosyası: "Lisans: X" başlığı, boş satır, metin. */
function onbellektenOku(ad) {
  const f = path.join(ONBELLEK, `${onbellekAdi(ad)}.txt`);
  if (!fs.existsSync(f)) return null;
  const ham = fs.readFileSync(f, "utf8");
  const bol = ham.indexOf("\n\n");
  const bas = ham.slice(0, bol);
  return {
    lisans: /^Lisans: (.*)$/m.exec(bas)?.[1] ?? "",
    kaynak: /^Kaynak: (.*)$/m.exec(bas)?.[1] ?? "",
    metin: ham.slice(bol + 2).trim(),
  };
}

/** Önceden derlenmiş dosyanın içine gömülü paket adları (kaynak haritasından ve elle listeden). */
export function gomuluPaketler(ad, dosya) {
  const adlar = new Set();
  const ekle = (x) => {
    if (adlar.has(x)) return;
    adlar.add(x);
    for (const y of GOMULU_ELLE[x] ?? []) ekle(y);
  };
  if (dosya.includes(`${NM}${ad}/dist/`)) for (const y of GOMULU_ELLE[ad] ?? []) ekle(y);
  const harita = `${dosya}.map`;
  if (fs.existsSync(harita)) {
    let kaynaklar = [];
    try {
      kaynaklar = JSON.parse(fs.readFileSync(harita, "utf8")).sources ?? [];
    } catch {
      kaynaklar = [];
    }
    for (const s of kaynaklar) {
      const p = paketBul(`/${s}`);
      if (p) ekle(p.ad);
    }
  }
  adlar.delete(ad);
  return [...adlar];
}

/** Paket kaydı: ad, sürüm, lisans, kaynak ve metin. Metin bulunamazsa null. */
function kayit(ad, dizin, icinde) {
  const pkg = dizin ? pkgOku(dizin) : null;
  let metin = dizin ? lisansMetni(dizin) : "";
  let lisans = lisansAdi(pkg);
  let kaynak = typeof pkg?.repository === "string" ? pkg.repository : (pkg?.repository?.url ?? pkg?.homepage ?? "");
  if (!metin) {
    const o = onbellektenOku(ad);
    if (!o) return null;
    metin = o.metin;
    lisans ||= o.lisans;
    kaynak ||= o.kaynak;
  }
  return { ad, surum: icinde ? "" : (pkg?.version ?? ""), lisans, kaynak: kaynak.replace(/^git\+/, ""), metin, icinde };
}

/** Pakete giren modül kimliklerinden bildirim kayıtlarını üretir. */
export function kayitlariTopla(modulKimlikleri) {
  const kayitlar = new Map();
  const eksik = [];
  for (const id of modulKimlikleri) {
    const sanal = SANAL.exec(id)?.[1];
    if (sanal) {
      if (kayitlar.has(sanal)) continue;
      const k = kayit(sanal, yereldeBul(sanal), "");
      if (k) kayitlar.set(sanal, k);
      else eksik.push(sanal);
      continue;
    }
    const p = paketBul(id);
    // Önce başka bir paketin içinde görülen bileşen doğrudan da kullanılıyorsa sürümlü kaydı yazılır.
    if (!p || (kayitlar.has(p.ad) && !kayitlar.get(p.ad).icinde)) continue;
    const k = kayit(p.ad, p.dizin, "");
    if (k) kayitlar.set(p.ad, k);
    else eksik.push(p.ad);
    for (const ic of gomuluPaketler(p.ad, p.dosya)) {
      if (kayitlar.has(ic)) continue;
      // Gömülü bileşenler yalnız depodaki önbellekten okunur; yerel pnpm deposu ortama göre değişir.
      const ki = kayit(ic, null, p.ad);
      if (ki) kayitlar.set(ic, ki);
      else eksik.push(ic);
    }
  }
  return {
    kayitlar: [...kayitlar.values()].sort((a, b) => a.ad.localeCompare(b.ad)),
    eksik: [...new Set(eksik)].filter((ad) => !kayitlar.has(ad)),
  };
}

/** Paket adını, Node'un çözümleme sırasıyla başlangıç dizininden yukarı doğru arar. */
function paketDiziniBul(ad, baslangic) {
  for (let d = baslangic; ; d = path.dirname(d)) {
    const aday = path.basename(d) === "node_modules" ? path.join(d, ad) : path.join(d, "node_modules", ad);
    if (fs.existsSync(path.join(aday, "package.json"))) return aday;
    if (path.dirname(d) === d) return null;
  }
}

/**
 * Bir paketin üretim bağımlılık ağacı (Electron ana sürecine paketlenen node_modules).
 * Dönen değerler package.json yollarıdır; kayitlariTopla bunları modül kimliği gibi işler.
 */
export function uretimBagimliliklari(kokDizin) {
  const dizinler = new Set();
  const ziyaret = (dizin) => {
    const pkg = pkgOku(dizin);
    for (const ad of Object.keys({ ...(pkg?.dependencies ?? {}), ...(pkg?.optionalDependencies ?? {}) })) {
      const d = paketDiziniBul(ad, dizin);
      if (!d) continue;
      const gercek = fs.realpathSync(d);
      if (dizinler.has(gercek)) continue;
      dizinler.add(gercek);
      ziyaret(gercek);
    }
  };
  ziyaret(kokDizin);
  return [...dizinler].map((d) => path.join(d, "package.json").replace(/\\/g, "/"));
}

/** Dağıtıma girebilecek izin verici lisanslar (SPDX). Karşılıklılık (copyleft) şartı taşıyanlar yoktur. */
export const IZINLI_LISANSLAR = new Set(["MIT", "ISC", "BSD-2-Clause", "BSD-3-Clause", "0BSD", "Apache-2.0", "Zlib", "BlueOak-1.0.0", "Unlicense", "CC0-1.0", "Python-2.0"]);

/** SPDX ifadesi izinli mi: "A OR B" seçeneklerinden biri, "A AND B" için hepsi izinli olmalı. */
export function lisansIzinli(ifade) {
  const temiz = (ifade ?? "").replace(/[()]/g, " ").trim();
  if (!temiz) return false;
  return temiz.split(/\s+OR\s+/i).some((secenek) => secenek.split(/\s+AND\s+/i).every((l) => IZINLI_LISANSLAR.has(l.trim())));
}

const CIZGI = "=".repeat(78);

export function bildirimMetni(kayitlar, urun = "OfficeHere") {
  const bas = [
    `${urun} — Üçüncü Taraf Yazılım Bildirimleri / Third-Party Software Notices`,
    "",
    `${urun} MIT lisanslıdır (bkz. LICENSE). Bu dağıtım aşağıdaki açık kaynak`,
    "bileşenleri içerir; her biri kendi lisansı altındadır ve telif hakları",
    "sahiplerine aittir. Birden çok lisans seçeneği sunan bileşenlerde (ör. MIT OR",
    "GPL-3.0) bu dağıtımda MIT seçeneği kullanılır.",
    "",
    "This distribution includes the open source components listed below, each",
    "under its own license. Where a component offers a choice of licenses (e.g.",
    "MIT OR GPL-3.0), the MIT option is used for this distribution.",
    "",
    `Bileşen sayısı / Components: ${kayitlar.length}`,
  ];
  const govde = kayitlar.map((k) => {
    const satir = [CIZGI, `${k.ad}${k.surum ? ` ${k.surum}` : ""} — ${k.lisans || "bkz. lisans metni"}`];
    if (k.icinde) satir.push(`(${k.icinde} paketinin içine gömülü / bundled inside ${k.icinde})`);
    if (k.kaynak) satir.push(k.kaynak);
    satir.push("-".repeat(78), k.metin);
    return satir.join("\n");
  });
  return `${[bas.join("\n"), ...govde].join("\n\n")}\n`;
}

/**
 * Vite eklentisi: derleme çıktısına THIRD_PARTY_NOTICES.txt ekler.
 * `uretimBagimliliklari`: yapılandırma dosyasına göre göreli bir paket dizini; o paketin
 * üretim bağımlılıkları da (ör. Electron ana sürecine giren electron-updater) eklenir.
 */
export function lisansBildirimi(secenekler = {}) {
  let yapilandirmaDizini = process.cwd();
  return {
    name: "officehere-lisans-bildirimi",
    apply: "build",
    configResolved(c) {
      if (c.configFile) yapilandirmaDizini = path.dirname(c.configFile);
    },
    generateBundle(_secenek, paket) {
      const kimlikler = new Set();
      if (secenekler.uretimBagimliliklari) {
        for (const id of uretimBagimliliklari(path.resolve(yapilandirmaDizini, secenekler.uretimBagimliliklari))) kimlikler.add(id);
      }
      for (const parca of Object.values(paket)) {
        if (parca.type !== "chunk") continue;
        for (const id of parca.moduleIds ?? Object.keys(parca.modules ?? {})) kimlikler.add(id);
      }
      const { kayitlar, eksik } = kayitlariTopla(kimlikler);
      const yasak = kayitlar.filter((k) => !lisansIzinli(k.lisans));
      if (yasak.length) {
        this.error(`İzin verilmeyen ya da belirsiz lisans: ${yasak.map((k) => `${k.ad} (${k.lisans || "lisans yok"})`).join(", ")}. Bkz. HUKUKI.md.`);
      }
      if (eksik.length) {
        this.error(`Lisans metni bulunamadı: ${eksik.join(", ")}. "node scripts/lisans-onbellek.mjs ${eksik.join(" ")}" ile önbelleğe alın.`);
      }
      this.emitFile({ type: "asset", fileName: "THIRD_PARTY_NOTICES.txt", source: bildirimMetni(kayitlar) });
    },
  };
}
