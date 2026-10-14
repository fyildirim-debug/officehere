// Derlem ve bulanık test betiklerinin ortak yolları ve yardımcıları.
//
// Tüm yollar depo köküne göre çözülür; betikler hangi dizinden çağrılırsa çağrılsın çalışır.
// Çıktılar varsayılan olarak <kök>/.derlem/ altına yazılır (git'e girmez); DERLEM_CIKTI
// ortam değişkeniyle başka bir dizin verilebilir.
import { existsSync, readdirSync, statSync } from "node:fs";
import { basename, dirname, extname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

/** Depo kökü. */
export const KOK = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
/** Üretilen her şeyin (paketler, çıktılar, raporlar) yazıldığı dizin. */
export const CIKTI = process.env.DERLEM_CIKTI ? resolve(process.env.DERLEM_CIKTI) : join(KOK, ".derlem");
/** Depoya konmuş derlem alt kümesi (LibreOffice / python ile üretilmiş, bizim içeriğimiz). */
export const DERLEM = join(KOK, "test-derlem");
/** Derlem üreticilerinin kullandığı görseller. */
export const GORSEL = join(DERLEM, "gorsel");
/** Betikle üretilen (depoya konmayan) büyük derlem dosyaları. */
export const URETILEN = join(CIKTI, "uretilen");
export const RAPOR = join(CIKTI, "rapor");
export const PAKET = join(CIKTI, "paket");
/** Temel çizgi (CI'ın karşılaştırdığı kaydedilmiş sonuçlar). */
export const TEMEL = join(KOK, "scripts/derlem/temel.json");

/** Uygulamaların açtığı uzantılar (turda hangi dosyanın hangi uygulamayla açılacağını belirler). */
export const UZANTI = {
  belge: ["docx", "docm", "dotx", "doc", "odt", "ott", "rtf", "md", "html", "htm", "txt", "fodt"],
  tablo: ["xlsx", "xlsm", "xltx", "xltm", "xls", "ods", "ots", "csv", "tsv"],
  sunum: ["pptx", "pptm", "potx", "ppsx", "ppt", "odp", "otp"],
};
export const UYGULAMALAR = Object.keys(UZANTI);

export function uygulamaOf(dosya) {
  const e = extname(dosya).slice(1).toLowerCase();
  for (const [app, exts] of Object.entries(UZANTI)) if (exts.includes(e)) return app;
  return null;
}

/** Depodaki örnek dosya (fikstür) dizinleri: uygulamaların kendi test dosyaları da tura girer. */
export const FIKSTUR_DIZINLERI = [
  "apps/belge/src/io/__fixtures__",
  "apps/belge/src/io/__fixtures__/bicimler",
  "apps/sunum/src/io/__fixtures__",
  "apps/tablo/src/io/__fixtures__",
  "apps/tablo/test/fixtures",
  "apps/tablo/test/fixtures/openpyxl",
  "packages/core/test/fixtures",
];

/** Platformdan bağımsız göreli yol (anahtar olarak kullanılır). */
export const posix = (p) => p.split(sep).join("/");

function dosyalar(dizin) {
  if (!existsSync(dizin)) return [];
  return readdirSync(dizin)
    .filter((f) => statSync(join(dizin, f)).isFile())
    .sort();
}

/**
 * Turun dosya listesi. Anahtar, temel çizgide dosyayı tanımlar:
 *  - depodaki dosyalar için depo köküne göre yol (test-derlem/lo/lo-rapor.docx),
 *  - betikle üretilenler için "uretilen/lo/lo-buyuk.xls".
 * Üretilen dosyalardan depodaki alt kümede aynı adla bulunanlar atlanır (çift sayılmasın).
 */
export function derlemDosyalari({ uygulama, suzgec } = {}) {
  const out = [];
  const ekle = (yol, anahtar, kaynak) => {
    const app = uygulamaOf(yol);
    if (!app || (uygulama && uygulama !== app)) return;
    if (suzgec && !anahtar.includes(suzgec)) return;
    out.push({ app, yol, anahtar, kaynak });
  };
  const depodakiAdlar = new Set();
  if (existsSync(DERLEM)) {
    for (const alt of readdirSync(DERLEM).sort()) {
      const d = join(DERLEM, alt);
      if (!statSync(d).isDirectory() || alt === "gorsel") continue;
      for (const f of dosyalar(d)) {
        depodakiAdlar.add(f);
        ekle(join(d, f), posix(relative(KOK, join(d, f))), "derlem");
      }
    }
  }
  if (existsSync(URETILEN)) {
    for (const alt of readdirSync(URETILEN).sort()) {
      const d = join(URETILEN, alt);
      if (!statSync(d).isDirectory()) continue;
      for (const f of dosyalar(d)) if (!depodakiAdlar.has(f)) ekle(join(d, f), `uretilen/${alt}/${f}`, "uretilen");
    }
  }
  for (const rel of FIKSTUR_DIZINLERI) for (const f of dosyalar(join(KOK, rel))) ekle(join(KOK, rel, f), `${rel}/${f}`, "fikstur");
  return out;
}

/** Basit argüman ayrıştırıcı: --ad değer, --ad=değer, --bayrak. */
export function argumanlar(argv = process.argv.slice(2)) {
  const a = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const s = argv[i];
    if (!s.startsWith("--")) {
      a._.push(s);
      continue;
    }
    const [k, v] = s.slice(2).split("=", 2);
    if (v !== undefined) a[k] = v;
    else if (argv[i + 1] && !argv[i + 1].startsWith("--")) a[k] = argv[++i];
    else a[k] = true;
  }
  return a;
}

export const ad = (p) => basename(p);
export const sure = (ms) => (ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(1)} sn`);
