// Kaynak denetimi (pnpm denetim): sayaç temelli, yalnız ARTIŞI yakalar.
//
// 1) Yerel ayarsız büyük/küçük harf dönüşümü: `.toUpperCase()` / `.toLowerCase()` Türkçede yanlış sonuç
//    verir ("istanbul".toUpperCase() → "ISTANBUL", "I".toLowerCase() → "i"). Kullanıcı metninde
//    `toLocaleUpperCase("tr-TR")` / `toLocaleLowerCase("tr-TR")` kullanılmalıdır. Bilerek ASCII olan
//    kullanımlar (dosya uzantısı, onaltılık renk, XML öznitelik adı…) satıra `// yerel-ayar: ascii`
//    yorumu eklenerek sayımdan çıkarılır.
// 2) `waitForTimeout` (uçtan uca testlerde sabit bekleme): kararsız testlerin başlıca nedeni; yerine
//    expect.poll / toHaveText / waitForFunction gibi koşul beklemeleri kullanılmalıdır.
//
// Mevcut kullanımlar dosya başına sayı olarak temel listededir (scripts/denetim/yerel-ayar-temel.json).
// Bir dosyada sayı artarsa ya da yeni dosyada kullanım çıkarsa denetim başarısız olur. Azalınca temel
// güncellenebilir: `pnpm denetim --guncelle` (artış varsa --zorla gerekir).
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, relative, sep } from "node:path";
import { fileURLToPath } from "node:url";

const KOK = join(dirname(fileURLToPath(import.meta.url)), "../..");
const TEMEL = join(KOK, "scripts/denetim/yerel-ayar-temel.json");
const args = new Set(process.argv.slice(2));

const ATLA = new Set(["node_modules", "dist", "dist-electron", "dist-web", "release", "coverage", "__fixtures__", "__tests__", "fixtures"]);
function gez(dizin, uzanti, out = []) {
  if (!existsSync(dizin)) return out;
  for (const f of readdirSync(dizin)) {
    if (ATLA.has(f) || f.startsWith(".")) continue;
    const p = join(dizin, f);
    if (statSync(p).isDirectory()) gez(p, uzanti, out);
    else if (uzanti.test(f)) out.push(p);
  }
  return out;
}
const rel = (p) => relative(KOK, p).split(sep).join("/");

const DENETIMLER = {
  yerelAyar: {
    ad: "yerel ayarsız toUpperCase()/toLowerCase()",
    dosyalar: () =>
      [...readdirSync(join(KOK, "apps")).map((a) => join(KOK, "apps", a, "src")), ...readdirSync(join(KOK, "packages")).map((p) => join(KOK, "packages", p, "src"))]
        .flatMap((d) => gez(d, /\.(ts|tsx|js|mjs)$/))
        .filter((f) => !/\.(test|spec)\.[cm]?[jt]sx?$/.test(f)),
    desen: /\.to(Upper|Lower)Case\(\s*\)/g,
    muaf: /yerel-ayar:\s*ascii/,
    oneri: 'Kullanıcı metninde toLocaleUpperCase("tr-TR") / toLocaleLowerCase("tr-TR") kullanın; bilerek ASCII ise satıra `// yerel-ayar: ascii` ekleyin.',
  },
  waitForTimeout: {
    ad: "uçtan uca testlerde waitForTimeout",
    dosyalar: () => [...gez(join(KOK, "e2e"), /\.[cm]?[jt]s$/), ...gez(join(KOK, "e2e-masaustu"), /\.[cm]?[jt]s$/)],
    desen: /\bwaitForTimeout\s*\(/g,
    muaf: null,
    oneri: "Sabit bekleme yerine koşul bekleyin: expect(...).toHaveText / expect.poll / page.waitForFunction.",
  },
};

function say() {
  const sonuc = {};
  const satirlar = {};
  for (const [k, d] of Object.entries(DENETIMLER)) {
    sonuc[k] = {};
    satirlar[k] = {};
    for (const f of d.dosyalar()) {
      let n = 0;
      const yerler = [];
      readFileSync(f, "utf8")
        .split("\n")
        .forEach((s, i) => {
          if (d.muaf?.test(s)) return;
          const m = s.match(d.desen)?.length ?? 0;
          if (m) {
            n += m;
            yerler.push(`${rel(f)}:${i + 1}: ${s.trim().slice(0, 120)}`);
          }
        });
      if (n) {
        sonuc[k][rel(f)] = n;
        satirlar[k][rel(f)] = yerler;
      }
    }
    sonuc[k] = Object.fromEntries(Object.entries(sonuc[k]).sort(([a], [b]) => a.localeCompare(b)));
  }
  return { sonuc, satirlar };
}

const { sonuc, satirlar } = say();
const temel = existsSync(TEMEL) ? JSON.parse(readFileSync(TEMEL, "utf8")) : {};
let hata = 0;
let azalan = 0;
for (const [k, d] of Object.entries(DENETIMLER)) {
  const t = temel[k] ?? {};
  const toplam = Object.values(sonuc[k]).reduce((a, b) => a + b, 0);
  const toplamT = Object.values(t).reduce((a, b) => a + b, 0);
  console.log(`${d.ad}: ${toplam} (temel ${toplamT}, ${Object.keys(sonuc[k]).length} dosya)`);
  for (const [f, n] of Object.entries(sonuc[k])) {
    if (n > (t[f] ?? 0)) {
      hata++;
      console.error(`  ARTIŞ ${f}: ${t[f] ?? 0} → ${n}`);
      for (const y of satirlar[k][f]) console.error(`    ${y}`);
    }
  }
  for (const [f, n] of Object.entries(t)) if ((sonuc[k][f] ?? 0) < n) azalan++;
  if (hata) console.error(`  → ${d.oneri}`);
}
if (args.has("--guncelle")) {
  if (hata && existsSync(TEMEL) && !args.has("--zorla")) {
    console.error("Temel güncellenmedi: artış var (bilerek kabul etmek için --zorla).");
    process.exit(1);
  }
  writeFileSync(
    TEMEL,
    JSON.stringify({ aciklama: "pnpm denetim temel sayıları (dosya başına). Yalnız azaltın; güncelleme: pnpm denetim --guncelle", ...sonuc }, null, 1) + "\n",
  );
  console.log(`temel güncellendi: ${rel(TEMEL)}`);
  process.exit(0);
}
if (azalan && !hata) console.log(`${azalan} dosyada sayı azaldı: \`pnpm denetim --guncelle\` ile temeli düşürün.`);
process.exit(hata ? 1 : 0);
