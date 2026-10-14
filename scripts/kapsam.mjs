// Test kapsamı özeti (pnpm test:kapsam sonunda çalışır).
//
// Her paketin vitest kapsam özetini (<paket>/coverage/coverage-summary.json, @vitest/coverage-v8) okur;
// paket başına satır/dal/işlev kapsamını ve uygulamaların dosya okuyucu/yazıcıları (apps/*/src/io/**)
// için satır kapsamını yazar. Eşik yoktur, yalnız rapor: coverage/ozet.md (CI'da adım özetine de eklenir).
import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, relative, sep } from "node:path";
import { fileURLToPath } from "node:url";

const KOK = join(dirname(fileURLToPath(import.meta.url)), "..");
const yuzde = (k, t) => (t ? `${((100 * k) / t).toFixed(1)}` : "-");
const satirlar = [];
const io = { kapsanan: 0, toplam: 0 };
for (const ust of ["apps", "packages"]) {
  for (const ad of readdirSync(join(KOK, ust)).sort()) {
    const yol = join(KOK, ust, ad, "coverage", "coverage-summary.json");
    if (!existsSync(yol)) continue;
    const ozet = JSON.parse(readFileSync(yol, "utf8"));
    const t = ozet.total;
    const s = { paket: `${ust}/${ad}`, satir: t.lines, dal: t.branches, islev: t.functions, io: { kapsanan: 0, toplam: 0 } };
    for (const [dosya, d] of Object.entries(ozet)) {
      if (dosya === "total") continue;
      const goreli = relative(join(KOK, ust, ad), dosya).split(sep).join("/");
      if (ust !== "apps" || !goreli.startsWith("src/io/")) continue;
      s.io.kapsanan += d.lines.covered;
      s.io.toplam += d.lines.total;
    }
    io.kapsanan += s.io.kapsanan;
    io.toplam += s.io.toplam;
    satirlar.push(s);
  }
}
if (!satirlar.length) {
  console.error("Kapsam özeti bulunamadı: önce `pnpm test:kapsam` (vitest --coverage) çalıştırın.");
  process.exit(1);
}
const L = ["# Test kapsamı", "", "Eşik yok; yalnız rapor. Satır kapsamı (kapsanan / toplam).", ""];
L.push("| paket | satır % | dal % | işlev % | `src/io/**` satır % |", "|---|---|---|---|---|");
for (const s of satirlar) {
  const ioHucre = s.io.toplam ? `${yuzde(s.io.kapsanan, s.io.toplam)} (${s.io.kapsanan}/${s.io.toplam})` : "-";
  L.push(`| ${s.paket} | ${yuzde(s.satir.covered, s.satir.total)} (${s.satir.covered}/${s.satir.total}) | ${yuzde(s.dal.covered, s.dal.total)} | ${yuzde(s.islev.covered, s.islev.total)} | ${ioHucre} |`);
}
L.push("", `**\`apps/*/src/io/**\` toplam satır kapsamı: %${yuzde(io.kapsanan, io.toplam)}** (${io.kapsanan}/${io.toplam})`, "");
const md = L.join("\n");
mkdirSync(join(KOK, "coverage"), { recursive: true });
writeFileSync(join(KOK, "coverage", "ozet.md"), md);
if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, md + "\n");
console.log(md);
