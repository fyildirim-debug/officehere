// LibreOffice denetimi (pnpm derlem:lo): derlem turunun kaydettiği dosyaları LibreOffice'e açtırır.
//
// LibreOffice yoksa uyarıp çıkar (hata değil). Önce `pnpm derlem` çalışmış olmalı (çıktılar ve
// .derlem/rapor/derlem.json). lo_pdf.py her çıktıyı ve özgününü PDF'e çevirir; bu betik özeti
// (.derlem/rapor/lo_pdf.md) yazar ve temel çizgiyle karşılaştırır: LibreOffice'in artık açamadığı
// çıktı ya da özgünden fazla boş sayfa = gerileme (kip "siki"de hata). lo_pdf.py ayrıca özgün ↔ çıktı piksel
// farkını ölçer (lo_piksel.md bu özete eklenir; temel çizgisine göre gerileme her kipte hata).
//
// Kullanım: pnpm derlem:lo [--temelsiz] [--kip siki|baslangic]
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { argumanlar, KOK, RAPOR } from "./ortak.mjs";
import { bulguRaporu, kipBelirle, loKarsilastir, loOzeti, temelOku } from "./temel.mjs";

const A = argumanlar();
const eksik = [["soffice", ["--version"]], ["pdfinfo", ["-v"]], ["pdftotext", ["-v"]], ["pdfimages", ["-v"]]].filter(([k, a]) => spawnSync(k, a, { stdio: "ignore" }).error);
if (eksik.length) {
  console.log(`LibreOffice denetimi atlandı: ${eksik.map(([k]) => k).join(", ")} bulunamadı (apt: libreoffice-writer libreoffice-calc libreoffice-impress poppler-utils).`);
  process.exit(0);
}
if (!existsSync(join(RAPOR, "derlem.json"))) {
  console.error("Önce `pnpm derlem` çalıştırın (kaydedilen dosyalar ve derlem.json gerekli).");
  process.exit(1);
}
const r = spawnSync("python3", [join(KOK, "scripts/derlem/lo_pdf.py")], { stdio: "inherit", cwd: KOK });
// 3: piksel farkında temel çizgiye göre gerileme (raporlar yazıldı; özet yine çıkarılır, sonunda hata).
if (r.status !== 0 && r.status !== 3) process.exit(r.status ?? 1);

const { sureSn, satirlar } = JSON.parse(readFileSync(join(RAPOR, "lo_pdf.json"), "utf8"));
const L = ["# LibreOffice denetimi (OfficeHere çıktıları → PDF)", "", `${satirlar.length} çıktı · ${sureSn} sn`, ""];
const sorunlu = satirlar.filter((s) => s.ohHata || s.bosFazla || (s.kayipOrani ?? 0) > 0.05);
L.push(`- LibreOffice'in açamadığı çıktı: **${satirlar.filter((s) => s.ohHata).length}** (özgünü açılamayan: ${satirlar.filter((s) => s.ozgunHata).length})`);
L.push(`- Özgünden fazla boş sayfalı çıktı: **${satirlar.filter((s) => s.bosFazla).length}**`);
L.push(`- PDF metninde %5'ten çok kelime kaybı: **${satirlar.filter((s) => (s.kayipOrani ?? 0) > 0.05).length}**`, "");
L.push("| çıktı | sayfa (özgün→OH) | boyut | boş sayfa (özgün→OH) | kayıp kelime oranı | hata |", "|---|---|---|---|---|---|");
for (const s of satirlar) {
  const isaret = sorunlu.includes(s) ? "**" : "";
  L.push(`| ${isaret}${s.anahtar}${isaret} | ${s.ozgunSayfa ?? "-"}→${s.ohSayfa ?? "-"} | ${s.ozgunBoyut ?? ""}→${s.ohBoyut ?? ""} | ${s.ozgunBos ?? "-"}→${s.ohBos ?? "-"} | ${s.kayipOrani ?? "-"} | ${(s.ohHata ?? "").replace(/\|/g, "/").slice(0, 100)} |`);
}
let md = L.join("\n") + "\n";
let kod = 0;
if (!A.temelsiz) {
  const temel = temelOku();
  const { md: kmd, hataVar } = bulguRaporu("LibreOffice denetimi", loKarsilastir(loOzeti(satirlar), temel.lo ?? {}, kipBelirle(A, temel)));
  md += "\n" + kmd + "\n";
  if (hataVar) kod = 1;
}
// Özgün ↔ çıktı piksel farkı (lo_pdf.py yazar; kendi temel çizgisi scripts/derlem/lo_piksel_temel.json).
if (existsSync(join(RAPOR, "lo_piksel.md"))) md += "\n" + readFileSync(join(RAPOR, "lo_piksel.md"), "utf8");
if (r.status === 3) kod = 1;
writeFileSync(join(RAPOR, "lo_pdf.md"), md);
console.log(`\nLibreOffice denetimi: ${satirlar.length} çıktı, ${satirlar.filter((s) => s.ohHata).length} açılamadı, ${satirlar.filter((s) => s.bosFazla).length} fazladan boş sayfalı (${sureSn} sn). Rapor: ${join(RAPOR, "lo_pdf.md")}`);
process.exitCode = kod;
