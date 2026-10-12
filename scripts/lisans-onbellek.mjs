// Yerelde kurulu olmayan (önceden derlenmiş paketlere gömülü) bileşenlerin
// lisans metinlerini npm kayıt deposundan indirip scripts/lisanslar/ altına yazar.
// Kullanım: node scripts/lisans-onbellek.mjs <paket> [<paket> ...]
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { ONBELLEK } from "./lisans-bildirimi.mjs";

const LISANS_DOSYASI = /^(licen[cs]e|copying|notice)([.\-_].*)?$/i;
const paketler = process.argv.slice(2);
if (!paketler.length) {
  console.error("Kullanım: node scripts/lisans-onbellek.mjs <paket> [<paket> ...]");
  process.exit(1);
}

fs.mkdirSync(ONBELLEK, { recursive: true });
let hata = 0;
for (const ad of paketler) {
  const gecici = fs.mkdtempSync(path.join(os.tmpdir(), "lisans-"));
  try {
    const bilgi = JSON.parse(execFileSync("npm", ["view", ad, "version", "license", "repository.url", "--json"], { encoding: "utf8" }));
    const arsiv = execFileSync("npm", ["pack", `${ad}@${bilgi.version}`, "--silent", "--pack-destination", gecici], { encoding: "utf8" }).trim().split("\n").pop();
    execFileSync("tar", ["-xzf", path.join(gecici, arsiv), "-C", gecici]);
    const kok = path.join(gecici, "package");
    const dosyalar = fs.readdirSync(kok).filter((f) => LISANS_DOSYASI.test(f)).sort();
    if (!dosyalar.length) {
      console.error(`${ad}: pakette lisans dosyası yok; metni elle scripts/lisanslar/${ad.replace("/", "+")}.txt olarak ekleyin.`);
      hata++;
      continue;
    }
    const metin = dosyalar.map((f) => fs.readFileSync(path.join(kok, f), "utf8").trim()).join("\n\n");
    const kaynak = (bilgi["repository.url"] ?? "").replace(/^git\+/, "");
    const bas = [`Paket: ${ad}@${bilgi.version}`, `Lisans: ${bilgi.license ?? ""}`, `Kaynak: ${kaynak}`].join("\n");
    fs.writeFileSync(path.join(ONBELLEK, `${ad.replace("/", "+")}.txt`), `${bas}\n\n${metin}\n`);
    console.log(`${ad}@${bilgi.version} (${bilgi.license}) yazıldı.`);
  } catch (e) {
    console.error(`${ad}: ${e.message}`);
    hata++;
  } finally {
    fs.rmSync(gecici, { recursive: true, force: true });
  }
}
process.exit(hata ? 1 : 0);
