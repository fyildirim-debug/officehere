// Paket boyutu: dist JS/CSS boyutları (ham + gzip + brotli), parça sayısı, ilk yüklemede istenen dosyalar.
// Ön koşul: pnpm build. Kullanım: node scripts/derlem/olcum/paket_boyut.mjs → .derlem/rapor/paket_boyut.json
import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { extname, join } from "node:path";
import { brotliCompressSync, gzipSync } from "node:zlib";
import { KOK, RAPOR } from "../ortak.mjs";

const out = {};
const gez = (d, acc = []) => {
  for (const f of readdirSync(d)) {
    const p = join(d, f);
    if (statSync(p).isDirectory()) gez(p, acc);
    else acc.push(p);
  }
  return acc;
};
for (const [app, dizin] of [["belge", "apps/belge/dist"], ["tablo", "apps/tablo/dist"], ["sunum", "apps/sunum/dist"], ["başlangıç", "apps/desktop/dist-electron/launcher"]]) {
  let dosyalar;
  try {
    dosyalar = gez(join(KOK, dizin));
  } catch {
    continue;
  }
  const html = dosyalar.find((f) => f.endsWith("index.html"));
  const htmlMetni = html ? readFileSync(html, "utf8") : "";
  const ilk = [...htmlMetni.matchAll(/(?:src|href)="\.?\/?([^"]+\.(?:js|css))"/g)].map((m) => m[1]);
  const satirlar = dosyalar
    .filter((f) => [".js", ".css"].includes(extname(f)))
    .map((f) => {
      const b = readFileSync(f);
      const rel = f.slice(join(KOK, dizin).length + 1).replaceAll("\\", "/");
      return { dosya: rel, KB: Math.round(b.length / 1024), gzipKB: Math.round(gzipSync(b, { level: 9 }).length / 1024), brKB: Math.round(brotliCompressSync(b).length / 1024), ilk: ilk.some((i) => rel.endsWith(i.replace(/^\.\//, ""))) };
    })
    .sort((a, b) => b.KB - a.KB);
  const topla = (k, f = () => true) => satirlar.filter(f).reduce((a, r) => a + r[k], 0);
  out[app] = { jsParça: satirlar.filter((r) => r.dosya.endsWith(".js")).length, cssParça: satirlar.filter((r) => r.dosya.endsWith(".css")).length, toplamKB: topla("KB"), toplamGzipKB: topla("gzipKB"), ilkYükKB: topla("KB", (r) => r.ilk), ilkYükGzipKB: topla("gzipKB", (r) => r.ilk), enBüyük: satirlar.slice(0, 8) };
}
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "paket_boyut.json"), JSON.stringify(out, null, 1));
for (const [k, v] of Object.entries(out)) {
  console.log(`\n${k}: ${v.jsParça} JS + ${v.cssParça} CSS parça; toplam ${v.toplamKB} KB (gzip ${v.toplamGzipKB} KB); ilk yük ${v.ilkYükKB} KB (gzip ${v.ilkYükGzipKB} KB)`);
  for (const r of v.enBüyük) console.log(`  ${r.ilk ? "*" : " "} ${r.dosya} ${r.KB} KB / gz ${r.gzipKB} / br ${r.brKB}`);
}
