// Bulanık test (pnpm bulanik): 756 bozuk örneği uygulama okuyucularıyla açar.
//
// Her örnek ayrı işçide, süre sınırıyla (5 sn × makine hızı oranı) açılır; açılırsa kaydedilebiliyor mu
// denenir. Sınıflar:
//   açıldı · uyarılı (açıldı, kullanıcıya uyarı verildi) · kaydedilemiyor (açıldı ama yazıcı çöktü)
//   anlaşılır hata (Türkçe, kullanıcıya gösterilebilir ileti) · ham hata (teknik ileti: "Corrupted zip…")
//   zaman aşımı (5 sn'yi aştı ama uzun denemede bitti) · takılma (uzun denemede de bitmedi)
//   yığın taşması (RangeError: Maximum call stack) · bellek taşması · işçi çöktü
// Zaman aşımına düşenler uzun sınırla (60 sn × oran) yeniden denenir: takılma mı, yalnız yavaş mı?
// Çıktı: .derlem/rapor/bulanik.json + bulanik.md; temel çizgiyle karşılaştırma (scripts/derlem/temel.mjs).
//
// Kullanım: pnpm bulanik [--uygulama belge|tablo|sunum] [--temelsiz] [--kip siki|baslangic] [--paketleme-yok]
//   Ortam: BULANIK_SINIR_MS (5000), BULANIK_UZUN_MS (60000)
import { execSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { availableParallelism } from "node:os";
import { join } from "node:path";
import { Havuz } from "../derlem/havuz.mjs";
import { hizOlc, hizOrani } from "../derlem/hiz.mjs";
import { argumanlar, KOK, PAKET, RAPOR, sure } from "../derlem/ortak.mjs";
import { paketle } from "../derlem/paketle.mjs";
import { bulanikKarsilastir, bulanikOzeti, bulguRaporu, kipBelirle, temelOku } from "../derlem/temel.mjs";
import { BULANIK, uret } from "./uret.mjs";

const A = argumanlar();
const t0 = performance.now();
const TEMEL_SINIR = Number(process.env.BULANIK_SINIR_MS ?? 5000);
const TEMEL_UZUN = Number(process.env.BULANIK_UZUN_MS ?? 60000);
const temel = temelOku();

if (!A["paketleme-yok"]) await paketle();
const liste = await uret();
const ornekler = liste.ornekler.filter((x) => !A.uygulama || x.app === A.uygulama);

// Makine hızı: temel çizgiyi üreten makineden yavaşsa sınırlar orantılı uzar (CI'da yanlış alarm olmasın).
const hizMs = hizOlc();
const oran = hizOrani(hizMs, temel.bulanik?.hizMs);
const SINIR = Math.round(TEMEL_SINIR * oran);
const UZUN = Math.round(TEMEL_UZUN * oran);
console.log(`${ornekler.length} örnek; makine hızı ${hizMs} ms (temel ${temel.bulanik?.hizMs ?? "-"} ms, oran ${oran.toFixed(2)}) → sınır ${SINIR} ms, uzun deneme ${UZUN} ms`);

const TR = /[çğıöşüÇĞİÖŞÜ]|dosya|bozuk|açıl|okun|biçim|parola|desteklen/i;
const yiginMi = (h) => h && h.ad === "RangeError" && /call stack/i.test(h.ileti);
function siniflandir(r) {
  if (r.durum === "zaman-asimi") return "zaman aşımı";
  if (r.durum === "bellek") return "bellek taşması";
  if (r.durum !== "tamam") return "işçi çöktü";
  const o = r.cikti;
  if (o.cokme) return yiginMi(o.cokme) ? "yığın taşması" : "ham hata";
  if (o.acildi) {
    if (o.yazmaHatasi) return yiginMi(o.yazmaHatasi) ? "yığın taşması" : "kaydedilemiyor";
    return o.uyari ? "uyarılı" : "açıldı";
  }
  if (yiginMi(o.hata)) return "yığın taşması";
  return TR.test(o.hata.ileti) ? "anlaşılır hata" : "ham hata";
}
const iletiOf = (r) => r.cikti?.yazmaHatasi?.ileti ?? r.cikti?.hata?.ileti ?? r.cikti?.cokme?.ileti ?? r.hata?.ileti ?? "";

const boyut = Math.max(1, Math.min(3, availableParallelism() - 1));
const havuz = new Havuz({ boyut, bellekMb: 2048, veri: { paket: PAKET } });
const sonuclar = [];
await Promise.all(
  ornekler.map(async (s) => {
    const r = await havuz.calistir({ tur: "bulanik", app: s.app, yol: join(BULANIK, s.yol), ad: s.ad }, SINIR);
    const sinif = siniflandir(r);
    const o = { app: s.app, ad: s.ad, yol: s.yol, tur: s.tur, tohum: s.tohum, sinif, sureMs: r.sureMs };
    if (r.cikti?.acildi) Object.assign(o, { bicim: r.cikti.bicim, uyari: r.cikti.uyari, acmaMs: r.cikti.ms, ozet: r.cikti.ozet });
    const ileti = iletiOf(r);
    if (ileti) o.ileti = ileti.slice(0, 300);
    if (r.cikti?.hata?.ad) o.hataAdi = r.cikti.hata.ad;
    sonuclar.push(o);
    if (!/^(açıldı|uyarılı|anlaşılır hata)$/.test(sinif)) process.stdout.write(`${sinif.padEnd(15)} ${s.app} ${s.ad} ${r.sureMs} ms ${ileti.slice(0, 100)}\n`);
  }),
);
await havuz.kapat();

// Zaman aşımına düşenleri uzun sınırla yeniden dene: takılma mı, yalnız yavaş mı?
const yavaslar = sonuclar.filter((r) => r.sinif === "zaman aşımı");
if (yavaslar.length) {
  console.log(`\n${yavaslar.length} zaman aşımı ${UZUN} ms sınırla yeniden deneniyor…`);
  const h2 = new Havuz({ boyut, bellekMb: 4096, veri: { paket: PAKET } });
  await Promise.all(
    yavaslar.map(async (o) => {
      const r = await h2.calistir({ tur: "bulanik", app: o.app, yol: join(BULANIK, o.yol), ad: o.ad }, UZUN);
      const s2 = siniflandir(r);
      o.uzun = { sinif: s2, sureMs: r.sureMs, ileti: iletiOf(r).slice(0, 200) || undefined };
      if (s2 === "zaman aşımı") o.sinif = "takılma";
      else if (s2 === "yığın taşması" || s2 === "bellek taşması" || s2 === "işçi çöktü") o.sinif = s2;
      console.log(`  ${o.sinif.padEnd(12)} ${o.app} ${o.ad} → ${s2} ${r.sureMs} ms`);
    }),
  );
  await h2.kapat();
}
sonuclar.sort((a, b) => a.yol.localeCompare(b.yol));
const toplamMs = Math.round(performance.now() - t0);
let commit = "";
try {
  commit = execSync("git rev-parse --short HEAD", { cwd: KOK, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
} catch {
  /* git yok */
}
const rapor = { tarih: new Date().toISOString(), commit, sureMs: toplamMs, kapsamli: !A.uygulama, hizMs, hizOrani: oran, temelSinirMs: TEMEL_SINIR, sinirMs: SINIR, uzunSinirMs: UZUN, ornekler: sonuclar };
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "bulanik.json"), JSON.stringify(rapor, null, 1));

/* ------------------------------------------------------------ özet (Markdown) */
const SINIFLAR = ["açıldı", "uyarılı", "kaydedilemiyor", "anlaşılır hata", "ham hata", "zaman aşımı", "takılma", "yığın taşması", "bellek taşması", "işçi çöktü"];
const apps = [...new Set(sonuclar.map((r) => r.app))];
const L = [];
L.push(`# Bulanık test`, "", `${rapor.tarih.slice(0, 16).replace("T", " ")} · commit ${commit || "-"} · ${sonuclar.length} örnek · ${sure(toplamMs)} · sınır ${SINIR} ms (uzun deneme ${UZUN} ms, hız oranı ${oran.toFixed(2)})`, "");
L.push(`| sınıf | ${apps.join(" | ")} | toplam |`, `|---|${apps.map(() => "---").join("|")}|---|`);
for (const c of SINIFLAR) {
  const n = sonuclar.filter((r) => r.sinif === c).length;
  if (n || ["açıldı", "anlaşılır hata", "ham hata", "zaman aşımı", "yığın taşması"].includes(c)) L.push(`| ${c} | ${apps.map((a) => sonuclar.filter((r) => r.app === a && r.sinif === c).length).join(" | ")} | ${n} |`);
}
L.push("", "## Ham (teknik) hata iletileri — en sık", "");
const iletiler = {};
for (const r of sonuclar.filter((r) => r.sinif === "ham hata" || r.sinif === "kaydedilemiyor")) {
  const k = `${r.app}: ${`${r.hataAdi ?? ""}: ${r.ileti ?? ""}`.replace(/\d+/g, "N").replace(/\s+/g, " ").slice(0, 110)}`;
  iletiler[k] = (iletiler[k] ?? 0) + 1;
}
for (const [k, n] of Object.entries(iletiler).sort((a, b) => b[1] - a[1]).slice(0, 25)) L.push(`- ${n}× ${k}`);
L.push("", "## Mutasyon türüne göre", "");
const turler = {};
for (const r of sonuclar) {
  const k = r.tur.replace(/-\d+$/, "").replace(/:.*/, "");
  turler[k] ??= {};
  turler[k][r.sinif] = (turler[k][r.sinif] ?? 0) + 1;
}
for (const [k, v] of Object.entries(turler)) L.push(`- ${k}: ${Object.entries(v).map(([c, n]) => `${c}=${n}`).join(", ")}`);
L.push("", "## Ciddi örnekler (takılma / zaman aşımı / yığın / bellek / çökme / kaydedilemeyen)", "");
for (const r of sonuclar.filter((r) => !/^(açıldı|uyarılı|anlaşılır hata|ham hata)$/.test(r.sinif)))
  L.push(`- **${r.sinif}** ${r.app} \`${r.yol}\` (${r.tur}, tohum ${r.tohum}, ${r.sureMs} ms${r.uzun ? `; uzun deneme: ${r.uzun.sinif} ${r.uzun.sureMs} ms` : ""}) ${(r.ileti ?? "").replace(/\s+/g, " ").slice(0, 160)}`);
L.push("", "## Açılış süresi (başarılı açılışlar) — en yavaş 10", "");
for (const r of sonuclar.filter((r) => r.acmaMs != null).sort((a, b) => b.acmaMs - a.acmaMs).slice(0, 10)) L.push(`- ${r.acmaMs} ms ${r.app} ${r.ad} (${r.tur}) ${r.ozet ?? ""}`);
let md = L.join("\n") + "\n";

let cikisKodu = 0;
if (!A.temelsiz) {
  const kip = kipBelirle(A, temel);
  const simdi = bulanikOzeti(rapor);
  const temelB = A.uygulama ? { ornekler: Object.fromEntries(Object.entries(temel.bulanik?.ornekler ?? {}).filter(([k]) => k.startsWith(A.uygulama + "/"))) } : temel.bulanik;
  const { md: kmd, hataVar } = bulguRaporu("Bulanık test", bulanikKarsilastir(simdi, temelB, kip));
  md += "\n" + kmd + "\n";
  if (hataVar) cikisKodu = 1;
}
writeFileSync(join(RAPOR, "bulanik.md"), md);
const say = (c) => sonuclar.filter((r) => r.sinif === c).length;
console.log(
  `\nBulanık test: ${sonuclar.length} örnek — açıldı ${say("açıldı")}, uyarılı ${say("uyarılı")}, kaydedilemiyor ${say("kaydedilemiyor")}, anlaşılır hata ${say("anlaşılır hata")}, ham hata ${say("ham hata")}, zaman aşımı ${say("zaman aşımı")}, takılma ${say("takılma")}, yığın taşması ${say("yığın taşması")}, bellek ${say("bellek taşması")}, işçi çöktü ${say("işçi çöktü")} (${sure(toplamMs)})`,
);
console.log(`Rapor: ${join(RAPOR, "bulanik.md")}`);
process.exitCode = cikisKodu;
