// Kaydetme dayanıklılığı deneyi: disk dolu / eşzamanlı kaydetme / yazma ortasında süreç ölümü.
//
// NOT: Ölçüm turundaki (v0.1.5) masaüstü ana süreç kaydetme işleyicilerinin MANTIK KOPYASIYLA çalışır
// ("oh:write-file": geçici dosya + rename; "oh:save-file-as": doğrudan writeFile). Kaydetme güvenliği
// düzeltmesinden sonra asıl uygulama davranışı apps/desktop testleriyle doğrulanır; bu betik senaryoların
// ve eski davranışın başvurusu olarak durur (yeni işleyici buraya taşınırsa aynı senaryolar ona uygulanabilir).
// Kullanım: node scripts/derlem/olcum/kaydet_sim.mjs [küçük_tmpfs_dizini]
//   Disk dolu testi için ör. (root): mount -t tmpfs -o size=256k tmpfs /mnt/dolu ; dizin verilmezse atlanır.
// Çıktı: .derlem/rapor/kaydet_sim.json
import fs from "node:fs/promises";
import { existsSync, readdirSync, readFileSync, writeFileSync, mkdirSync, rmSync } from "node:fs";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { CIKTI, RAPOR } from "../ortak.mjs";

const DOLU = process.argv[2] ?? null;
const WORK = join(CIKTI, "kaydet");
rmSync(WORK, { recursive: true, force: true });
mkdirSync(WORK, { recursive: true });
const h = (b) => createHash("sha1").update(b).digest("hex").slice(0, 10);
const rnd = (n, seed) => { const b = Buffer.alloc(n); let s = seed; for (let i = 0; i < n; i++) { s = (s * 1103515245 + 12345) & 0x7fffffff; b[i] = s >> 16; } return b; };

// --- index.ts'deki işleyicilerin birebir mantığı
async function ohWriteFile(filePath, data) {
  try {
    const tmp = `${filePath}.oh-${process.pid}.tmp`;
    await fs.writeFile(tmp, Buffer.from(data));
    await fs.rename(tmp, filePath);
    return true;
  } catch {
    return false;
  }
}
async function ohSaveFileAs(filePath, data) {
  await fs.writeFile(filePath, Buffer.from(data));
  return { path: filePath };
}
const out = {};

// 1) Disk dolu
if (DOLU && existsSync(DOLU)) {
  for (const f of readdirSync(DOLU)) rmSync(`${DOLU}/${f}`, { force: true });
  const orig = rnd(60 * 1024, 1);
  writeFileSync(`${DOLU}/belge.docx`, orig);
  // boş alanı ~40 KB'a indir
  try { writeFileSync(`${DOLU}/dolgu`, Buffer.alloc(150 * 1024)); } catch {}
  const yeni = rnd(120 * 1024, 2);
  const ok = await ohWriteFile(`${DOLU}/belge.docx`, yeni);
  const after = readFileSync(`${DOLU}/belge.docx`);
  out.diskDolu_writeFile = { dönüş: ok, özgünKorundu: h(after) === h(orig), artıkGeçiciDosya: readdirSync(DOLU).filter((f) => f.includes(".tmp")).map((f) => `${f} (${readFileSync(`${DOLU}/${f}`).length} B)`) };
  for (const f of readdirSync(DOLU)) if (f.includes(".tmp")) rmSync(`${DOLU}/${f}`);
  let err = null;
  try { await ohSaveFileAs(`${DOLU}/belge.docx`, yeni); } catch (e) { err = e.code; }
  const after2 = readFileSync(`${DOLU}/belge.docx`);
  out.diskDolu_saveFileAs_üzerineYaz = { hata: err, özgünKorundu: h(after2) === h(orig), dosyaBoyu: after2.length, beklenen: yeni.length };
  for (const f of readdirSync(DOLU)) rmSync(`${DOLU}/${f}`, { force: true });
}

// 2) Eşzamanlı iki kayıt (otomatik kaydet + Ctrl+S) aynı dosyaya: geçici ad yalnız PID içerir → çakışır
{
  const N = 200;
  let bozuk = 0, ikisiDeTrue = 0, falseDönen = 0;
  const A = rnd(3 * 1024 * 1024, 3), B = rnd(1 * 1024 * 1024, 4);
  const ha = h(A), hb = h(B);
  const örnek = [];
  for (let i = 0; i < N; i++) {
    const p = `${WORK}/yarış.xlsx`;
    writeFileSync(p, rnd(1000, 5));
    const [r1, r2] = await Promise.all([ohWriteFile(p, A), ohWriteFile(p, B)]);
    const got = readFileSync(p);
    const hg = h(got);
    if (r1 && r2) ikisiDeTrue++;
    if (!r1 || !r2) falseDönen++;
    if (hg !== ha && hg !== hb) {
      bozuk++;
      if (örnek.length < 3) örnek.push(`boy=${got.length} (A=${A.length}, B=${B.length}) dönüş=${r1}/${r2}`);
    }
    for (const f of readdirSync(WORK)) if (f.includes(".tmp")) rmSync(`${WORK}/${f}`);
  }
  out.eşzamanlıKayıt = { deneme: N, bozukSonuç: bozuk, ikisiDeBaşarılı: ikisiDeTrue, birisiFalse: falseDönen, örnek };
}

// 3) Yazma sırasında süreç ölürse (ör. çökme/elektrik): alt süreçte yaz, 2-40 ms sonra SIGKILL
async function killMid(mode, delays) {
  let korunan = 0, bozuk = 0, yeni = 0;
  const orig = rnd(200 * 1024, 7);
  for (const d of delays) {
    const p = `${WORK}/cokme-${mode}.docx`;
    writeFileSync(p, orig);
    const code = mode === "atomik"
      ? `const fs=require("fs/promises");(async()=>{const d=Buffer.alloc(64*1024*1024,7);const tmp=${JSON.stringify(p)}+".oh-"+process.pid+".tmp";await fs.writeFile(tmp,d);await fs.rename(tmp,${JSON.stringify(p)});})()`
      : `const fs=require("fs/promises");(async()=>{const d=Buffer.alloc(64*1024*1024,7);await fs.writeFile(${JSON.stringify(p)},d);})()`;
    const c = spawn(process.execPath, ["-e", code]);
    const exited = new Promise((r) => c.on("exit", r));
    await new Promise((r) => setTimeout(r, d));
    c.kill("SIGKILL");
    await exited;
    const got = readFileSync(p);
    if (h(got) === h(orig)) korunan++;
    else if (got.length === 64 * 1024 * 1024) yeni++;
    else bozuk++;
  }
  return { deneme: delays.length, özgünKorundu: korunan, tamYeniSürüm: yeni, YARIM_BOZUK: bozuk };
}
const delays = Array.from({ length: 30 }, (_, i) => 20 + i * 5);
out.çökme_writeFile_atomik = await killMid("atomik", delays);
out.çökme_saveFileAs_doğrudan = await killMid("dogrudan", delays);
out.artıkGeçiciDosyalar = readdirSync(WORK).filter((f) => f.includes(".tmp")).length;
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "kaydet_sim.json"), JSON.stringify(out, null, 1));
console.log(JSON.stringify(out, null, 1));
