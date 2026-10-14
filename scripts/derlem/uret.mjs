// Derlemi üretir: LibreOffice UNO (uret_lo.py) ve python-docx/openpyxl/python-pptx (uret_py.py) ile
// belgeler, ardından LibreOffice'in formül sonuçları (gercek_lo.py).
//
// Üretilenler <CIKTI>/uretilen/ altına yazılır (depoya girmez). Depodaki küçük alt küme (test-derlem/)
// --test-derlem ile yenilenir: SECIM listesindeki dosyalar test-derlem/ altına kopyalanır.
// LibreOffice yoksa yalnız python üretimi yapılır; python kütüphaneleri yoksa o adım atlanır.
//
// Kullanım:
//   pnpm derlem:uret                 # büyük derlem dosyaları + formül sonuçları
//   pnpm derlem:uret --test-derlem   # ayrıca depodaki alt kümeyi yeniden üret (gözden geçirip commit edin)
import { spawn, spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:net";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { argumanlar, CIKTI, DERLEM, derlemDosyalari, GORSEL, KOK, sure, URETILEN } from "./ortak.mjs";

const A = argumanlar();
const BURADA = join(KOK, "scripts/derlem");
/** Depoya konan alt küme (≈2,3 MB): her biçimden en değerli örnekler; büyükler betikle üretilir. */
export const SECIM = {
  lo: [
    "lo-rapor.docx", "lo-rapor.odt", "lo-rapor.rtf", "lo-rapor.doc", "lo-rapor.html",
    "lo-karma.docx", "lo-karma.odt", "lo-karma.rtf", "lo-karma.doc",
    "lo-uzun.docx",
    "lo-butce.xlsx", "lo-butce.ods", "lo-butce.xls", "lo-butce.csv",
    "lo-buyuk.xlsx", "lo-buyuk.ods",
    "lo-sunu.pptx", "lo-sunu.odp", "lo-sunu.ppt",
    "lo-sunu24.pptx", "lo-sunu24.odp",
    "envanter.json",
  ],
  py: ["py-belge.docx", "py-tablo.xlsx", "py-sunu.pptx"],
};

const t0 = performance.now();
const komutVar = (k, args = ["--version"]) => spawnSync(k, args, { stdio: "ignore" }).status === 0;
const python = (args, ek = {}) => spawnSync("python3", args, { stdio: ["ignore", "inherit", "inherit"], cwd: KOK, ...ek });
const ortam = (port) => ({ ...process.env, LO_PORT: String(port), DERLEM_GORSEL: GORSEL, LC_ALL: "C.UTF-8", LANG: "C.UTF-8" });

async function bosPort() {
  return new Promise((res) => {
    const s = createServer();
    s.listen(0, "127.0.0.1", () => {
      const p = s.address().port;
      s.close(() => res(p));
    });
  });
}

/** LibreOffice'i UNO dinleyicisi olarak başlatır; bağlantı kurulabilene dek bekler. */
async function loBaslat() {
  const port = Number(process.env.LO_PORT ?? 0) || (await bosPort());
  const profil = join(CIKTI, "lo-profil-uret");
  mkdirSync(profil, { recursive: true });
  const lo = spawn("soffice", ["--headless", "--invisible", "--nologo", "--norestore", "--nodefault", `--accept=socket,host=127.0.0.1,port=${port};urp;`, `-env:UserInstallation=${pathToFileURL(profil).href}`], { stdio: "ignore", detached: true });
  const dene = `import uno,sys\nl=uno.getComponentContext()\nl.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",l).resolve("uno:socket,host=127.0.0.1,port=${port};urp;StarOffice.ComponentContext")`;
  for (let i = 0; i < 90; i++) {
    if (spawnSync("python3", ["-c", dene], { stdio: "ignore" }).status === 0) return { lo, port };
    await new Promise((r) => setTimeout(r, 1000));
  }
  loDurdur({ lo, port });
  throw new Error("LibreOffice UNO dinleyicisi 90 sn içinde hazır olmadı.");
}
function loDurdur({ lo, port }) {
  spawnSync("python3", ["-c", `import uno\nl=uno.getComponentContext()\nc=l.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",l).resolve("uno:socket,host=127.0.0.1,port=${port};urp;StarOffice.ComponentContext")\ntry:\n  c.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",c).terminate()\nexcept Exception:\n  pass`], { stdio: "ignore", timeout: 20000 });
  try {
    process.kill(-lo.pid, "SIGKILL");
  } catch {
    /* zaten kapandı */
  }
}

// 1) Görseller (belgelere gömülür)
if (!existsSync(join(GORSEL, "foto.jpg"))) {
  if (python([join(BURADA, "uret_gorsel.py"), GORSEL]).status !== 0) throw new Error("Görseller üretilemedi (python3 + Pillow gerekli).");
}

const loVar = komutVar("soffice") && spawnSync("python3", ["-c", "import uno"], { stdio: "ignore" }).status === 0;
const pyVar = spawnSync("python3", ["-c", "import docx, openpyxl, pptx"], { stdio: "ignore" }).status === 0;
let lo = null;
if (loVar) {
  lo = await loBaslat();
  console.log(`LibreOffice dinleyicisi hazır (port ${lo.port})`);
} else console.log("LibreOffice (soffice + python3-uno) bulunamadı: LibreOffice derlemi ve formül sonuçları atlanıyor.");

try {
  // 2) LibreOffice derlemi
  if (lo) {
    rmSync(join(URETILEN, "lo"), { recursive: true, force: true });
    const r = python([join(BURADA, "uret_lo.py"), join(URETILEN, "lo")], { env: ortam(lo.port) });
    if (r.status !== 0) throw new Error("uret_lo.py başarısız");
    // Formül sonuçları temiz bir LibreOffice örneğiyle alınır (üretim sırasında çökmüş olabilir).
    loDurdur(lo);
    lo = await loBaslat();
  }
  // 3) python derlemi
  if (pyVar) {
    rmSync(join(URETILEN, "py"), { recursive: true, force: true });
    const r = python([join(BURADA, "uret_py.py"), join(URETILEN, "py")], { env: ortam(0) });
    if (r.status !== 0) throw new Error("uret_py.py başarısız");
  } else console.log("python-docx / openpyxl / python-pptx yok: python derlemi atlanıyor (depodaki test-derlem/py kullanılır).");

  // 4) Depodaki alt kümeyi yenile
  if (A["test-derlem"]) {
    for (const [alt, adlar] of Object.entries(SECIM)) {
      const kaynak = join(URETILEN, alt);
      if (!existsSync(kaynak)) {
        console.log(`uyarı: ${kaynak} yok; test-derlem/${alt} yenilenmedi`);
        continue;
      }
      const hedef = join(DERLEM, alt);
      mkdirSync(hedef, { recursive: true });
      for (const f of readdirSync(hedef)) if (!adlar.includes(f)) rmSync(join(hedef, f));
      for (const f of adlar) if (existsSync(join(kaynak, f))) copyFileSync(join(kaynak, f), join(hedef, f));
      console.log(`test-derlem/${alt}: ${adlar.length} dosya yenilendi`);
    }
  }

  // 5) LibreOffice formül sonuçları (tablo dosyaları)
  if (lo) {
    const tablolar = derlemDosyalari({ uygulama: "tablo" });
    const r = spawnSync("python3", [join(BURADA, "gercek_lo.py"), ...tablolar.map((d) => `${d.anahtar}=${d.yol}`)], { cwd: KOK, env: ortam(lo.port), maxBuffer: 1 << 30, encoding: "utf8", stdio: ["ignore", "pipe", "inherit"] });
    if (r.status !== 0) throw new Error("gercek_lo.py başarısız");
    const tum = JSON.parse(r.stdout);
    const depo = {};
    const uretilen = {};
    for (const [k, v] of Object.entries(tum)) (k.startsWith("uretilen/") ? uretilen : depo)[k] = v;
    mkdirSync(URETILEN, { recursive: true });
    writeFileSync(join(URETILEN, "gercek_lo.json"), JSON.stringify({ ...depo, ...uretilen }));
    // Depodaki dosyaların sonuçları depoya da yazılır: LibreOffice'siz makinede tur yine karşılaştırabilsin.
    if (A["test-derlem"]) writeFileSync(join(DERLEM, "gercek_lo.json"), JSON.stringify(Object.fromEntries(Object.entries(depo).sort(([a], [b]) => a.localeCompare(b)))) + "\n");
    console.log(`formül sonuçları: ${Object.keys(tum).length} kitap`);
  }
} finally {
  if (lo) loDurdur(lo);
}
console.log(`derlem üretildi (${sure(performance.now() - t0)}): ${URETILEN}`);
