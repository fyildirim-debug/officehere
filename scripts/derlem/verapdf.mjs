// PDF/A doğrulaması (pnpm derlem:pdfa). Yalnız geliştirme/CI aracıdır: veraPDF (GPLv3+/MPLv2+) ayrı
// süreç olarak çalıştırılır, pakete girmez ve dağıtılmaz.
//
// Derlemdeki (test-derlem/) her belge masaüstü uygulamasında (Electron) açılır, kullanıcının yolundan
// (Dosya › Dışarı Aktar › PDF › seçenek kutusu) PDF/A-2b olarak dışa aktarılır ve veraPDF ile doğrulanır.
// Belge dosyalarında ayrıca: yer imi sayısı = başlık sayısı, belgedeki Türkçe harfler pdftotext ile geri
// okunur. Saydamlıksız şablon (yeni belge: başlık, Türkçe metin, sayfa sonları) PDF/A-1b olarak ve sayfa
// aralığıyla dışa aktarılıp aynı denetimlerden geçer; Resmî Yazışma şablonlarının (8) her biri de PDF/A-1b
// olarak sorusuz dışa aktarılmalıdır. Her 2b çıktısı ayrıca uygulamanın 1b yolundan geçirilir:
// gerçek engeli (saydamlık, sRGB dışı profil, uzun belge) olanlar raporlanır, engelsizler veraPDF 1b'den geçmelidir.
//
// Önce: pnpm web && pnpm --filter @officehere/desktop build:electron   (Linux'ta: xvfb-run -a pnpm derlem:pdfa)
// Seçenekler:
//   --kur              veraPDF'i indirir (sürüm ve SHA-256 sabit) ve .derlem/verapdf altına kurar
//   --yalniz-kur       kurduktan sonra tur çalıştırılmaz (CI önbelleği)
//   --verapdf <yol>    veraPDF komutu (varsayılan: $VERAPDF, .derlem/verapdf/verapdf, PATH'teki verapdf)
//   --dosya <alt dize> yalnız yolu bunu içeren derlem dosyaları
//   --buyuk            büyük derlem dosyalarını (lo-buyuk.*) da dahil eder
// Çıktılar: .derlem/pdfa/*.pdf, .derlem/rapor/pdfa.md ve pdfa.json. Bir doğrulama düşerse çıkış kodu 1.
import { execFileSync, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { argumanlar, CIKTI, derlemDosyalari, KOK, RAPOR, sure } from "./ortak.mjs";

const VERAPDF = {
  surum: "1.30.3",
  url: "https://software.verapdf.org/releases/1.30/verapdf-greenfield-1.30.3-installer.zip",
  sha256: "b512bf24945503a33fc10738ecffc1f2dccaaf8593a448e94ee9ac86a9559a59",
};
const KURULUM = join(CIKTI, "verapdf");
const TR = ["ı", "İ", "ş", "Ş", "ğ", "Ğ", "ç", "Ç", "ö", "Ö", "ü", "Ü"];
const A = argumanlar();

/* ------------------------------------------------------------------ */
/* veraPDF                                                              */
/* ------------------------------------------------------------------ */

async function kur() {
  const komut = join(KURULUM, "verapdf");
  if (existsSync(komut)) return console.log(`veraPDF ${VERAPDF.surum} kurulu: ${komut}`);
  const gecici = mkdtempSync(join(tmpdir(), "verapdf-"));
  const zip = join(gecici, "kurulum.zip");
  console.log(`veraPDF ${VERAPDF.surum} indiriliyor…`);
  const yanit = await fetch(VERAPDF.url);
  if (!yanit.ok) throw new Error(`veraPDF indirilemedi: HTTP ${yanit.status}`);
  const veri = Buffer.from(await yanit.arrayBuffer());
  const ozet = createHash("sha256").update(veri).digest("hex");
  if (ozet !== VERAPDF.sha256) throw new Error(`veraPDF SHA-256 uyuşmuyor: ${ozet}`);
  writeFileSync(zip, veri);
  execFileSync("unzip", ["-q", zip, "-d", gecici]);
  const yanitlar = join(gecici, "auto.xml");
  // IzPack kurulumu: yalnız komut satırı betikleri ve çekirdek (GUI dışı paketler kapalı).
  writeFileSync(
    yanitlar,
    `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<AutomatedInstallation langpack="eng">
  <com.izforge.izpack.panels.htmlhello.HTMLHelloPanel id="welcome"/>
  <com.izforge.izpack.panels.target.TargetPanel id="install_dir"><installpath>${KURULUM}</installpath></com.izforge.izpack.panels.target.TargetPanel>
  <com.izforge.izpack.panels.packs.PacksPanel id="sdk_pack_select">
    <pack index="0" name="veraPDF GUI" selected="true"/>
    <pack index="1" name="veraPDF Mac and *nix Scripts" selected="true"/>
    <pack index="2" name="veraPDF Validation model" selected="false"/>
    <pack index="3" name="veraPDF Documentation" selected="false"/>
    <pack index="4" name="veraPDF Sample Plugins" selected="false"/>
  </com.izforge.izpack.panels.packs.PacksPanel>
  <com.izforge.izpack.panels.install.InstallPanel id="install"/>
  <com.izforge.izpack.panels.finish.FinishPanel id="finish"/>
</AutomatedInstallation>
`,
  );
  execFileSync("java", ["-jar", join(gecici, `verapdf-greenfield-${VERAPDF.surum}`, `verapdf-izpack-installer-${VERAPDF.surum}.jar`), yanitlar], { stdio: "inherit" });
  rmSync(gecici, { recursive: true, force: true });
  if (!existsSync(komut)) throw new Error("veraPDF kurulamadı");
  console.log(`veraPDF kuruldu: ${komut}`);
}

function veraKomutu() {
  for (const aday of [A.verapdf, process.env.VERAPDF, join(KURULUM, "verapdf")]) if (typeof aday === "string" && existsSync(aday)) return aday;
  if (spawnSync("verapdf", ["--version"]).status === 0) return "verapdf";
  return null;
}

/** Dosyaları tek veraPDF çağrısında doğrular: get(yol) → { gecti, kurallar } (adlar NFC ile eşleştirilir). */
function dogrula(komut, flavour, yollar) {
  const out = new Map();
  if (!yollar.length) return { get: () => undefined };
  const r = spawnSync(komut, ["--flavour", flavour, "--format", "json", ...yollar], { maxBuffer: 256 * 1024 * 1024, encoding: "utf8" });
  const json = JSON.parse(r.stdout);
  for (const is of json.report.jobs) {
    const v = Array.isArray(is.validationResult) ? is.validationResult[0] : is.validationResult;
    const kurallar = (v?.details?.ruleSummaries ?? []).map((k) => `${k.clause}-${k.testNumber} (${k.failedChecks}): ${k.description}`);
    out.set(is.itemDetails.name.normalize("NFC"), { gecti: v?.compliant === true, kurallar: v ? kurallar : [is.taskException?.exceptionMessage ?? "doğrulanamadı"] });
  }
  return { get: (yol) => out.get(yol.normalize("NFC")) };
}

const pdfMetni = (yol) => execFileSync("pdftotext", ["-enc", "UTF-8", yol, "-"], { maxBuffer: 64 * 1024 * 1024 }).toString("utf8");
const sayfaMetinleri = (yol) => pdfMetni(yol).split("\f").slice(0, -1);

/* ------------------------------------------------------------------ */
/* PDF yapısı (pdf-lib; masaüstü paketinin bağımlılığı)                 */
/* ------------------------------------------------------------------ */

const pdfLib = createRequire(join(KOK, "apps/desktop/package.json"))("@cantoo/pdf-lib");

async function yapi(yol) {
  const { PDFDocument, PDFName, PDFDict } = pdfLib;
  const doc = await PDFDocument.load(readFileSync(yol), { updateMetadata: false });
  const yerImleri = [];
  const gez = (dugum) => {
    for (let it = dugum instanceof PDFDict ? dugum.lookup(PDFName.of("First")) : undefined; it instanceof PDFDict; it = it.lookup(PDFName.of("Next"))) {
      yerImleri.push(it.lookup(PDFName.of("Title")).decodeText());
      gez(it);
    }
  };
  gez(doc.catalog.lookup(PDFName.of("Outlines")));
  let gomulmemis = 0;
  for (const [, o] of doc.context.enumerateIndirectObjects()) {
    if (!(o instanceof PDFDict) || o.get(PDFName.of("Type"))?.toString() !== "/Font") continue;
    const alt = o.get(PDFName.of("Subtype"))?.toString();
    if (alt === "/Type0" || alt === "/Type3") continue;
    const fd = o.lookup(PDFName.of("FontDescriptor"));
    if (!(fd instanceof PDFDict) || !["FontFile", "FontFile2", "FontFile3"].some((k) => fd.has(PDFName.of(k)))) gomulmemis++;
  }
  return { sayfa: doc.getPageCount(), yerImleri, gomulmemis };
}

/* ------------------------------------------------------------------ */
/* Electron ile dışa aktarma                                            */
/* ------------------------------------------------------------------ */

async function uygulamaAc() {
  const { _electron: electron } = await import("playwright");
  const masaustu = join(KOK, "apps/desktop");
  const app = await electron.launch({
    executablePath: createRequire(join(masaustu, "package.json"))("electron"),
    args: [masaustu, "--no-sandbox"],
    env: { ...process.env, OH_APPS_DIR: join(KOK, "dist-web"), OH_USER_DATA: mkdtempSync(join(tmpdir(), "oh-pdfa-")), ELECTRON_DISABLE_SECURITY_WARNINGS: "1", OH_UPDATE_MODE: "off" },
  });
  const ilk = await app.firstWindow();
  await ilk.waitForSelector(".oh-titlebar");
  // Konum ve ileti kutuları sahte: PDF globalThis.__hedef yoluna yazılır, ileti kutuları kaydedilir.
  await app.evaluate(({ dialog, shell }) => {
    const g = globalThis;
    g.__iletiler = [];
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: g.__hedef });
    dialog.showMessageBox = async (...a) => {
      g.__iletiler.push((a.length > 1 ? a[1] : a[0]).message);
      return { response: 0, checkboxChecked: false };
    };
    shell.openPath = async () => "";
    shell.showItemInFolder = () => undefined;
  });
  return { app, ilk };
}

const HAZIR = { belge: ".bl-body .ProseMirror", tablo: ".oh-ribbon", sunum: ".thumb-pane .thumb-item" };

async function pencereAc(app, ilk, dosya, tur) {
  const [win] = await Promise.all([app.waitForEvent("window", { timeout: 60_000 }), dosya ? app.evaluate(({ app: a }, f) => a.emit("second-instance", {}, [process.execPath, ".", f], process.cwd()), dosya) : ilk.evaluate((t) => window.officehereNative.newWindow(t), tur)]);
  await win.waitForSelector(".oh-ribbon", { timeout: 60_000 });
  // Dosya okununca başlık çubuğunda adı görünür.
  if (dosya) await win.locator(".oh-titlebar").filter({ hasText: basename(dosya).replace(/\.[^.]+$/, "") }).waitFor({ timeout: 60_000 });
  await win.waitForSelector(HAZIR[tur], { timeout: 60_000 });
  // Açılışta çıkan kutular (uyarılar) kapatılır.
  for (let i = 0; i < 3 && (await win.locator("[role=dialog]").count()); i++) await win.keyboard.press("Escape");
  return win;
}

async function disaAktar(app, win, hedef, { standard, aralik }) {
  await app.evaluate((_e, h) => {
    globalThis.__hedef = h;
    globalThis.__iletiler = [];
  }, hedef);
  await win.locator(".oh-ribbon-file").click();
  await win.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
  await win.locator(".oh-bs-action", { hasText: /PDF.*Oluştur/ }).click();
  const kutu = win.getByRole("dialog", { name: "PDF Olarak Dışa Aktar" });
  await kutu.getByRole("radio", { name: standard }).check();
  if (aralik) await kutu.getByRole("textbox", { name: "Sayfa aralığı" }).fill(aralik);
  await kutu.getByRole("combobox", { name: "Dışa aktardıktan sonra" }).selectOption("none");
  await kutu.getByRole("button", { name: "Dışa aktar" }).click();
  const t0 = Date.now();
  while (!existsSync(hedef)) {
    if (Date.now() - t0 > 120_000) throw new Error("PDF 120 sn içinde yazılmadı");
    await new Promise((r) => setTimeout(r, 200));
  }
  return app.evaluate(() => globalThis.__iletiler);
}

/** Kaydedilmemiş değişiklik sorusu beklenmesin: pencere doğrudan kapatılır. */
async function kapat(app, win) {
  const bw = await app.browserWindow(win).catch(() => null);
  await bw?.evaluate((w) => w.destroy()).catch(() => undefined);
}

/** Belge gövdesindeki başlıklar ve metin (yer imi ve harf denetimi için); ekran süslemeleri (bölüm sonu işareti vb.) basılmaz. */
const belgeBilgisi = (win) =>
  win.evaluate(() => {
    const govde = document.querySelector(".bl-body .ProseMirror")?.cloneNode(true);
    govde?.querySelectorAll(".ProseMirror-widget").forEach((e) => e.remove());
    return { basliklar: [...(govde?.querySelectorAll("h1, h2, h3, h4, h5, h6") ?? [])].map((h) => h.textContent.trim()), metin: govde?.textContent ?? "" };
  });

/* ------------------------------------------------------------------ */
/* Tur                                                                  */
/* ------------------------------------------------------------------ */

if (A.kur) await kur();
if (A["yalniz-kur"]) process.exit(0);
const komut = veraKomutu();
if (!komut) {
  console.error("veraPDF bulunamadı: `node scripts/derlem/verapdf.mjs --kur` ile kurun ya da --verapdf <yol> verin.");
  process.exit(2);
}
if (spawnSync("pdftotext", ["-v"]).error) {
  console.error("pdftotext bulunamadı (poppler-utils).");
  process.exit(2);
}
if (!existsSync(join(KOK, "dist-web/belge/index.html")) || !existsSync(join(KOK, "apps/desktop/dist-electron/main/pdfa.js"))) {
  console.error("Önce derleyin: pnpm web && pnpm --filter @officehere/desktop build:electron");
  process.exit(2);
}

const dosyalar = derlemDosyalari({ suzgec: typeof A.dosya === "string" ? A.dosya : undefined }).filter((f) => f.kaynak === "derlem" && (A.buyuk || !/buyuk/.test(f.anahtar)));
const ciktiDizini = join(CIKTI, "pdfa");
rmSync(ciktiDizini, { recursive: true, force: true });
mkdirSync(ciktiDizini, { recursive: true });
mkdirSync(RAPOR, { recursive: true });

const t0 = performance.now();
const { app, ilk } = await uygulamaAc();
const sonuclar = [];
console.log(`${dosyalar.length} derlem dosyası PDF/A-2b olarak dışa aktarılıyor…`);
for (const f of dosyalar) {
  const hedef = join(ciktiDizini, `${f.anahtar.replace(/^test-derlem\//, "").replace(/[/\\]/g, "_")}.pdf`);
  const s = { anahtar: f.anahtar, app: f.app, pdf: hedef, flavour: "2b", hatalar: [] };
  const b = performance.now();
  let win;
  try {
    win = await pencereAc(app, ilk, f.yol, f.app);
    const bilgi = f.app === "belge" ? await belgeBilgisi(win) : null;
    await disaAktar(app, win, hedef, { standard: /^PDF\/A-2b/ });
    const y = await yapi(hedef);
    s.sayfa = y.sayfa;
    if (y.gomulmemis) s.hatalar.push(`${y.gomulmemis} yazı tipi gömülmemiş`);
    if (bilgi) {
      s.basliklar = bilgi.basliklar.length;
      s.yerImleri = y.yerImleri.length;
      if (s.yerImleri !== s.basliklar) {
        const eksik = bilgi.basliklar.filter((h) => !y.yerImleri.includes(h));
        s.hatalar.push(`yer imi ${s.yerImleri} ≠ başlık ${s.basliklar}${eksik.length ? ` (yer imi olmayan: ${eksik.slice(0, 5).join(" | ")})` : ""}`);
      }
      const metin = pdfMetni(hedef);
      s.turkce = TR.filter((c) => bilgi.metin.includes(c));
      const eksik = s.turkce.filter((c) => !metin.includes(c));
      if (eksik.length) s.hatalar.push(`pdftotext'te eksik harf: ${eksik.join(" ")}`);
    }
  } catch (e) {
    s.hatalar.push(`dışa aktarılamadı: ${e.message.split("\n")[0]}`);
    s.pdf = null;
  }
  if (win) await kapat(app, win);
  s.sureMs = performance.now() - b;
  sonuclar.push(s);
  console.log(`  ${s.hatalar.length ? "✗" : "·"} ${f.anahtar} ${s.sayfa ?? "-"} sayfa ${sure(s.sureMs)}${s.hatalar.length ? ` — ${s.hatalar.join("; ")}` : ""}`);
}

// Saydamlıksız şablon: başlıklar, Türkçe metin, sayfa sonları → PDF/A-1b ve sayfa aralığı "2-3".
{
  const s = { anahtar: "şablon (yeni belge)", app: "belge", pdf: join(ciktiDizini, "sablon-1b.pdf"), flavour: "1b", hatalar: [] };
  const b = performance.now();
  try {
    const win = await pencereAc(app, ilk, null, "belge");
    await win.locator(".bl-body .ProseMirror").first().click();
    const sayfa = async (n, baslik, metin, son) => {
      await win.keyboard.press("Control+Alt+1");
      await win.keyboard.type(baslik);
      await win.keyboard.press("Enter");
      await win.keyboard.type(metin);
      if (!son) await win.keyboard.press("Control+Enter");
    };
    await sayfa(1, "Giriş", "Birinci sayfa metni.", false);
    await sayfa(2, "Çalışma Takvimi", "Ilık ağaç gölgesinde şiir okuyan öğrenciler: ığüşöç İĞÜŞÖÇ.", false);
    await sayfa(3, "Sonuç ve Öneriler", "Üçüncü sayfa; çözüm önerileri.", true);
    await win.locator(".bl-page").nth(2).waitFor();
    const iletiler = await disaAktar(app, win, s.pdf, { standard: /^PDF\/A-1b/, aralik: "2-3" });
    if (iletiler.length) s.hatalar.push(`beklenmeyen ileti: ${iletiler.join("; ")}`);
    const y = await yapi(s.pdf);
    const metinler = sayfaMetinleri(s.pdf);
    s.sayfa = y.sayfa;
    s.basliklar = 2;
    s.yerImleri = y.yerImleri.length;
    if (y.sayfa !== 2) s.hatalar.push(`sayfa aralığı 2-3: ${y.sayfa} sayfa`);
    if (!metinler[0]?.includes("Çalışma Takvimi") || !metinler[1]?.includes("Sonuç ve Öneriler") || metinler.join("").includes("Birinci")) s.hatalar.push("sayfa aralığı yanlış sayfaları içeriyor");
    if (JSON.stringify(y.yerImleri) !== JSON.stringify(["Çalışma Takvimi", "Sonuç ve Öneriler"])) s.hatalar.push(`yer imleri: ${y.yerImleri.join(" | ")}`);
    s.turkce = TR;
    const eksik = TR.filter((c) => !metinler.join("").includes(c));
    if (eksik.length) s.hatalar.push(`pdftotext'te eksik harf: ${eksik.join(" ")}`);
    await kapat(app, win);
  } catch (e) {
    s.hatalar.push(`dışa aktarılamadı: ${e.message.split("\n")[0]}`);
    s.pdf = null;
  }
  s.sureMs = performance.now() - b;
  sonuclar.push(s);
  console.log(`  ${s.hatalar.length ? "✗" : "·"} ${s.anahtar} ${s.sayfa ?? "-"} sayfa${s.hatalar.length ? ` — ${s.hatalar.join("; ")}` : ""}`);
}
// Resmî Yazışma şablonları (Dosya › Yeni › Resmî yazışma): saydamlık üretmezler, PDF/A-1b sorusuz geçmelidir.
const SABLONLAR = ["Üst yazı", "Dağıtımlı yazı", "Olur yazısı", "Tekit yazısı", "Ek ve dağıtım listesi", "Dilekçe", "Tutanak ve rapor", "Zarf"];
for (const [i, ad] of SABLONLAR.entries()) {
  const s = { anahtar: `resmî şablon: ${ad}`, app: "belge", pdf: join(ciktiDizini, `resmi-${i + 1}-1b.pdf`), flavour: "1b", hatalar: [] };
  const b = performance.now();
  let win;
  try {
    win = await pencereAc(app, ilk, null, "belge");
    await win.locator(".oh-ribbon-file").click();
    await win.locator(".oh-bs-item", { hasText: "Yeni" }).first().click();
    await win.locator(".oh-template", { hasText: "Resmî yazışma" }).click();
    const kutu = win.getByRole("dialog", { name: "Resmî Yazışma Şablonları" });
    await kutu.getByRole("button", { name: ad, exact: true }).click();
    const tarih = kutu.getByPlaceholder("GG.AA.YYYY");
    if (await tarih.count()) await tarih.first().fill("15.10.2026");
    await kutu.getByRole("button", { name: "Oluştur", exact: true }).click();
    await kutu.waitFor({ state: "hidden" });
    await win.locator(".bl-body .ProseMirror p").first().waitFor();
    const iletiler = await disaAktar(app, win, s.pdf, { standard: /^PDF\/A-1b/ });
    if (iletiler.length) s.hatalar.push(`beklenmeyen ileti: ${iletiler.join("; ")}`);
    const y = await yapi(s.pdf);
    s.sayfa = y.sayfa;
    if (y.gomulmemis) s.hatalar.push(`${y.gomulmemis} yazı tipi gömülmemiş`);
    if (!/\S/.test(pdfMetni(s.pdf))) s.hatalar.push("pdftotext metin bulamadı");
  } catch (e) {
    s.hatalar.push(`dışa aktarılamadı: ${e.message.split("\n")[0]}`);
    s.pdf = null;
  }
  if (win) await kapat(app, win);
  s.sureMs = performance.now() - b;
  sonuclar.push(s);
  console.log(`  ${s.hatalar.length ? "✗" : "·"} ${s.anahtar} ${s.sayfa ?? "-"} sayfa${s.hatalar.length ? ` — ${s.hatalar.join("; ")}` : ""}`);
}
await app.close().catch(() => undefined);

// 1b: her 2b çıktısı uygulamanın kendi 1b yolundan (prepareForPdfA1 + applyPdfStandard) geçirilir; gerçek
// engeli (saydamlık, sRGB dışı profil, uzun belge) olanlar raporlanır, engelsizler veraPDF 1b'den geçmelidir.
const pdfa = createRequire(import.meta.url)(join(KOK, "apps/desktop/dist-electron/main/pdfa.js"));
const icc = pdfa.loadSrgbProfile();
for (const s of sonuclar.filter((x) => x.pdf && x.flavour === "2b")) {
  const doc = await pdfa.loadPdf(readFileSync(s.pdf));
  const engel = pdfa.prepareForPdfA1(doc, icc);
  if (engel.length) {
    s.birB = `engel: ${engel.join(" ").replace(/•\s*/g, "").replace(/\s+/g, " ").slice(0, 120)}`;
    continue;
  }
  pdfa.applyPdfStandard(doc, { standard: "pdfa-1b", meta: { title: basename(s.anahtar), producer: "OfficeHere derlem" }, icc });
  s.pdf1b = s.pdf.replace(/\.pdf$/, ".1b.pdf");
  writeFileSync(s.pdf1b, await pdfa.savePdf(doc));
}

console.log("veraPDF doğruluyor…");
for (const flavour of ["2b", "1b"]) {
  const grup = sonuclar.filter((s) => s.pdf && s.flavour === flavour);
  const v = dogrula(komut, flavour, grup.map((s) => s.pdf));
  for (const s of grup) {
    const r = v.get(s.pdf);
    s.vera = r?.gecti ? "GEÇTİ" : "KALDI";
    if (!r?.gecti) s.hatalar.push(`veraPDF ${flavour}: ${(r?.kurallar ?? ["sonuç yok"]).join("; ")}`);
  }
}
{
  const grup = sonuclar.filter((s) => s.pdf1b);
  const v = dogrula(komut, "1b", grup.map((s) => s.pdf1b));
  for (const s of grup) {
    const r = v.get(s.pdf1b);
    s.birB = r?.gecti ? "GEÇTİ" : "KALDI";
    if (!r?.gecti) s.hatalar.push(`veraPDF 1b (dönüştürülen): ${(r?.kurallar ?? ["sonuç yok"]).join("; ")}`);
  }
}

const kalan = sonuclar.filter((s) => s.hatalar.length);
const birB = sonuclar.filter((s) => s.pdf1b);
const satir = (s) =>
  `| ${s.anahtar} | ${s.app} | ${s.flavour} | ${s.sayfa ?? "-"} | ${s.vera ?? "-"} | ${s.birB ?? "-"} | ${s.basliklar === undefined ? "-" : `${s.yerImleri}/${s.basliklar}`} | ${s.turkce ? s.turkce.join("") || "-" : "-"} | ${sure(s.sureMs)} |`;
const md = [
  "# PDF/A doğrulaması (veraPDF)",
  "",
  `${sonuclar.length} PDF · veraPDF ${VERAPDF.surum} · geçen ${sonuclar.length - kalan.length} · kalan ${kalan.length} · ${sure(performance.now() - t0)}`,
  `PDF/A-1b'ye dönüştürülebilen: ${birB.length}/${sonuclar.filter((s) => s.flavour === "2b").length} (veraPDF 1b geçen ${birB.filter((s) => s.birB === "GEÇTİ").length}); kalanların engeli tabloda.`,
  "",
  "| Dosya | Uygulama | PDF/A | Sayfa | veraPDF | 1b (dönüştürülen) | Yer imi/başlık | Türkçe harfler | Süre |",
  "| --- | --- | --- | --- | --- | --- | --- | --- | --- |",
  ...sonuclar.map(satir),
  "",
  ...(kalan.length ? ["## Kalanlar", "", ...kalan.map((s) => `- **${s.anahtar}**: ${s.hatalar.join("; ")}`), ""] : []),
].join("\n");
writeFileSync(join(RAPOR, "pdfa.md"), md);
writeFileSync(join(RAPOR, "pdfa.json"), JSON.stringify(sonuclar, null, 2));
console.log(`\n${sonuclar.length - kalan.length}/${sonuclar.length} PDF geçti; 1b'ye dönüştürülebilen ${birB.length}, veraPDF 1b geçen ${birB.filter((s) => s.birB === "GEÇTİ").length}. Rapor: ${join(RAPOR, "pdfa.md")}`);
for (const s of kalan) console.log(`  ✗ ${basename(s.anahtar)}: ${s.hatalar.join("; ")}`);
process.exit(kalan.length ? 1 : 0);
