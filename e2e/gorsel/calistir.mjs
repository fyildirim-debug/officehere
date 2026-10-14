// Görsel regresyon çalıştırıcısı (pnpm gorsel / pnpm gorsel:guncelle).
//
// Temel görüntüler (e2e/gorsel/__ekran__/linux/) yalnız CI ile birebir aynı ortamda üretilir ve karşılaştırılır:
// Docker imajı mcr.microsoft.com/playwright:v1.56.1-noble + sürümü sabit Carlito/Caladea/Liberation/DejaVu.
// Bu yüzden betik varsayılan olarak her şeyi o imajda çalıştırır (depo /src'ye bağlanır, /work'e kopyalanır,
// bağımlılıklar kurulur, uygulamalar derlenir, sonuç ve güncellenen temeller geri kopyalanır).
//
// Kullanım:
//   pnpm gorsel                      karşılaştır (fark varsa çıkış kodu 1; farklar test-results/gorsel/)
//   pnpm gorsel:guncelle             yalnız eşiği aşan ya da eksik temelleri yeniden yaz
//   pnpm gorsel:guncelle -- --hepsi  temel dizinini silip hepsini yeniden üret (artık kullanılmayanlar da gider)
//   Diğer argümanlar Playwright'a geçer (ör. pnpm gorsel -- -g Sunum).
// Ortam:
//   OH_GORSEL_KONTEYNER=1  zaten imajın içindeyiz (CI işi ya da bu betiğin Docker adımı): doğrudan çalıştır.
//   OH_GORSEL_YEREL=1      Docker'sız, bu makinenin Chromium'uyla yalnız belirlenimcilik denemesi; temeller
//                          OH_GORSEL_EKRAN'a (varsayılan: geçici dizin) yazılır, depodakilerle karşılaştırılmaz.
//   OH_GORSEL_KAYDIR=1     kasıtlı 1 px kaydırma: içeriği olan her görüntü kırmızıya dönmeli (öz denetim).
//   OH_E2E_PORT            önizleme sunucularının taban portu (varsayılan 4173).
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { appendFileSync, existsSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const KOK = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const IMAJ = "mcr.microsoft.com/playwright:v1.56.1-noble";
// .github/workflows/ci.yml "gorsel" işindeki yazı tipi satırıyla aynı tutun (noble, Release cebi).
const YAZI_TIPLERI = "fonts-crosextra-carlito=20230309-2 fonts-crosextra-caladea=20200211-2 fonts-liberation=1:2.1.5-3 fonts-dejavu-core=2.37-8";
const UYGULAMALAR = ["belge", "tablo", "sunum"];

const argv = process.argv.slice(2);
const guncelle = argv.includes("--guncelle");
const hepsi = argv.includes("--hepsi");
const pwArg = argv.filter((a) => a !== "--guncelle" && a !== "--hepsi" && a !== "--");

function dur(ileti, kod = 1) {
  console.error(ileti);
  process.exit(kod);
}

function ozet(dizin) {
  const m = new Map();
  if (!existsSync(dizin)) return m;
  for (const f of readdirSync(dizin)) if (f.endsWith(".png")) m.set(f, createHash("sha256").update(readFileSync(join(dizin, f))).digest("hex"));
  return m;
}

function dogrudan() {
  const yerel = process.env.OH_GORSEL_YEREL === "1";
  const ekran = process.env.OH_GORSEL_EKRAN ?? (yerel ? join(tmpdir(), "officehere-gorsel-yerel") : join(KOK, "e2e/gorsel/__ekran__"));
  const dizin = join(resolve(KOK, ekran), process.platform);
  const eksik = UYGULAMALAR.filter((a) => !existsSync(join(KOK, "apps", a, "dist/index.html")));
  if (eksik.length) dur(`Önce uygulamaları derleyin (${eksik.join(", ")} yok): pnpm -r --filter ./apps/belge --filter ./apps/tablo --filter ./apps/sunum run build`);
  if (guncelle && !process.env.OH_GORSEL_EKRAN && !yerel) {
    // Depodaki temeller yalnız resmi imajda yazılır; başka makinenin yazı tipleriyle üretilen temel CI'da kırmızıdır.
    const carlito = spawnSync("fc-list", [":family=Carlito"], { encoding: "utf8" }).stdout?.trim();
    if (!existsSync("/ms-playwright") || !carlito) dur(`Temel görüntüler yalnız ${IMAJ} imajında (Carlito vb. kurulu) üretilir. Docker ile: pnpm gorsel:guncelle (OH_GORSEL_KONTEYNER olmadan).`);
  }
  const once = ozet(dizin);
  if (!guncelle && once.size === 0) {
    console.log(`Temel görüntü yok (${dizin}); karşılaştırma atlandı. Üretmek için: pnpm gorsel:guncelle (ya da CI'da "gorsel_temel" girdisiyle elle tetikleme; bkz. e2e/gorsel/__ekran__/README.md).`);
    return 0;
  }
  if (guncelle && hepsi) rmSync(dizin, { recursive: true, force: true });
  const kip = guncelle ? (hepsi ? "all" : "changed") : "none";
  const r = spawnSync("pnpm", ["exec", "playwright", "test", "-c", "playwright.gorsel.config.ts", `--update-snapshots=${kip}`, ...pwArg], {
    cwd: KOK,
    stdio: "inherit",
    env: { ...process.env, OH_GORSEL_EKRAN: ekran },
    shell: process.platform === "win32",
  });
  if (guncelle) {
    const sonra = ozet(dizin);
    const yeni = [...sonra.keys()].filter((f) => !once.has(f));
    const degisen = [...sonra.keys()].filter((f) => once.has(f) && once.get(f) !== sonra.get(f));
    const silinen = [...once.keys()].filter((f) => !sonra.has(f));
    const L = [`Görsel temeller: ${dizin}`, `- toplam ${sonra.size} görüntü; yeni ${yeni.length}, değişen ${degisen.length}, silinen ${silinen.length}`];
    for (const [ad, liste] of [["yeni", yeni], ["değişen", degisen], ["silinen", silinen]]) if (liste.length) L.push(`- ${ad}: ${liste.sort().join(", ")}`);
    console.log("\n" + L.join("\n"));
    if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, `### ${L[0]}\n\n${L.slice(1).join("\n")}\n`);
  }
  return r.status ?? 1;
}

function docker() {
  const info = spawnSync("docker", ["info", "--format", "{{.ServerVersion}}"], { encoding: "utf8", timeout: 20_000 });
  if (info.error || info.status !== 0) {
    const neden = info.error?.code === "ENOENT" ? "docker komutu bulunamadı" : (info.stderr || info.stdout || String(info.error ?? "")).trim().split("\n")[0] || "Docker hizmeti yanıt vermiyor";
    dur(
      [
        "Görsel regresyon testleri, temel görüntülerin üretildiği ortamla birebir aynı Docker imajında çalışır:",
        `  ${IMAJ} + Carlito, Caladea, Liberation, DejaVu yazı tipleri (sürüm sabit).`,
        `Bu makinede Docker kullanılamıyor (${neden}).`,
        "",
        "Seçenekler:",
        "  1) Docker'ı kurun ya da hizmetini başlatın ve yeniden deneyin: pnpm gorsel",
        "  2) Temelleri CI'da üretin: GitHub › Actions › CI › Run workflow, \"gorsel_temel\" işaretli. İş bitince",
        "     \"gorsel-temel\" artefaktındaki PNG'leri e2e/gorsel/__ekran__/linux/ altına koyup commit'leyin.",
        "     (PR'a \"gorsel-degisti\" etiketi eklemek de aynı artefaktı üretir.)",
        "  3) Yalnız bu makinede belirlenimcilik denemesi (temeller depoya yazılmaz, CI ile karşılaştırılamaz):",
        "     OH_GORSEL_YEREL=1 pnpm gorsel:guncelle && OH_GORSEL_YEREL=1 pnpm gorsel",
      ].join("\n"),
      2,
    );
  }
  // Güncellemede /work'teki temel dizini (depodakilerin kopyası + yeni yazılanlar) ana makinedekinin yerine geçer.
  const geri = guncelle
    ? "if [ -d e2e/gorsel/__ekran__/linux ]; then rm -rf /src/e2e/gorsel/__ekran__/linux && mkdir -p /src/e2e/gorsel/__ekran__ && cp -r e2e/gorsel/__ekran__/linux /src/e2e/gorsel/__ekran__/; fi"
    : "true";
  const betik = [
    "set -euo pipefail",
    "export DEBIAN_FRONTEND=noninteractive ELECTRON_SKIP_BINARY_DOWNLOAD=1",
    `apt-get update -qq && apt-get install -y -qq --no-install-recommends ${YAZI_TIPLERI} >/dev/null && fc-cache -f`,
    "mkdir -p /work",
    "tar -C /src --exclude=node_modules --exclude=dist --exclude=dist-electron --exclude=dist-web --exclude=release --exclude=.git --exclude=.derlem --exclude=test-results --exclude=playwright-report -cf - . | tar -C /work -xf -",
    "cd /work",
    "npm install -g --silent pnpm@10.28.0",
    "pnpm config set store-dir /pnpm-store >/dev/null",
    "pnpm install --frozen-lockfile --prefer-offline",
    "pnpm -r --filter ./apps/belge --filter ./apps/tablo --filter ./apps/sunum run build",
    "set +e; node e2e/gorsel/calistir.mjs \"$@\"; kod=$?; set -e",
    "rm -rf /src/test-results/gorsel /src/playwright-report/gorsel && mkdir -p /src/test-results /src/playwright-report",
    "if [ -d test-results/gorsel ]; then cp -r test-results/gorsel /src/test-results/; fi",
    "if [ -d playwright-report/gorsel ]; then cp -r playwright-report/gorsel /src/playwright-report/; fi",
    geri,
    'if [ -n "${OH_UID:-}" ]; then chown -R "$OH_UID:${OH_GID:-$OH_UID}" /src/test-results /src/playwright-report /src/e2e/gorsel/__ekran__ 2>/dev/null || true; fi',
    "exit $kod",
  ].join("\n");
  const ortam = ["-e", "CI=1", "-e", "OH_GORSEL_KONTEYNER=1"];
  for (const k of ["OH_E2E_PORT", "OH_GORSEL_KAYDIR"]) if (process.env[k]) ortam.push("-e", `${k}=${process.env[k]}`);
  if (process.getuid) ortam.push("-e", `OH_UID=${process.getuid()}`, "-e", `OH_GID=${process.getgid()}`);
  const args = ["run", "--rm", "--init", "--ipc=host", "-v", `${KOK}:/src`, "-v", "officehere-gorsel-pnpm:/pnpm-store", ...ortam, IMAJ, "bash", "-c", betik, "gorsel"];
  args.push(...(guncelle ? ["--guncelle"] : []), ...(hepsi ? ["--hepsi"] : []), ...pwArg);
  console.log(`Docker: ${IMAJ} (ilk koşuda imaj indirilir)`);
  return spawnSync("docker", args, { stdio: "inherit" }).status ?? 1;
}

process.exitCode = process.env.OH_GORSEL_KONTEYNER === "1" || process.env.OH_GORSEL_YEREL === "1" ? dogrudan() : docker();
