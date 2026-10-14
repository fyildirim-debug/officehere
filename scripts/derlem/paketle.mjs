// Uygulamaların dosya okuyucu/yazıcılarını (src/io) Node'da çalışacak tek dosyalık ESM paketlerine çevirir.
//
// Tarayıcı derlemesine yakın olsun diye paketlerin "browser" alanları kullanılır (exceljs, jszip, mammoth
// tarayıcı sürümleri); DOM API'leri işçide jsdom ile sağlanır (dom.mjs). Çıktı: <CIKTI>/paket/<uygulama>-io.mjs
// (+ kaynak haritası; takılma yığınlarını özgün dosya:satıra çevirmek için).
//
// Kullanım: node scripts/derlem/paketle.mjs   (tur.mjs ve bulanik/calistir.mjs bunu kendiliğinden çağırır)
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { CIKTI, KOK, PAKET, posix } from "./ortak.mjs";

/** Her uygulamadan dışa alınan adlar; ad değişirse paketleme açık bir hatayla durur. */
const GIRIS = {
  belge: [["apps/belge/src/io/index.ts", ["openDocument", "writeDocx", "writeOdt", "writeRtf", "writeMarkdown", "docToText", "validateContent"]]],
  tablo: [
    ["apps/tablo/src/io/formats.ts", ["openBook", "writeBook", "writableAs"]],
    ["apps/tablo/src/io/xlsx.ts", ["writeXlsx"]],
    ["apps/tablo/src/model/workbook.ts", ["WorkbookModel"]],
    ["packages/formula/src/index.ts", ["cellKey", "keyRow", "keyCol", "isError"]],
  ],
  sunum: [["apps/sunum/src/io/formats.ts", ["readPresentationFile", "writePresentationFormat"]]],
};

export async function paketle({ sessiz = false } = {}) {
  const { build } = await import(pathToFileURL(join(KOK, "node_modules/vite/dist/node/index.js")).href);
  const giris = join(CIKTI, "giris");
  mkdirSync(giris, { recursive: true });
  mkdirSync(PAKET, { recursive: true });
  const sureler = {};
  for (const [app, parcalar] of Object.entries(GIRIS)) {
    const t = performance.now();
    const kaynak = parcalar.map(([dosya, adlar]) => `export { ${adlar.join(", ")} } from ${JSON.stringify(posix(join(KOK, dosya)))};`).join("\n");
    const girisDosyasi = join(giris, `${app}.ts`);
    writeFileSync(girisDosyasi, kaynak + "\n");
    await build({
      configFile: false,
      root: join(KOK, "apps", app),
      logLevel: "warn",
      publicDir: false,
      mode: "production",
      define: { "process.env.NODE_ENV": '"production"' },
      ssr: { noExternal: true, target: "webworker", resolve: { conditions: ["browser", "module", "import", "default"], externalConditions: [] } },
      resolve: { mainFields: ["browser", "module", "main"], conditions: ["browser"] },
      build: {
        ssr: girisDosyasi,
        outDir: PAKET,
        emptyOutDir: false,
        minify: false,
        sourcemap: true,
        target: "node22",
        rollupOptions: { output: { entryFileNames: `${app}-io.mjs`, format: "esm", codeSplitting: false } },
      },
    });
    sureler[app] = Math.round(performance.now() - t);
  }
  if (!sessiz) console.log(`paketlendi: ${Object.entries(sureler).map(([k, v]) => `${k} ${v} ms`).join(", ")}`);
  return sureler;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) await paketle();
