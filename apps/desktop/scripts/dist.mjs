// Paketleme: web uygulamalarını ve masaüstü kabuğunu derler, electron-builder'ı çalıştırır.
// Kullanım: pnpm --filter @officehere/desktop dist -- --win   (ya da --linux, --mac)
import { execSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const desktop = path.join(here, "..");
const root = path.join(desktop, "..", "..");
const sh = (cmd, cwd = root) => execSync(cmd, { stdio: "inherit", cwd });

for (const app of ["belge", "tablo", "sunum"]) sh(`pnpm --filter @officehere/${app} build`);
sh("pnpm --filter @officehere/desktop build");

const args = process.argv.slice(2).filter((a) => a !== "--");

// Mağaza (Store) paketi: OH_STORE=1 pakete işaret koyar (updates.ts güncelleştirme denetimini kapatır);
// kimlik ve yayıncı Partner Center'dan gelir (verilmezse electron-builder.yml'deki yer tutucular).
if (process.env.OH_STORE === "1") {
  args.push("-c.extraMetadata.ohDagitim=store");
  for (const [env, key] of [
    ["OH_STORE_KIMLIK", "identityName"],
    ["OH_STORE_YAYINCI", "publisher"],
  ]) {
    const value = process.env[env]?.trim();
    if (!value) continue;
    if (!/^[\w.=, -]+$/.test(value)) throw new Error(`${env} geçersiz karakter içeriyor: ${value}`);
    args.push(`"-c.appx.${key}=${value}"`);
  }
}

sh(`pnpm exec electron-builder --config electron-builder.yml ${args.join(" ")}`, desktop);
