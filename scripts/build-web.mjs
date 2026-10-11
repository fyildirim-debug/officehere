// Web sürümü: başlangıç ekranı kökte, uygulamalar alt dizinlerde (dist-web/).
// Herhangi bir statik sunucuyla yayınlanabilir.
import { execSync } from "node:child_process";
import { cp, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const out = path.join(root, "dist-web");
const sh = (cmd) => execSync(cmd, { stdio: "inherit", cwd: root });

for (const app of ["belge", "tablo", "sunum"]) sh(`pnpm --filter @officehere/${app} build`);
sh("pnpm --filter @officehere/desktop build:launcher");

await rm(out, { recursive: true, force: true });
await mkdir(out, { recursive: true });
await cp(path.join(root, "apps/desktop/dist-electron/launcher"), out, { recursive: true });
for (const app of ["belge", "tablo", "sunum"]) {
  await cp(path.join(root, `apps/${app}/dist`), path.join(out, app), { recursive: true });
}
console.log(`Web sürümü hazır: ${out}`);
