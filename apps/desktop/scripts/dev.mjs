// Geliştirme: üç uygulamanın ve başlangıç ekranının Vite sunucularını başlatır,
// ana süreci derler ve Electron'u OH_DEV=1 ile açar.
import { spawn } from "node:child_process";
import net from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const desktop = path.join(here, "..");
const root = path.join(desktop, "..", "..");
const children = [];
const shell = process.platform === "win32";

function run(cmd, args, opts = {}) {
  const p = spawn(cmd, args, { stdio: "inherit", shell, ...opts });
  children.push(p);
  return p;
}

function waitPort(port, timeout = 60_000) {
  const start = Date.now();
  return new Promise((resolve, reject) => {
    const tryOnce = () => {
      const s = net.connect(port, "127.0.0.1");
      s.once("connect", () => (s.end(), resolve()));
      s.once("error", () => {
        s.destroy();
        if (Date.now() - start > timeout) reject(new Error(`Port ${port} açılmadı`));
        else setTimeout(tryOnce, 300);
      });
    };
    tryOnce();
  });
}

for (const app of ["belge", "tablo", "sunum"]) run("pnpm", ["--filter", `@officehere/${app}`, "dev"], { cwd: root });
run("pnpm", ["exec", "vite"], { cwd: desktop });

const tsc = run("pnpm", ["exec", "tsc", "-p", "tsconfig.electron.json"], { cwd: desktop });
await new Promise((r) => tsc.once("exit", r));
await Promise.all([5173, 5174, 5175, 5176].map((p) => waitPort(p)));

const electron = run("pnpm", ["exec", "electron", "."], { cwd: desktop, env: { ...process.env, OH_DEV: "1" } });
electron.once("exit", () => {
  for (const c of children) c.kill();
  process.exit(0);
});

process.on("SIGINT", () => {
  for (const c of children) c.kill();
  process.exit(0);
});
