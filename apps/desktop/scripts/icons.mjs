// Uygulama ve dosya türü simgelerini SVG'den üretir (PNG + ICO).
// Kullanım: pnpm --filter @officehere/desktop icons   (Playwright Chromium gerekir)
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const build = path.join(here, "..", "build");

let chromium;
try {
  ({ chromium } = require("playwright"));
} catch {
  ({ chromium } = require("/opt/node22/lib/node_modules/playwright"));
}

const SUITE = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#4a4aa6"/><stop offset="1" stop-color="#30307a"/></linearGradient></defs>
  <rect x="2" y="2" width="60" height="60" rx="14" fill="url(#g)"/>
  <rect x="13" y="15" width="38" height="8" rx="2.5" fill="#6d9bff"/>
  <rect x="13" y="28" width="38" height="8" rx="2.5" fill="#48c98a"/>
  <rect x="13" y="41" width="38" height="8" rx="2.5" fill="#ff8a66"/>
</svg>`;

const FILETYPES = {
  belge: { color: "#1f5fbf", letter: "B" },
  tablo: { color: "#1d7a46", letter: "T" },
  sunum: { color: "#c4462b", letter: "S" },
};

const GLYPH = {
  // Belge: metin satırları
  belge: '<g stroke="#c8c8c8" stroke-width="2.4" stroke-linecap="round"><path d="M18 14h22M18 20h28M18 26h28M18 32h20"/></g>',
  // Tablo: hücre ızgarası
  tablo: '<g fill="none" stroke="#c8c8c8" stroke-width="1.8"><rect x="17" y="11" width="30" height="24"/><path d="M17 19h30M17 27h30M27 11v24M37 11v24"/></g>',
  // Sunum: slayt çerçevesi
  sunum: '<g fill="none" stroke="#c8c8c8" stroke-width="2"><rect x="17" y="12" width="30" height="20" rx="1.5"/><path d="M23 19h18M26 25h12"/></g>',
};

const fileSvg = (name, { color, letter }) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <path d="M10 3h32l12 12v46H10z" fill="#fff" stroke="#9a9a9a" stroke-width="1.5"/>
  <path d="M42 3v12h12" fill="#ececec" stroke="#9a9a9a" stroke-width="1.5"/>
  ${GLYPH[name]}
  <path d="M10 40h44v21H10z" fill="${color}"/>
  <text x="32" y="57" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-weight="700" font-size="17" fill="#fff">${letter}</text>
</svg>`;

/** PNG verilerinden ICO dosyası oluşturur (Vista+ PNG gömülü ICO). */
function ico(pngs) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(pngs.length, 4);
  const entries = [];
  let offset = 6 + pngs.length * 16;
  for (const { size, data } of pngs) {
    const e = Buffer.alloc(16);
    e.writeUInt8(size >= 256 ? 0 : size, 0);
    e.writeUInt8(size >= 256 ? 0 : size, 1);
    e.writeUInt8(0, 2);
    e.writeUInt8(0, 3);
    e.writeUInt16LE(1, 4);
    e.writeUInt16LE(32, 6);
    e.writeUInt32LE(data.length, 8);
    e.writeUInt32LE(offset, 12);
    offset += data.length;
    entries.push(e);
  }
  return Buffer.concat([header, ...entries, ...pngs.map((p) => p.data)]);
}

/** PNG verilerinden macOS ICNS dosyası oluşturur (ic07/ic08/ic09/ic10 PNG girdileri). */
function icns(entries) {
  const parts = entries.map(({ type, data }) => {
    const h = Buffer.alloc(8);
    h.write(type, 0, "ascii");
    h.writeUInt32BE(data.length + 8, 4);
    return Buffer.concat([h, data]);
  });
  const body = Buffer.concat(parts);
  const head = Buffer.alloc(8);
  head.write("icns", 0, "ascii");
  head.writeUInt32BE(body.length + 8, 4);
  return Buffer.concat([head, body]);
}

const ICNS_SIZES = [
  ["ic07", 128],
  ["ic08", 256],
  ["ic09", 512],
  ["ic10", 1024],
];

const browser = await chromium.launch();
const page = await browser.newPage();

async function render(svg, size) {
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(`<html><body style="margin:0;background:transparent">${svg.replace("<svg ", `<svg width="${size}" height="${size}" `)}</body></html>`);
  return page.screenshot({ omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
}

const ICO_SIZES = [16, 24, 32, 48, 64, 128, 256];

await mkdir(path.join(build, "icons"), { recursive: true });
await mkdir(path.join(build, "filetypes"), { recursive: true });

await writeFile(path.join(build, "icon.png"), await render(SUITE, 1024));
const suitePngs = [];
for (const s of ICO_SIZES) suitePngs.push({ size: s, data: await render(SUITE, s) });
await writeFile(path.join(build, "icon.ico"), ico(suitePngs));
for (const s of [16, 32, 48, 64, 128, 256, 512]) {
  await writeFile(path.join(build, "icons", `${s}x${s}.png`), await render(SUITE, s));
}

for (const [name, spec] of Object.entries(FILETYPES)) {
  const svg = fileSvg(name, spec);
  const pngs = [];
  for (const s of ICO_SIZES) pngs.push({ size: s, data: await render(svg, s) });
  await writeFile(path.join(build, "filetypes", `${name}.ico`), ico(pngs));
  await writeFile(path.join(build, "filetypes", `${name}.png`), await render(svg, 512));
  const icnsEntries = [];
  for (const [type, size] of ICNS_SIZES) icnsEntries.push({ type, data: await render(svg, size) });
  await writeFile(path.join(build, "filetypes", `${name}.icns`), icns(icnsEntries));
}

{
  const suiteIcns = [];
  for (const [type, size] of ICNS_SIZES) suiteIcns.push({ type, data: await render(SUITE, size) });
  await writeFile(path.join(build, "icon.icns"), icns(suiteIcns));
}

// Mağaza (AppX) paketi görselleri: electron-builder build/appx altındakileri kullanır; yoksa
// kendi örnek görsellerini koyar ve mağaza "varsayılan görsel" diye reddeder. Arka plan saydam.
{
  await mkdir(path.join(build, "appx"), { recursive: true });
  const tile = async (w, h, icon) => {
    await page.setViewportSize({ width: w, height: h });
    const svg = SUITE.replace("<svg ", `<svg width="${icon}" height="${icon}" `);
    await page.setContent(`<html><body style="margin:0;background:transparent;display:flex;align-items:center;justify-content:center;width:${w}px;height:${h}px">${svg}</body></html>`);
    return page.screenshot({ omitBackground: true, clip: { x: 0, y: 0, width: w, height: h } });
  };
  for (const [name, w, h, icon] of [
    ["StoreLogo", 50, 50, 50],
    ["Square44x44Logo", 44, 44, 44],
    ["Square150x150Logo", 150, 150, 96],
    ["Wide310x150Logo", 310, 150, 96],
  ]) {
    await writeFile(path.join(build, "appx", `${name}.png`), await tile(w, h, icon));
  }
}

await browser.close();
console.log("Simgeler üretildi:", build);
