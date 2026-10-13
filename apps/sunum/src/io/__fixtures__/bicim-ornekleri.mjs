#!/usr/bin/env node
// Sunu biçimi test örneklerini LibreOffice ile üretir (yeniden üretilebilsin diye).
// Kullanım: node apps/sunum/src/io/__fixtures__/bicim-ornekleri.mjs
// Kaynak: elle yazılmış düz OpenDocument sunusu (.fodp) → .odp, .otp, .ppt, .ppsx;
// ayrıca mevcut lo-proje.pptx → lo-proje.ppt (tablo ve notlu daha büyük örnek).
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
// Node 22.18+ TypeScript türlerini kendisi ayıklar.
import { readCfb } from "../../../../../packages/core/src/cfb.ts";
import { writeCfb } from "../ppt/testCfb.ts";

const here = dirname(fileURLToPath(import.meta.url));
const work = mkdtempSync(join(tmpdir(), "oh-sunum-ornek-"));

// 8×8 mavi PNG.
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGOQqziBFTEMLQkABkZXgRDDjAEAAAAASUVORK5CYII=";

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
].join(" ");

const frame = (cls, x, y, w, h, body) =>
  `<draw:frame presentation:class="${cls}" presentation:style-name="pr1" svg:x="${x}cm" svg:y="${y}cm" svg:width="${w}cm" svg:height="${h}cm"><draw:text-box>${body}</draw:text-box></draw:frame>`;
const notes = (t) => `<presentation:notes><draw:page-thumbnail svg:x="2cm" svg:y="2cm" svg:width="17cm" svg:height="9.5cm"/><draw:frame presentation:class="notes" svg:x="2cm" svg:y="13cm" svg:width="17cm" svg:height="12cm"><draw:text-box><text:p>${t}</text:p></draw:text-box></draw:frame></presentation:notes>`;

const fodp = `<?xml version="1.0" encoding="UTF-8"?>
<office:document ${NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">
<office:meta><dc:title>Biçim Örneği</dc:title><meta:initial-creator>OfficeHere Test</meta:initial-creator></office:meta>
<office:styles>
<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none"/><style:text-properties fo:font-family="Liberation Sans" fo:font-size="18pt"/></style:style>
</office:styles>
<office:automatic-styles>
<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm" style:print-orientation="landscape"/></style:page-layout>
<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:background-visible="true" presentation:background-objects-visible="true"/></style:style>
<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#fff2cc" presentation:background-visible="true" presentation:background-objects-visible="true"/></style:style>
<style:style style:name="pr1" style:family="presentation"><style:graphic-properties draw:fill="none" draw:stroke="none"/></style:style>
<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#c00000" draw:stroke="solid" svg:stroke-color="#800000" svg:stroke-width="0.05cm" draw:textarea-horizontal-align="center" draw:textarea-vertical-align="middle" draw:auto-grow-height="false" draw:auto-grow-width="false"/></style:style>
<style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#70ad47" draw:stroke="none"/></style:style>
<style:style style:name="gr3" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-color="#2f5597" svg:stroke-width="0.08cm"/></style:style>
<style:style style:name="gr4" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none"/></style:style>
<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>
<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#1f4e79" fo:font-size="32pt"/></style:style>
<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
<style:style style:name="T3" style:family="text"><style:text-properties fo:color="#ffffff" fo:font-weight="bold"/></style:style>
<style:style style:name="T4" style:family="text"><style:text-properties fo:color="#c55a11" style:font-name="Liberation Serif" fo:font-family="'Liberation Serif'" fo:font-size="24pt"/></style:style>
</office:automatic-styles>
<office:master-styles>
<style:master-page style:name="Varsayilan" style:page-layout-name="PM1" draw:style-name="dp1"/>
</office:master-styles>
<office:body><office:presentation>
<draw:page draw:name="Kapak" draw:style-name="dp1" draw:master-page-name="Varsayilan" presentation:presentation-page-layout-name="AL1T0">
${frame("title", 2, 4, 24, 3, '<text:p text:style-name="P1"><text:span text:style-name="T1">Çeyrek Sonuçları — Iğdır Şubesi</text:span></text:p>')}
${frame("subtitle", 2, 8, 24, 2.5, '<text:p text:style-name="P1">Satış ve Pazarlama Ekibi · Şükrü Öztürk</text:p>')}
${notes("Açılış: herkese hoş geldiniz deyin.")}
</draw:page>
<draw:page draw:name="Gündem" draw:style-name="dp1" draw:master-page-name="Varsayilan" presentation:presentation-page-layout-name="AL1T1">
${frame("title", 1.5, 0.8, 25, 2.5, "<text:p>Gündem</text:p>")}
${frame(
  "outline",
  1.5,
  3.6,
  25,
  10.5,
  "<text:list><text:list-item><text:p>Geçen çeyreğin özeti</text:p></text:list-item><text:list-item><text:p>Yeni ürünler</text:p><text:list><text:list-item><text:p>Akıllı saat çığır açıyor</text:p></text:list-item><text:list-item><text:p>Kablosuz kulaklık</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Soru ve cevap</text:p></text:list-item></text:list>",
)}
${notes("Konuşmacı notu: sözü Ayşe Hanım'a bırakın.")}
</draw:page>
<draw:page draw:name="Görsel ve Şekil" draw:style-name="dp1" draw:master-page-name="Varsayilan" presentation:presentation-page-layout-name="AL1T19">
${frame("title", 1.5, 0.8, 25, 2.5, "<text:p>Görsel ve Şekiller</text:p>")}
<draw:frame draw:name="Mavi Resim" svg:x="2cm" svg:y="4cm" svg:width="6cm" svg:height="6cm"><draw:image><office:binary-data>${PNG}</office:binary-data></draw:image><svg:desc>Mavi kare</svg:desc></draw:frame>
<draw:custom-shape draw:name="Kırmızı Kutu" draw:style-name="gr1" svg:x="10cm" svg:y="4cm" svg:width="7cm" svg:height="3cm"><text:p><text:span text:style-name="T3">Kırmızı Kutu</text:span></text:p><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="rectangle" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/></draw:custom-shape>
<draw:custom-shape draw:name="Yeşil Elips" draw:style-name="gr2" draw:transform="rotate (0.5235987755983) translate (19.5cm 6.5cm)" svg:width="5cm" svg:height="3cm"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="ellipse"/></draw:custom-shape>
<draw:line draw:name="Mavi Çizgi" draw:style-name="gr3" svg:x1="10cm" svg:y1="9cm" svg:x2="26cm" svg:y2="9cm"/>
<draw:frame draw:name="Biçimli Metin" draw:style-name="gr4" svg:x="10cm" svg:y="10cm" svg:width="16cm" svg:height="3cm"><draw:text-box><text:p>Düz metin, <text:span text:style-name="T2">eğik ve altı çizili</text:span>, <text:span text:style-name="T4">serif turuncu</text:span>.</text:p></draw:text-box></draw:frame>
</draw:page>
<draw:page draw:name="Kapanış" draw:style-name="dp2" draw:master-page-name="Varsayilan" presentation:presentation-page-layout-name="AL1T19">
${frame("title", 2, 5, 24, 3, '<text:p text:style-name="P1">Teşekkürler!</text:p>')}
</draw:page>
</office:presentation></office:body>
</office:document>
`;

const src = join(work, "lo-bicim.fodp");
writeFileSync(src, fodp);

// Önizleme resmi kapalı (yoksa .ppt özet akışı yüzlerce KB olur).
mkdirSync(join(work, "profil", "user"), { recursive: true });
writeFileSync(
  join(work, "profil", "user", "registrymodifications.xcu"),
  `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Save/Document"><prop oor:name="GenerateThumbnail" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
`,
);

function convert(file, filter, outName) {
  execFileSync("timeout", ["120", "soffice", `-env:UserInstallation=file://${work}/profil`, "--headless", "--convert-to", filter, file, "--outdir", work], { stdio: "inherit" });
  const ext = filter.split(":")[0];
  const base = file.split("/").pop().replace(/\.[^.]+$/, "");
  const produced = join(work, `${base}.${ext}`);
  const target = join(here, outName);
  if (ext === "ppt") {
    // Özet akışındaki önizleme resmi atılır (dosya küçük kalsın).
    const cfb = readCfb(readFileSync(produced));
    const streams = cfb.entries.filter((e) => e.type === "stream" && !/SummaryInformation$/.test(e.name)).map((e) => ({ name: e.name, data: cfb.read(e.path) }));
    writeFileSync(target, writeCfb(streams));
  } else copyFileSync(produced, target);
  const size = statSync(target).size;
  if (size > 100_000) throw new Error(`${outName} çok büyük: ${size} bayt`);
  console.log(`${outName}: ${size} bayt`);
}

convert(src, "odp", "lo-bicim.odp");
convert(src, "otp:impress8_template", "lo-bicim.otp");
convert(src, "ppt", "lo-bicim.ppt");
convert(src, 'ppsx:Impress MS PowerPoint 2007 XML AutoPlay', "lo-bicim.ppsx");

// Daha büyük .ppt örneği: mevcut LibreOffice PPTX örneğinden.
const proje = join(work, "lo-proje.pptx");
writeFileSync(proje, readFileSync(join(here, "lo-proje.pptx")));
convert(proje, "ppt", "lo-proje.ppt");
convert(proje, "odp", "lo-proje.odp");

rmSync(work, { recursive: true, force: true });
