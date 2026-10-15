import { describe, expect, it } from "vitest";
import { OpcPackage } from "../../src/opc/package";
import { SourceRegistry } from "../../src/opc/source";
import { readZip, ZipWriter } from "../../src/opc/zip";
import { manifestXml, mergeSettings, OdfCarry, odfFamily, odfInfo, parseManifest, pruneNamespaces, renameStyleRefs, rootNamespaces, styleRefs, withNamespaces } from "../../src/odf";

const MIME_ODT = "application/vnd.oasis.opendocument.text";
const NS_DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const NS_XLINK = "http://www.w3.org/1999/xlink";
const NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const NS_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NS_LOEXT = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

async function zip(parts: [string, string | Uint8Array, boolean?][]): Promise<Uint8Array> {
  const z = new ZipWriter();
  for (const [n, d, store] of parts) z.add(n, d, { compress: !store });
  return z.finish();
}

const manifest = (mime: string, extra: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/>${extra}</manifest:manifest>`;
const entry = (p: string, t: string, v = "") => `<manifest:file-entry manifest:full-path="${p}"${v ? ` manifest:version="${v}"` : ""} manifest:media-type="${t}"/>`;

/** LibreOffice'in yazdığına benzer kaynak: ayarlar, RDF, küçük resim ve gömülü grafik. */
async function source(mime = MIME_ODT): Promise<Uint8Array> {
  const chart = `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>`;
  return zip([
    ["mimetype", mime, true],
    ["content.xml", "<x/>"],
    ["styles.xml", "<x/>"],
    ["settings.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewLeft" config:type="int">7</config:config-item></config:config-item-set><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="AddParaTableSpacing" config:type="boolean">false</config:config-item></config:config-item-set></office:settings></office:document-settings>`],
    ["manifest.rdf", `<?xml version="1.0" encoding="utf-8"?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"/>`],
    ["Thumbnails/thumbnail.png", new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5, 6, 7, 8])],
    ["Object 1/content.xml", chart],
    ["Object 1/styles.xml", chart],
    ["ObjectReplacements/Object 1", new Uint8Array([0x56, 0x43, 0x4c, 0x4d, 0x54, 0x46, 1, 0])],
    [
      "META-INF/manifest.xml",
      manifest(
        mime,
        entry("content.xml", "text/xml") +
          entry("styles.xml", "text/xml") +
          entry("settings.xml", "text/xml") +
          entry("manifest.rdf", "application/rdf+xml") +
          entry("Thumbnails/thumbnail.png", "image/png") +
          entry("Object 1/content.xml", "text/xml") +
          entry("Object 1/styles.xml", "text/xml") +
          entry("Object 1/", "application/vnd.oasis.opendocument.chart", "1.3") +
          entry("ObjectReplacements/Object 1", 'application/x-openoffice-gdimetafile;windows_formatname="GDIMetaFile"'),
      ),
    ],
  ]);
}

/** Yazıcı çıktısına benzer paket (JSZip gibi: mimetype ilk). */
function written(extra: [string, string | Uint8Array, boolean?][] = [], mime = MIME_ODT): Promise<Uint8Array> {
  return zip([
    ["mimetype", mime, true],
    ["content.xml", "<content/>"],
    ["styles.xml", "<styles/>"],
    ...extra,
    ["META-INF/manifest.xml", manifest(mime, entry("content.xml", "text/xml") + entry("styles.xml", "text/xml"))],
  ]);
}

const FRAME = `<draw:frame xmlns:draw="${NS_DRAW}" xmlns:xlink="${NS_XLINK}" xmlns:svg="${NS_SVG}" xmlns:loext="${NS_LOEXT}" draw:style-name="fr1" svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><loext:p/></draw:object><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
const FR1 = `<style:style xmlns:style="${NS_STYLE}" style:name="fr1" style:family="graphic" style:parent-style-name="OLE"><style:graphic-properties style:wrap="none"/></style:style>`;

async function entries(bytes: Uint8Array) {
  const z = readZip(bytes);
  const pkg = await OpcPackage.open(bytes);
  return { names: z.entries.map((e) => e.name), first: z.entries[0], pkg, manifest: parseManifest(await pkg.text("META-INF/manifest.xml")) };
}

describe("ODF manifest", () => {
  it("girdileri okur ve yazar (öznitelik sırası ve kaçışlar)", () => {
    const xml = manifest(MIME_ODT, entry("Object 1/", "application/vnd.oasis.opendocument.chart", "1.3") + entry("ObjectReplacements/Object 1", "application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;"));
    const list = parseManifest(xml);
    expect(list).toEqual([
      { path: "/", mediaType: MIME_ODT, version: "1.3" },
      { path: "Object 1/", mediaType: "application/vnd.oasis.opendocument.chart", version: "1.3" },
      { path: "ObjectReplacements/Object 1", mediaType: 'application/x-openoffice-gdimetafile;windows_formatname="GDIMetaFile"' },
    ]);
    expect(parseManifest(manifestXml(list))).toEqual(list);
  });

  it("aile: şablonlar dahil; ODF olmayan null", () => {
    expect(odfFamily("application/vnd.oasis.opendocument.spreadsheet-template")).toBe("spreadsheet");
    expect(odfFamily("application/vnd.oasis.opendocument.presentation")).toBe("presentation");
    expect(odfFamily("application/zip")).toBeNull();
  });
});

describe("ODF XML parçacıkları", () => {
  it("kullanılan öneklerin bildirimleri eklenir, kullanılmayanlar atılır", () => {
    const decls = { draw: NS_DRAW, xlink: NS_XLINK, chart: "urn:chart", of: "urn:of", table: "urn:table" };
    const out = withNamespaces(`<draw:frame draw:name="a"><draw:object xlink:href="x" chart:class="chart:bar"/></draw:frame>`, decls);
    expect(rootNamespaces(out)).toEqual({ draw: NS_DRAW, xlink: NS_XLINK, chart: "urn:chart" });
    const pruned = pruneNamespaces(`<a:x xmlns:a="u1" xmlns:b="u2" xmlns:c="u3" xmlns:a="u1" c:f="of:=1"/>`);
    expect(pruned).toBe(`<a:x xmlns:a="u1" xmlns:c="u3" c:f="of:=1"/>`);
  });

  it("stil başvuruları ve yeniden adlandırma (aile seçimi)", () => {
    const x = `<draw:frame draw:style-name="P1" draw:text-style-name="P1" presentation:class-names="pr1 pr2"/>`;
    expect([...styleRefs(x)]).toEqual(["P1", "pr1", "pr2"]);
    const r = renameStyleRefs(x, new Map([["P1", [["graphic", "g"], ["paragraph", "p"]]], ["pr2", [["presentation", "q"]]]]));
    expect(r).toBe(`<draw:frame draw:style-name="g" draw:text-style-name="p" presentation:class-names="pr1 q"/>`);
  });

  it("settings.xml birleştirme: yazıcının kümeleri geçerli, kaynağın diğerleri eklenir", () => {
    const out = `<?xml version="1.0"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ActiveTable" config:type="string">A</config:config-item></config:config-item-set></office:settings></office:document-settings>`;
    const src = `<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ActiveTable" config:type="string">B</config:config-item></config:config-item-set><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="AutoCalculate" config:type="boolean">true</config:config-item></config:config-item-set></office:settings></office:document-settings>`;
    const m = mergeSettings(out, src)!;
    expect(m).toContain(">A</config:config-item>");
    expect(m).not.toContain(">B<");
    expect(m).toContain('config:name="ooo:configuration-settings"');
    expect(mergeSettings(out, out)).toBeNull();
  });
});

describe("OdfCarry", () => {
  it("aynı aile: settings.xml ve manifest.rdf taşınır, Thumbnails düşer, manifest tam, mimetype ilk ve sıkıştırmasız", async () => {
    const reg = new SourceRegistry();
    const id = reg.register(await OpcPackage.open(await source()));
    const c = new OdfCarry({ family: "text", source: id, registry: reg });
    const out = await entries(await c.finish(await written()));
    expect(out.first.name).toBe("mimetype");
    expect(out.first.method).toBe(0);
    expect(out.names).toContain("settings.xml");
    expect(out.names).toContain("manifest.rdf");
    expect(out.names.some((n) => n.startsWith("Thumbnails/"))).toBe(false);
    expect(out.names.some((n) => n.startsWith("Object"))).toBe(false);
    const paths = out.manifest.map((e) => e.path);
    for (const n of out.names) if (n !== "mimetype" && n !== "META-INF/manifest.xml") expect(paths).toContain(n);
    expect(out.manifest.find((e) => e.path === "manifest.rdf")?.mediaType).toBe("application/rdf+xml");
    expect(c.report.dropped).toContainEqual({ what: "Thumbnails/", reason: "thumbnail" });
    expect(await out.pkg.text("settings.xml")).toContain("AddParaTableSpacing");
  });

  it("yazıcı settings.xml yazdıysa birleştirilir", async () => {
    const reg = new SourceRegistry();
    const id = reg.register(await OpcPackage.open(await source()));
    const own = `<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewLeft" config:type="int">1</config:config-item></config:config-item-set></office:settings></office:document-settings>`;
    const out = await entries(await new OdfCarry({ family: "text", source: id, registry: reg }).finish(await written([["settings.xml", own]])));
    const s = (await out.pkg.text("settings.xml"))!;
    expect(s).toContain(">1</config:config-item>");
    expect(s).toContain("AddParaTableSpacing");
    expect(s.match(/ooo:view-settings/g)).toHaveLength(1);
  });

  it("gömülü nesne: alt belge ve yedek görüntü yeni adla kopyalanır, stiller yeniden adlanır, ikinci örnek ayrı kopya alır", async () => {
    const reg = new SourceRegistry();
    const id = reg.register(await OpcPackage.open(await source()));
    const c = new OdfCarry({ family: "text", source: id, registry: reg });
    const a = c.object({ xml: FRAME, styles: [FR1] })!;
    const b = c.object({ src: id, xml: FRAME, styles: [FR1] })!;
    expect(a.xml).toContain('xlink:href="./Object 1"');
    expect(a.xml).toContain('xlink:href="./ObjectReplacements/Object 1"');
    expect(b.xml).toContain('xlink:href="./Object 2"');
    expect(b.xml).toContain('xlink:href="./ObjectReplacements/Object 2"');
    // Aynı stil bir kez yazılır, iki nesne de onu anar.
    expect(a.styles).toHaveLength(1);
    expect(b.styles).toHaveLength(0);
    const name = /style:name="([^"]+)"/.exec(a.styles[0])![1];
    expect(name).toMatch(/^oh\d+-fr1$/);
    expect(a.xml).toContain(`draw:style-name="${name}"`);
    expect(b.xml).toContain(`draw:style-name="${name}"`);
    // Başka bölümde (ör. üst bilgi stilleri) yeniden yazılır.
    expect(c.object({ xml: FRAME, styles: [FR1] }, "styles")!.styles).toHaveLength(1);
    const out = await entries(await c.finish(await written()));
    for (const n of ["Object 1/content.xml", "Object 1/styles.xml", "ObjectReplacements/Object 1", "Object 2/content.xml", "ObjectReplacements/Object 2", "Object 3/content.xml"]) expect(out.names).toContain(n);
    const man = new Map(out.manifest.map((e) => [e.path, e]));
    expect(man.get("Object 2/")).toEqual({ path: "Object 2/", mediaType: "application/vnd.oasis.opendocument.chart", version: "1.3" });
    expect(man.get("ObjectReplacements/Object 2")?.mediaType).toContain("gdimetafile");
    expect(c.report.objects).toBe(3);
  });

  it("kaynak yoksa ya da başvurulan dosya eksikse null; OPC kaynak ve başka aile paket düzeyinde taşınmaz", async () => {
    const reg = new SourceRegistry();
    expect(new OdfCarry({ family: "text", source: "yok", registry: reg }).object({ xml: FRAME, styles: [] })).toBeNull();
    const id = reg.register(await OpcPackage.open(await source()));
    const broken = FRAME.replace("./Object 1", "./Object 9");
    expect(new OdfCarry({ family: "text", source: id, registry: reg }).object({ xml: broken, styles: [] })).toBeNull();
    // Hesap tablosu kaynağı, metin hedefi: settings.xml taşınmaz ama nesne taşınabilir.
    const other = reg.register(await OpcPackage.open(await source("application/vnd.oasis.opendocument.spreadsheet")));
    const c = new OdfCarry({ family: "text", source: other, registry: reg });
    const w = await written();
    expect(await c.finish(w)).toBe(w);
    // OPC (docx) kaynak: ODF sayılmaz.
    const docx = await zip([["[Content_Types].xml", "<Types/>"], ["word/document.xml", "<w/>"]]);
    const did = reg.register(await OpcPackage.open(docx));
    const d = new OdfCarry({ family: "text", source: did, registry: reg });
    expect(d.source).toBeUndefined();
    expect(await d.finish(w)).toBe(w);
  });

  it("odfInfo: mimetype, aile ve manifest", async () => {
    const info = await odfInfo(await OpcPackage.open(await source()));
    expect(info?.family).toBe("text");
    expect(info?.manifest.some((e) => e.path === "Object 1/")).toBe(true);
  });
});
