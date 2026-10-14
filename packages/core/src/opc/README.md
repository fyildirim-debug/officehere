# `@officehere/core` OPC paket katmanı

Geçiş katmanının temeli: açılan `.docx/.xlsx/.pptx` paketini bellekte tutar, kaydederken yazıcının
ürettiği pakete modelin tanımadığı parçaları taşır. Tasarım: teknik tasarım K§1.2.1–1.2.3, 1.2.6.

**API dondurulmuştur (2. aşama, dalga 2).** Dalga 3'te Belge (P3.1/P3.2), Tablo (P3.3) ve Sunum (P3.4)
bu imzalara göre çalışır. İmza değişikliği gerekirse core sahibine (P2.1) iletilir; uygulama tarafında
politika genişletmesi için `overrides` kullanılır, `policy.ts` düzenlenmez.

```ts
import { OpcPackage, registerSource, getSource, carryParts, Carry, collectRefs, reattach, XmlPart } from "@officehere/core";
// ya da yalnız bu katman: import … from "@officehere/core/opc";
```

Bağımlılık yoktur. ZIP açma/sıkıştırma tarayıcının ve Node'un yerel `DecompressionStream` /
`CompressionStream` akışlarıyla yapılır (Electron, tarayıcı, Worker, Node 22, vitest `node` ortamı).

## Dosyalar ve sahiplik

| Dosya | İçerik | Sahibi |
|---|---|---|
| `zip.ts` | `readZip` (merkez dizin), `inflateEntry`, `ZipWriter` (ham kopya + yeni girdi), `crc32`, `OpcError` | P2.1 |
| `package.ts` | `OpcPackage`, `ContentTypes`, `RelsGraph`, `RelsBuilder`, `parseRels`, `explicitRelIds`, `rewriteRelIds`, yol yardımcıları | P2.1 |
| `policy.ts` | Politika tablosu: `classifyRel`, `OWN_TYPES`, `DROP_TYPES`, `MACRO_TYPES`, makro içerik türleri | P2.1 |
| `carry.ts` | Paket düzeyi `carryParts`, öğe düzeyi `Carry`, `collectRefs`, `reattach` | P2.1 |
| `source.ts` | `PackageSource` kaydı (`registerSource`, `getSource`, `releaseSource`), `sha256Hex` | P2.1 |
| `xmlPart.ts` | `XmlPart` (Tablo'dan taşındı) + sıra tabloları: `WORKSHEET_ORDER`, `WORKBOOK_ORDER`, `SETTINGS_ORDER`, `PRESENTATION_ORDER` | P2.1 (Tablo'da `apps/tablo/src/io/xmlPart.ts` yalnız yeniden dışa aktarım) |

Uygulama okuyucu/yazıcıları (`apps/*/src/io/*`) bu katmanı **çağırır**, değiştirmez.

## Akış

```
açılış:   bayt ──► OpcPackage.open ──► registerSource ──► model.kaynak = "kaynak-3"   (model yalnız kimlik taşır)
             └──► (uygulamanın bugünkü okuyucusu, değişmeden)
kaydetme: model ──► yazıcı (JSZip, bugünkü gibi) ──► bayt
                    │  öğe düzeyi: opak XML → reattach(…, carry.instance(kaynak), rels.add…)
                    └► carryParts({ app, source: model.kaynak, output: bayt, carry, macro, partMap }) ──► { bytes, report }
```

Biçim değişince (docx → odt, xlsx → ods, docx → rtf) `carryParts` çağrılmaz.

## API

### Paket

```ts
class OpcPackage {
  static open(data: Uint8Array | ArrayBuffer, opts?: { maxPartSize?: number }): Promise<OpcPackage>;
  readonly bytes: Uint8Array;               // özgün baytlar (değiştirilmez; IndexedDB'ye bu yazılır)
  readonly entries: readonly ZipEntry[];    // ZIP girdileri (dizinler, yinelenenler, bozuklar hariç)
  readonly skippedEntries: number;
  readonly graph: RelsGraph;
  readonly contentTypes: ContentTypes;      // değiştirmeden önce clone()
  readonly mainPart: string | null;         // kökteki officeDocument hedefi
  readonly mainContentType: string | undefined;
  readonly macroEnabled: boolean;           // makrolu ana tür ya da vbaProject ilişkisi
  readonly signed: boolean;                 // dijital imza var → "Kaydederseniz imza kaldırılır" uyarısı
  readonly isOpc: boolean;                  // [Content_Types].xml var mı
  readonly names: string[];
  has(name): boolean;                       // parça adları büyük/küçük harf duyarsız
  canonicalName(name): string | undefined;
  entry(name): ZipEntry | undefined;
  contentType(name): string | undefined;
  sizeOf(name): number;
  data(name): Promise<Uint8Array | null>;   // açılmış bayt; dönen dizi değiştirilmemeli
  text(name): Promise<string | null>;       // UTF-8/UTF-16 (BOM'lu) metin
  rels(source): Promise<Relationship[]>;    // "" = paket kökü; .rels yoksa/bozuksa []
  explicitRelIds(part): Promise<Set<string>>; // önbellekli
}

interface Relationship { id: string; type: string; target: string; external: boolean; part: string | null }

class RelsGraph { rels(source); targets(source); reachable(roots = [""]): Promise<Set<string>> }   // döngü korumalı
class ContentTypes { static parse(xml); get(part); getOverride; getDefault; setOverride; setDefault; deleteOverride; clone(); toXml() }
class RelsBuilder {                         // yeni kimlik asla var olanla/eskiyle çakışmaz
  constructor(source: string, existing?: Relationship[]);
  addPart(type, absolutePart): string;      // hedefi kaynağa göre göreli yazar, yeni rId döner
  addExternal(type, url): string;
  relationships; size; xml();
}
function explicitRelIds(xml, opts?: { prefixes?: string[] }): Set<string>;   // r:* (Transitional+Strict), VML o:relid, öneksiz relId
function rewriteRelIds(xml, map: Map<string,string>, opts?): string;          // tek geçiş, takas güvenli
function parseRels(xml, source): Relationship[]; function relsXml(list): string;
function resolveTarget(source, target): string | null;   // kökün dışı, boş, şemalı (http:, file:) → null
function relativeTarget(source, part): string; function relsPathOf(part): string; function partOfRels(path): string | null;
```

### Kaynak kaydı

```ts
interface PackageSource { id: string; pkg: OpcPackage; name?: string; at: number }
function registerSource(pkg, opts?: { id?: string; name?: string }): string;   // "kaynak-N" ya da verilen id (ör. SHA-256)
function getSource(id): OpcPackage | undefined;
function releaseSource(id): boolean;                                           // belge kapanınca
function openSource(bytes, opts?): Promise<PackageSource>;                     // open + register
function sha256Hex(bytes): Promise<string>;                                    // IndexedDB "kaynaklar" anahtarı
class SourceRegistry { … }  const packageSources: SourceRegistry;              // Worker/test kendi örneğini kurabilir
```

Model **yalnız kimliği** tutar (`structuredClone` ve geri al yığını değişmez). Kalıcılık (taslak/kurtarma,
IndexedDB) uygulamanın işidir: `pkg.bytes` saklanır, kurtarmada `openSource(bytes, { id })` ile aynı
kimlik yeniden kaydedilir. Kaynak bulunamazsa `carryParts` hata vermez, `report.sourceMissing = true`.

### Paket düzeyi taşıma

```ts
function carryParts(opts: {
  app: "belge" | "tablo" | "sunum";
  source: OpcPackage | string | null | undefined;  // paket ya da kayıt kimliği
  output: Uint8Array | ArrayBuffer | OpcPackage;    // yazıcının çıktısı
  macro?: boolean;                                  // hedef .docm/.xlsm/.pptm mi
  partMap?: Iterable<[string, string]> | Record<string, string>;  // kaynak parça → çıktı parçası (sayfa, slayt…)
  carry?: Carry;                                    // öğe düzeyi plan
  overrides?: { own?: string[]; carry?: string[]; drop?: string[] };  // kısa ad ya da tam tür
  ownedCustomProperty?: (name: string) => boolean;  // varsayılan /^OfficeHere/i
  registry?: SourceRegistry;
}): Promise<{ bytes: Uint8Array; report: CarryReport }>;

interface CarryReport {
  carried: { from; to; type; kind: "implicit" | "closure" | "element" | "external" }[];
  dropped: { part; type; reason: DropReason; source }[];
  owned: RelInfo[];      // örtük ama yazıcının (taşınmadı)
  explicit: RelInfo[];   // açık: modelin / öğe düzeyinin işi
  customProperties: number; macro: boolean; sourceMissing: boolean; collected: number; ms: number;
}
type DropReason = "signature" | "macro" | "comments" | "stale" | "thumbnail" | "missing" | "unreadable" | "structural" | "conflict" | "policy";
```

Kurallar (her biri birim testli, `test/opc/carry.test.ts`):

1. **Açık / örtük.** İlişki kimliği kaynak parçanın XML'inde geçiyorsa (`r:id`, `r:embed`, `r:dm`, VML
   `o:relid`, `relId`) açıktır → `report.explicit`, dokunulmaz. Geçmiyorsa örtüktür → politika.
   Kaynak parça okunamazsa (sınır, bozuk) hepsi açık sayılır (taşımamak güvenli olan).
2. **Politika** (`classifyRel`): `drop` > `own` > `carry`. Sıra: `overrides` → tam tür (Sunum yeni yorumları)
   → kısa ad. Tablo aşağıda.
3. **Yazıcı aynı türü ürettiyse** (aynı eşlenen parçada aynı kısa tür) taşınmaz. `customXml` kök ad alanına
   göre eşlenir: aynı ad alanlı öğe yazıcıda varsa yazıcınınki geçerli.
4. **Kapanış.** Taşınan parçanın `.rels`'i kimlikleri korunarak kopyalanır; iç hedefler de taşınır (döngü
   korumalı). Hedef eşlenen parçaysa (kök, ana parça, `partMap`) yeni eşdeğerine bağlanır; tekil sahip tür
   (styles, settings, theme…) çıktıda aynı ad ve türle varsa ona bağlanır; eşlemesiz yapısal parçaya (slayt,
   sayfa, asıl slayt, ana belge…) başvuru atılır (`structural`); hedefi olmayan ilişki atılır (`missing`).
5. **Adlar.** Paket düzeyinde ad korunur, çakışırsa `dizin/oh_tN_ad`. Öğe düzeyinde önek her zaman
   kullanılır; kaynakta zaten `oh_tK_` varsa önek birikmez.
6. **[Content_Types].xml** her zaman ilk parça. Kaynağın Override'ı taşınır; kaynak uzantı varsayılanı
   çıktıda yoksa Default eklenir, çakışıyorsa Override yazılır; türü olmayan parça `application/octet-stream`.
   `macro: true` iken ana parça türü makrolu karşılığına çevrilir (`report.macro`).
7. **custom.xml.** Yazıcının özellikleri korunur, kaynakta olup yazıcıda olmayan kullanıcı özellikleri eklenir
   (`OfficeHere*` adları hariç), `pid` 2'den yeniden numaralanır. Yazıcı custom.xml yazmadıysa yeni parça kurulur.
8. **Çöp toplama.** Yalnız eşlenen parçalardan ulaşılanlar taşınır (silinen slaytın parçaları kendiliğinden
   gitmez). Öğe düzeyi kopyalar yeni paketin kökünden ulaşılamıyorsa atılır (`report.collected`).
9. **Ham kopya.** Taşınan ve yazıcının değişmeyen parçaları yeniden sıkıştırılmaz; yalnız değişen `.rels`,
   içerik türleri, custom.xml ve yamalanan parçalar yeniden yazılır.

Politika tablosu (`policy.ts`):

| | Belge | Tablo | Sunum |
|---|---|---|---|
| **own** | styles, settings, numbering, footnotes, endnotes, comments, header, footer, hyperlink, image, chart, core/extended-properties | styles, sharedStrings, worksheet, chartsheet, theme, calcChain, comments, drawing, vmlDrawing, table, pivot*, slicer*, sheetMetadata, externalLink, image, hyperlink, core/extended | slide, slideMaster, slideLayout, notesMaster, notesSlide, handoutMaster, theme, presProps, viewProps, tableStyles, commentAuthors, comments, image, chart, hyperlink, core/extended |
| **drop** | imza (origin/signature/certificate) → `signature`; thumbnail → `thumbnail`; stylesWithEffects → `stale`; commentsExtended/Ids/Extensible, people, threadedComment, person, Sunum 2018 yorumları → `comments`; makrosuz hedefte vbaProject, wordVbaData, keyMapCustomizations, attachedToolbars → `macro` | | |
| **carry** | geri kalan her şey (theme, fontTable, webSettings, glossary, customXml, customUI, vbaProject (makrolu), xmlMaps, connections, bilinmeyenler…) | | |

`diagramDrawing` her uygulamada `own`: SmartArt çizimi öğe düzeyinin (`Carry`) işidir.

### Öğe düzeyi taşıma

```ts
class Carry {
  constructor(registry?: SourceRegistry);
  instance(source: string | OpcPackage): CarryInstance;   // her opak öğe örneği için bir tane
  patch(newPath: string, fn: (xml: string) => string): void;  // kopyalanan parçayı yazmadan önce dönüştür
  readonly size: number;
}
interface CarryInstance {
  readonly source: OpcPackage | undefined;
  copy(absoluteSourcePart: string): string | null;   // yeni ad ("…/oh_tN_ad"); parça yoksa null
  mapped(absoluteSourcePart: string): string | undefined;
}
interface RawRef { id: string; type: string; target: string; external: boolean }  // target: kaynakta mutlak ad ya da adres
function collectRefs(pkg, part, xml, opts?): Promise<RawRef[]>;            // okurken: opak XML'in ilişkileri
function reattach(xml, refs, inst, addRel: (r: { type; target; external }) => string, opts?)
  : { xml: string; ids: Map<string, string>; missing: RawRef[] };          // yazarken: kopyala + yeni rId + XML'i yeniden yaz
```

Kopyaların baytları `carryParts({ carry })` sırasında kaynak paketten ham kopyalanır, kapanışları
(ör. SmartArt çiziminin resimleri) ve içerik türleriyle. Belge genelinde benzersiz olması gereken XML
kimlikleri (`wp:docPr/@id`, `cNvPr/@id`, `w14:paraId`, yer işaretleri, SmartArt `modelId`) uygulamanın işidir.

### XmlPart

`new XmlPart(xml, ORDER)`: `get/all/has/set/add/map/ensureNamespace/toString`, `upsertExt/getExt`,
`childElements/rebuildElement/attrOf/topLevelNames/orderViolation`. Sahip parçalar (Belge `settings.xml`,
Sunum `presentation.xml`, Tablo `workbook.xml`/sayfalar) baştan yazılmaz; özgün parça okunur, yalnız
modelin öğeleri değiştirilir, bilinmeyenler (compat, rsids, docVars, mathPr, w14/w15, extLst) yerinde kalır.

## Hata davranışı

- Tek hata türü `OpcError` (`code`: `not-zip` | `corrupt` | `too-large` | `unsupported`).
- `OpcPackage.open`: ZIP değilse `not-zip`/`corrupt` fırlatır. Bozuk ya da eksik `[Content_Types].xml`
  ve `.rels` hata değildir (boş sayılır); okunamayan ZIP girdisi atlanır (`skippedEntries`).
- `data/text`: açık boyut `maxPartSize`'ı (varsayılan 256 MB) aşarsa `too-large` — bildirilen boyut yalan
  olsa da açma sınırda kesilir (sıkıştırma bombası). `.rels` ve içerik türleri için sınır 16 MB.
- `carryParts`: **kaynakla ilgili hiçbir sorun fırlatmaz** (bozuk/döngülü `.rels`, mutlak ya da kökün
  dışına çıkan hedef, ayrılmış ad, okunamayan parça, çok büyük parça); rapora yazar. Yalnız `output`
  ZIP değilse `OpcError` fırlatır (yazıcı hatası). Bulanık testli (`test/opc/bulanik.test.ts`).
- Taşınan parça açılmadan ham kopyalanır: büyük medya ve vbaProject CPU harcamaz.

## Örnekler

**Belge (P3.1)** — okuma ve yazma:

```ts
// reader.ts
const kaynak = registerSource(await OpcPackage.open(bytes), { name: file.name });
doc.kaynak = kaynak;                                   // model yalnız kimliği taşır
if (getSource(kaynak)!.signed) warnings.push("Kaydederseniz dijital imza kaldırılır.");

// writer.ts (docx/docm hedefi)
const out = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
if (!doc.kaynak) return out;
const { bytes, report } = await carryParts({ app: "belge", source: doc.kaynak, output: out, macro: ext === "docm" });
if (report.dropped.some((d) => d.reason === "macro")) warnings.push("Makrolar .docx biçiminde saklanamaz.");
if (report.dropped.some((d) => d.reason === "comments")) warnings.push("Yorum yanıt zinciri kaldırıldı.");
return bytes;

// settings.xml yaması
const s = new XmlPart((await getSource(doc.kaynak)!.text("word/settings.xml")) ?? DEFAULT, SETTINGS_ORDER);
s.set("defaultTabStop", `<w:defaultTabStop w:val="${tab}"/>`);
zip.file("word/settings.xml", s.toString());
```

**Belge opak öğe (P3.2)** — okurken ilişkileri sakla, yazarken yeniden kur:

```ts
const refs = await collectRefs(pkg, "word/document.xml", opaqueXml);     // node.attrs = { xml, refs }
// partWriter: tek Carry bütün yazma turu için
const carry = new Carry();
const { xml } = reattach(node.attrs.xml, node.attrs.refs, carry.instance(doc.kaynak), (r) =>
  r.external ? docRels.addExternal(r.type, r.target) : docRels.addPart(r.type, r.target));
// … writer sonunda: carryParts({ app: "belge", source: doc.kaynak, output, carry })
```

**Tablo (P3.3)** — sayfa eşlemesi ve makro:

```ts
const partMap = wb.sheets.map((s, i) => [s.sourcePath, `xl/worksheets/sheet${i + 1}.xml`] as const).filter(([a]) => a);
const { bytes } = await carryParts({ app: "tablo", source: wb.kaynak, output: await postProcess(...), partMap, macro: ext === "xlsm" });
```

**Sunum (P3.4)** — SmartArt kopyası ve slayt ilişki kimliği yaması:

```ts
const inst = carry.instance(shape.raw.src);
const { xml, ids } = reattach(shape.raw.xml, shape.raw.refs, inst, (r) => slideRels.add(r));
for (const e of shape.raw.extra) {                       // dsp:dataModelExt relId: slaydın ilişkisini anar
  const drawing = inst.copy(e.target)!;
  const newId = slideRels.addPart(e.type, drawing);
  carry.patch(inst.mapped(e.patchIn)!, (s) => s.replace(`relId="${e.id}"`, `relId="${newId}"`));
}
const partMap = slides.map((s) => [s.sourcePath, s.outPath]);   // silinen slayt eşlenmez → parçaları taşınmaz
```

## Ölçüm (P2.1)

POI 28 docx/docm, bugünkü `writeDocx` + `carryParts` (`asama2/deneme/p21/belge_tasima_core.mjs`):
parça kaybı 185 → 72, ilişki kaybı 209 → 97, OPC 0 hata, 28/28 kendi okuyucumuzla açılır ve metin aynı,
2. tur kararlı, LibreOffice 28/28 (PDF sayfa ve sözcük sayısı bugünkü çıktıyla aynı), taşıma ortalaması
3 ms (dosya başına 10 koşunun medyanı; en kötü tek koşu 20 ms). Kalan 72: bilinçli düşürmeler 12
(stylesWithEffects 7, imza 2, küçük resim 2, keyMap 1), yazıcının sahip olduğu boş dipnot/sonnot 21,
üst/alt bilgi varyantları 23 (model), OLE ve önizlemeleri 14 (öğe düzeyi, P3.2), açıklama 1 (model),
`image/jpg` 1 (yazıcı aynı resmi `.jpeg` olarak yazıyor; gerçek kayıp değil).
