// Paket düzeyi taşıma politikası: örtük ilişkiyle bağlı bir parça için ne yapılacağı.
//
//   own   (SAHİP)  Model kaynak alınır; yazıcı üretmediyse taşınmaz (ör. dipnot yok → dipnot parçası yok).
//   drop  (DÜŞÜR)  Taşınırsa geçersiz ya da eskimiş kalır (imza, yorum genişletmeleri, küçük resim…);
//                  rapora gerekçesiyle yazılır, uygulama kullanıcıyı uyarır.
//   carry (TAŞI)   Geri kalan her şey, bilinmeyenler dahil (ileriye dönük güvenli varsayılan).
//
// Kurallar önce tam ilişki türü adresine, sonra adresin son bölümüne (kısa ad: "theme", "customXml"…)
// bakar; böylece Strict (purl.oclc.org) türleri de aynı kısa adla eşleşir.

export type OpcApp = "belge" | "tablo" | "sunum";
export type PolicyAction = "own" | "carry" | "drop";

/**
 * Düşürme gerekçesi:
 * - signature: dijital imza (kaydedince geçersizleşir)
 * - macro: makro projesi, hedef biçim makrolu değil
 * - comments: yorum yanıt zinciri / kişi listesi (model bunları taşımıyor)
 * - stale: eskimiş kopya (stylesWithEffects)
 * - thumbnail: küçük resim (eskir)
 * - missing: ilişkinin hedefi kaynakta yok
 * - unreadable: parça okunamadı (bozuk ZIP girdisi)
 * - structural: taşınan parçadan modelin yapısal parçasına (slayt, sayfa, üst bilgi…) eşlemesiz başvuru
 * - conflict: öğe düzeyi kopyanın adı yazıcının bir parçasıyla çakıştı (yazıcınınki geçerli)
 * - policy: uygulamanın verdiği ek kural (PolicyOverrides.drop)
 */
export type DropReason = "signature" | "macro" | "comments" | "stale" | "thumbnail" | "missing" | "unreadable" | "structural" | "conflict" | "policy";

export interface PolicyDecision {
  action: PolicyAction;
  reason?: DropReason;
}

/** Uygulamanın tabloyu genişletmesi: kısa ad ("theme") ya da tam tür adresi. Sıra: drop > own > carry. */
export interface PolicyOverrides {
  own?: readonly string[];
  carry?: readonly string[];
  drop?: readonly string[];
}

const OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const PK = "http://schemas.openxmlformats.org/package/2006/relationships/";
const MS = "http://schemas.microsoft.com/office/";

/** Sık kullanılan ilişki türleri (Transitional). */
export const REL_TYPES = {
  officeDocument: `${OD}officeDocument`,
  coreProperties: `${PK}metadata/core-properties`,
  extendedProperties: `${OD}extended-properties`,
  customProperties: `${OD}custom-properties`,
  thumbnail: `${PK}metadata/thumbnail`,
  signatureOrigin: `${PK}digital-signature/origin`,
  signature: `${PK}digital-signature/signature`,
  customXml: `${OD}customXml`,
  customXmlProps: `${OD}customXmlProps`,
  theme: `${OD}theme`,
  styles: `${OD}styles`,
  settings: `${OD}settings`,
  fontTable: `${OD}fontTable`,
  webSettings: `${OD}webSettings`,
  numbering: `${OD}numbering`,
  glossaryDocument: `${OD}glossaryDocument`,
  image: `${OD}image`,
  hyperlink: `${OD}hyperlink`,
  vbaProject: `${MS}2006/relationships/vbaProject`,
  wordVbaData: `${MS}2006/relationships/wordVbaData`,
  keyMapCustomizations: `${MS}2006/relationships/keyMapCustomizations`,
  stylesWithEffects: `${MS}2007/relationships/stylesWithEffects`,
  commentsExtended: `${MS}2011/relationships/commentsExtended`,
  people: `${MS}2011/relationships/people`,
  commentsIds: `${MS}2016/09/relationships/commentsIds`,
  commentsExtensible: `${MS}2018/08/relationships/commentsExtensible`,
  diagramDrawing: `${MS}2007/relationships/diagramDrawing`,
} as const;

/** İlişki türünün kısa adı: adresin son bölümü ("…/relationships/theme" → "theme"). */
export function relTypeName(type: string): string {
  const t = type.trim();
  const i = t.lastIndexOf("/");
  return i >= 0 ? t.slice(i + 1) : t;
}

/** Makrolu biçim dışında taşınmayan türler. */
export const MACRO_TYPES: ReadonlySet<string> = new Set(["vbaProject", "vbaData", "wordVbaData", "keyMapCustomizations", "attachedToolbars", "vbaProjectSignature", "vbaProjectSignatureAgile", "vbaProjectSignatureV3"]);

/** Her uygulamada düşürülen türler (kısa ad → gerekçe). */
export const DROP_TYPES: ReadonlyMap<string, DropReason> = new Map<string, DropReason>([
  ["origin", "signature"],
  ["signature", "signature"],
  ["certificate", "signature"],
  ["thumbnail", "thumbnail"],
  ["stylesWithEffects", "stale"],
  ["commentsExtended", "comments"],
  ["commentsIds", "comments"],
  ["commentsExtensible", "comments"],
  ["people", "comments"],
  ["threadedComment", "comments"],
  ["person", "comments"],
]);

/** Tam adresle düşürülen türler (kısa adı başka bir türle aynı olanlar: Sunum'un yeni yorumları). */
const DROP_FULL: ReadonlyMap<string, DropReason> = new Map<string, DropReason>([
  [`${MS}2018/10/relationships/comments`, "comments"],
  [`${MS}2018/10/relationships/authors`, "comments"],
]);

/** Uygulama başına yazıcının sahip olduğu türler (kısa ad). */
export const OWN_TYPES: Readonly<Record<OpcApp, ReadonlySet<string>>> = {
  belge: new Set(["officeDocument", "styles", "settings", "numbering", "footnotes", "endnotes", "comments", "header", "footer", "hyperlink", "image", "chart", "core-properties", "extended-properties", "diagramDrawing"]),
  tablo: new Set([
    "officeDocument", "styles", "sharedStrings", "worksheet", "chartsheet", "dialogsheet", "theme", "calcChain", "comments", "drawing", "vmlDrawing", "table", "pivotTable",
    "pivotCacheDefinition", "pivotCacheRecords", "slicer", "slicerCache", "sheetMetadata", "core-properties", "extended-properties", "externalLink", "image", "hyperlink", "diagramDrawing",
  ]),
  sunum: new Set([
    "officeDocument", "slide", "slideMaster", "slideLayout", "notesMaster", "notesSlide", "handoutMaster", "theme", "presProps", "viewProps", "tableStyles", "commentAuthors", "comments",
    "core-properties", "extended-properties", "image", "chart", "hyperlink", "diagramDrawing",
  ]),
};

/**
 * Modelin yapısını oluşturan türler. Taşınan bir parça bunlardan birine başvuruyorsa ve uygulama
 * eşleme vermediyse ilişki atılır (yapısal parçanın eskimiş kopyası pakete sokulmaz).
 */
export const STRUCTURAL_TYPES: ReadonlySet<string> = new Set(["officeDocument", "slide", "slideMaster", "slideLayout", "notesMaster", "handoutMaster", "worksheet", "chartsheet", "dialogsheet"]);

/** Pakette tek olan sahip türler: taşınan parça bunlara başvurursa yazıcının aynı adlı parçasına bağlanır. */
export const SINGLETON_TYPES: ReadonlySet<string> = new Set([
  "styles", "settings", "numbering", "fontTable", "webSettings", "theme", "sharedStrings", "calcChain", "presProps", "viewProps", "tableStyles", "commentAuthors", "core-properties",
  "extended-properties",
]);

function matches(list: readonly string[] | undefined, type: string, short: string): boolean {
  return !!list && (list.includes(type) || list.includes(short));
}

/**
 * Örtük ilişkinin politikası. `macro`: hedef biçim makrolu mu (docm/xlsm/pptm).
 * customXml ve custom-properties burada "carry" döner; kök ad alanı eşlemesi ve özellik birleştirme
 * taşıma sırasında yapılır.
 */
export function classifyRel(app: OpcApp, type: string, opts: { macro?: boolean; overrides?: PolicyOverrides } = {}): PolicyDecision {
  const short = relTypeName(type);
  const o = opts.overrides;
  if (matches(o?.drop, type, short)) return { action: "drop", reason: "policy" };
  if (matches(o?.own, type, short)) return { action: "own" };
  if (matches(o?.carry, type, short)) return { action: "carry" };
  const full = DROP_FULL.get(type);
  if (full) return { action: "drop", reason: full };
  const drop = DROP_TYPES.get(short);
  if (drop) return { action: "drop", reason: drop };
  if (MACRO_TYPES.has(short)) return opts.macro ? { action: "carry" } : { action: "drop", reason: "macro" };
  if (OWN_TYPES[app].has(short)) return { action: "own" };
  return { action: "carry" };
}

const MACRO_CT: [string, string][] = [
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "application/vnd.ms-word.document.macroEnabled.main+xml"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml"],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml", "application/vnd.ms-excel.template.macroEnabled.main+xml"],
  ["application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"],
  ["application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml", "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"],
  ["application/vnd.openxmlformats-officedocument.presentationml.template.main+xml", "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"],
];

/** Ana parça içerik türünün makrolu karşılığı (zaten makroluysa ya da bilinmiyorsa aynısı). */
export function macroContentType(ct: string): string {
  const k = ct.trim().toLowerCase(); // yerel-ayar: ascii
  return MACRO_CT.find(([plain]) => plain === k)?.[1] ?? ct;
}

/** Ana parça içerik türünün makrosuz karşılığı. */
export function plainContentType(ct: string): string {
  const k = ct.trim().toLowerCase(); // yerel-ayar: ascii
  return MACRO_CT.find(([, macro]) => macro.toLowerCase() === k)?.[0] ?? ct; // yerel-ayar: ascii
}

/** İçerik türü makrolu bir ana parça mı? */
export function isMacroContentType(ct: string | undefined): boolean {
  return !!ct && /\.macroenabled/i.test(ct);
}
