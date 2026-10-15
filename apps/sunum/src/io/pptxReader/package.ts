/**
 * PPTX paket erişimi: korumalı parça okuma (eksik, bozuk, çok büyük, kısmen
 * okunabilir), uyarı yazan metin okuma ve paket içinde parça arama. Tüm parça
 * okuyucular pakete bu katman üzerinden erişir.
 */
import { LoadWarnings } from "../loadWarnings";
import { attr, kids, parseXml } from "../xml";
import { openZip, PackageError, PartError, xmlHealth, type SafeZip } from "../zipSafe";

/** PPTX okuma hatası: Türkçe ileti (+ günlük için teknik ayrıntı). */
export class PptxError extends PackageError {}

export type PartProblem = "missing" | "corrupt" | "tooLarge" | "partial";
export interface Part {
  root?: Element;
  xml?: string;
  problem?: PartProblem;
  detail?: string;
}

/** Yol sonundaki sayıya göre sıralama ("slide10" "slide9"dan sonra). */
export const byNumber = (a: string, b: string) => {
  const na = Number(/(\d+)\.xml$/i.exec(a)?.[1] ?? 0);
  const nb = Number(/(\d+)\.xml$/i.exec(b)?.[1] ?? 0);
  return na - nb || a.localeCompare(b);
};

/** Okuma boyunca paylaşılan paket erişimi ve açılış uyarıları. */
export interface PptxPackage {
  zip: SafeZip;
  W: LoadWarnings;
  /** Metin parçası; bozuk ya da sınırı aşan parça atlanır (`label` verilirse uyarı, yoksa ayrıntı yazılır). */
  text(p: string, label?: string): Promise<string | undefined>;
  /** XML parçası + sağlık durumu (eksik, bozuk, çok büyük, kısmen okunabilir). */
  part(p: string | undefined, roots: string | string[]): Promise<Part>;
  /** Paketteki, desene uyan parçalar (sayı sırasıyla). */
  parts(re: RegExp): string[];
}

/** ZIP paketini açar; paket hatası PptxError olarak atılır. */
export async function openPptxZip(data: ArrayBuffer | Uint8Array): Promise<SafeZip> {
  try {
    return await openZip(data, "Dosya geçerli bir PPTX (ZIP) paketi değil.");
  } catch (e) {
    if (e instanceof PackageError) throw new PptxError(e.message, e.detail);
    throw e;
  }
}

export function pptxPackage(zip: SafeZip, W: LoadWarnings): PptxPackage {
  const text = async (p: string, label?: string): Promise<string | undefined> => {
    try {
      return await zip.text(p);
    } catch (e) {
      if (!(e instanceof PartError)) throw e;
      if (label) W.add(label, `${p}: ${e.detail ?? e.message}`);
      else W.detail(`${p}: ${e.detail ?? e.message}`);
      return undefined;
    }
  };
  const part = async (p: string | undefined, roots: string | string[]): Promise<Part> => {
    if (!p) return { problem: "missing", detail: "ilişki hedefi yok" };
    let xml: string | undefined;
    try {
      xml = await zip.text(p);
    } catch (e) {
      if (!(e instanceof PartError)) throw e;
      return { problem: e.kind, detail: `${p}: ${e.detail ?? e.message}` };
    }
    if (xml === undefined) return { problem: "missing", detail: `${p}: pakette yok` };
    if (!xml.trim()) return { problem: "corrupt", detail: `${p}: boş parça` };
    return checkedPart(p, xml, roots);
  };
  const parts = (re: RegExp) => zip.paths.filter((p) => re.test(p)).sort(byNumber);
  return { zip, W, text, part, parts };
}

/**
 * İçerik türleri (parça yolu → tür): bozuk ilişkilerde parçaların türünü doğrulamak için.
 * Tanınmayan parçaların taşınması kaydederken core `carryParts` ile yapılır (türler kaynak paketten).
 */
export async function readContentTypes(pk: PptxPackage): Promise<Map<string, string>> {
  const contentTypes = new Map<string, string>();
  const ctXml = await pk.text("[Content_Types].xml");
  if (ctXml) for (const o of kids(parseXml(ctXml).documentElement, "Override")) contentTypes.set((attr(o, "PartName") ?? "").replace(/^\//, ""), attr(o, "ContentType") ?? "");
  else pk.W.add("Paketin içerik türleri bölümü ([Content_Types].xml) eksik ya da bozuk; dosya onarılarak açıldı.");
  return contentTypes;
}

/** Ayrıştırılan parçanın kök öğesi beklenen mi, XML tamamen okunabildi mi? */
function checkedPart(p: string, xml: string, roots: string | string[]): Part {
  const doc = parseXml(xml);
  const health = xmlHealth(doc, roots);
  if (health === "bad") {
    const r = doc.documentElement?.localName;
    return { problem: "corrupt", detail: `${p}: ${!r || r === "parsererror" ? "XML ayrıştırılamadı" : `beklenmeyen kök öğe <${r}>`}` };
  }
  return health === "partial" ? { root: doc.documentElement, xml, problem: "partial", detail: `${p}: XML yalnız kısmen ayrıştırılabildi` } : { root: doc.documentElement, xml };
}
