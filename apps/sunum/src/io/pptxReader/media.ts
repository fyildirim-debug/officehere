/**
 * Ortam dosyaları (resim, video, ses → veri adresi) ve grafik parçaları
 * (önbellekli verileriyle); şekiller okunmadan önce bir kez okunur.
 */
import type { ChartData, SchemeColorName } from "../../model/types";
import { EXT_MIME, MEDIA_MIME } from "../pptxWriter";
import { readChartXml } from "../readChart";
import { parseRels, relsPathFor, type Rel } from "../xml";
import { PartError } from "../zipSafe";
import type { PptxPackage } from "./package";

const MEDIA_RE = /\.(png|jpe?g|gif|bmp|svg|webp|tiff?|emf|wmf|mp4|m4v|webm|ogv|mov|mp3|m4a|wav|ogg|oga|aac|flac)$/i;

function bytesToBase64(bytes: Uint8Array): string {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(s);
}

/** Paketteki ortam dosyaları (yol → veri adresi) ve bozuk/çok büyük olanlar. */
export async function readMedia(pk: PptxPackage): Promise<{ media: Map<string, string>; badMedia: Set<string> }> {
  const { zip, W } = pk;
  const media = new Map<string, string>();
  const badMedia = new Set<string>();
  const mediaFiles = zip.paths.filter((p) => MEDIA_RE.test(p));
  await Promise.all(
    mediaFiles.map(async (p) => {
      try {
        const bytes = await zip.bytes(p);
        if (!bytes) return;
        const ext = p.split(".").pop()!.toLowerCase(); // yerel-ayar: ascii
        media.set(p, `data:${EXT_MIME[ext] ?? MEDIA_MIME[ext] ?? "application/octet-stream"};base64,${bytesToBase64(bytes)}`);
      } catch (e) {
        if (!(e instanceof PartError)) throw e;
        badMedia.add(p);
        W.detail(`${p}: ${e.detail ?? e.message}`);
      }
    }),
  );
  if (badMedia.size) W.add(`${badMedia.size} resim ya da ortam dosyası bozuk ya da çok büyük olduğu için yüklenemedi; yerlerinde boş çerçeve görünür.`);
  return { media, badMedia };
}

const DIAGRAM_CT = /drawingml\.diagram(Data|Drawing)\+xml$/i;

/**
 * SmartArt veri ve çizim parçaları (yol → XML) ve çizimlerin ilişkileri: SmartArt yedek
 * görünümü (çizimdeki şekiller) ve dsp:dataModelExt başvurusu için.
 */
export async function readDiagrams(pk: PptxPackage, contentTypes: Map<string, string>): Promise<{ diagrams: Map<string, string>; diagramRels: Map<string, Map<string, Rel>> }> {
  const diagrams = new Map<string, string>();
  const diagramRels = new Map<string, Map<string, Rel>>();
  const paths = pk.zip.paths.filter((p) => DIAGRAM_CT.test(contentTypes.get(p) ?? "") || /^ppt\/diagrams\/(data|drawing)[^/]*\.xml$/i.test(p));
  for (const p of paths) {
    const xml = await pk.text(p);
    if (!xml) continue;
    diagrams.set(p, xml);
    diagramRels.set(p, parseRels(await pk.text(relsPathFor(p)), p));
  }
  return { diagrams, diagramRels };
}

/** Grafik parçaları (yol → veri); okunamayanlar yedek resimleriyle (varsa) gösterilir. */
export async function readCharts(pk: PptxPackage, media: Map<string, string>, clrMap: Record<string, SchemeColorName>, lnWidths: number[]): Promise<Map<string, ChartData>> {
  const { zip, W } = pk;
  const charts = new Map<string, ChartData>();
  let badCharts = 0;
  for (const cp of zip.paths.filter((p) => /^ppt\/charts\/[^/]+\.xml$/i.test(p))) {
    const cx = await pk.text(cp);
    if (!cx) {
      badCharts++;
      continue;
    }
    try {
      const data = readChartXml(cx, { rels: new Map(), media, clrMap, idMap: new Map(), lnWidths });
      if (data) charts.set(cp, data);
    } catch (e) {
      // Okunamayan grafik yedek resmiyle (varsa) gösterilir.
      badCharts++;
      W.detail(`${cp}: ${(e as Error)?.message ?? e}`);
    }
  }
  if (badCharts) W.add("Bazı grafikler okunamadı; varsa yedek resimleri gösteriliyor.");
  return charts;
}
