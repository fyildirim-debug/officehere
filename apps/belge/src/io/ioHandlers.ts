// Belge dosya G/Ç işleyicileri: Worker'da (io.worker.ts) ya da Worker olmayan ortamda (testler, Node betikleri)
// süreç içinde aynı RPC ile çalışır. Açılan .docx/.odt'nin özgün paketi (geçiş katmanı) bu iş parçacığının kaynak
// kaydına girer; yazma da burada yapıldığından taşınacak parçalar oradan okunur. Kayıt yoksa (Worker yeniden
// kuruldu, kurtarılan taslak) yazma, eksik kaynakları ister; ana iş parçacığı baytları gönderir.
import { withTransfer, type IoContext } from "@officehere/core";
import { getSource, openSource, releaseSource } from "@officehere/core/opc";
import type { BelgeDocument } from "../model/document";
import { DomRequired, readDocument, type FormatId } from "./index";
import { readDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odt/writer";
import { writeRtf } from "./rtf/writer";
import { writeMarkdown } from "./markdown/writer";
import { LoadWarnings, ODF_OBJECTS_FALLBACK, odtSaveNotes, saveNotes, type PlainWarnings } from "./warnings";
import { hasOdfObjects } from "../editor/opaque";

/** Açma sonucu: okunan belge ya da "ana iş parçacığında aç" (HTML, Markdown, .docx yedek dönüştürücüsü). */
export type OpenReply = { main: true } | { main?: false; doc: BelgeDocument; format: FormatId; notes: string[]; warnings: PlainWarnings };

/** Yazma sonucu: baytlar ve kaydetme bildirimleri ya da bu iş parçacığında bulunmayan kaynak paketler. */
export type WriteReply = { bytes: Uint8Array; notes: string[]; missing?: undefined } | { missing: string[] };

/** Belgenin andığı kaynak paketler: ayarlardaki kimlik ve opak öğelerin (başka belgeden yapıştırılmış olabilir) kaynağı. */
export function sourceIds(doc: BelgeDocument): Set<string> {
  const ids = new Set<string>();
  if (doc.settings.kaynak) ids.add(doc.settings.kaynak);
  const walk = (v: unknown) => {
    if (Array.isArray(v)) {
      for (const x of v) walk(x);
      return;
    }
    if (!v || typeof v !== "object") return;
    const n = v as { type?: unknown; attrs?: { src?: unknown } };
    if (n.type === "opaqueBlock" || n.type === "opaqueInline") {
      if (typeof n.attrs?.src === "string" && n.attrs.src) ids.add(n.attrs.src);
      return;
    }
    for (const x of Object.values(v)) if (x && typeof x === "object") walk(x);
  };
  walk(doc.content);
  walk(doc.settings);
  return ids;
}

/** Otomatik kaydedilen .docx yeniden açılabiliyor mu (kaynak kaydedilmez; desteklenmeyen öğe uyarıları sorun değil). */
export async function canReopenDocxBytes(bytes: ArrayBuffer): Promise<boolean> {
  const warnings = new LoadWarnings();
  const doc = await readDocx(bytes, warnings, { kaynak: false });
  return !!doc?.content && !(["corrupt", "unclosed", "simplified", "flattened"] as const).some((k) => warnings.has(k));
}

export const handlers = {
  /** Bayt → belge modeli (DOM gerektiren biçimlerde `{ main: true }`). */
  async open({ bytes, name }: { bytes: ArrayBuffer; name: string }, ctx: IoContext): Promise<OpenReply> {
    ctx.progress(0, "Dosya okunuyor");
    try {
      const r = await readDocument(name, bytes, { noDom: true });
      ctx.progress(1, "Belge hazırlanıyor");
      return { doc: r.doc, format: r.format, notes: r.notes, warnings: r.warnings.toPlain() };
    } catch (err) {
      if (err instanceof DomRequired) return { main: true };
      throw err;
    }
  },

  /**
   * Belge modeli → .docx (ya da `format: "odt"` ile .odt) baytları. `sources`: ana iş parçacığının gönderdiği kaynak
   * paketler (null: onda da yok, geçiş katmanı olmadan yazılır). Bu iş parçacığında olmayan ve gönderilmeyen kaynak
   * varsa `{ missing }`.
   */
  async write({ doc, macro, sources, format }: { doc: BelgeDocument; macro?: boolean; sources?: Record<string, Uint8Array | null>; format?: "docx" | "odt" }) {
    const absent = new Set<string>();
    for (const [id, bytes] of Object.entries(sources ?? {})) {
      if (getSource(id)) continue;
      if (!bytes) absent.add(id);
      else await openSource(bytes, { id }).catch(() => absent.add(id));
    }
    const missing = [...sourceIds(doc)].filter((id) => !getSource(id) && !absent.has(id));
    if (missing.length) return { missing } as WriteReply;
    const notes: string[] = [];
    let bytes: Uint8Array;
    if (format === "odt") bytes = await writeOdt(doc, { onReport: (r) => notes.push(...odtSaveNotes(r)) });
    else {
      bytes = await writeDocx(doc, { macro, onReport: (r) => notes.push(...saveNotes(r)) });
      // .odt'den gelen gömülü nesneler .docx'te yalnız yedek görünümleriyle yazılır.
      if (hasOdfObjects(doc.content) || hasOdfObjects(doc.settings)) notes.push(ODF_OBJECTS_FALLBACK);
    }
    return withTransfer<WriteReply>({ bytes, notes }, [bytes.buffer]);
  },

  /** .rtf ve .md (DOM gerektirmeyen metin yazıcıları; .odt `write` ile, kaynak paketlerle yazılır). */
  async encode({ format, doc }: { format: "rtf" | "md"; doc: BelgeDocument }) {
    return format === "rtf" ? writeRtf(doc) : writeMarkdown(doc);
  },

  /** Otomatik kayıt doğrulaması (ayrı Worker'da çağrılır). */
  verify({ bytes }: { bytes: ArrayBuffer }): Promise<boolean> {
    return canReopenDocxBytes(bytes);
  },

  /** Belge kapandı: özgün paketi bırak. */
  release({ id }: { id: string }): boolean {
    return releaseSource(id);
  },
};

export type IoHandlersType = typeof handlers;
