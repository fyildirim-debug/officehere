/**
 * Tablo dosya G/Ç istemcileri. Ayrıştırma ve yazma Worker'da; ana iş parçacığı yalnız kitabı ara
 * gösterimden (IR) parça parça kurar ya da yazmak için IR anlık görüntüsünü parça parça çıkarır,
 * böylece büyük kitapta uzun görev oluşmaz. Açma, yazma ve otomatik kayıt doğrulaması ayrı
 * istemcilerdedir: açmanın iptali (Worker'ı sonlandırır) süren bir kaydı ya da doğrulamayı kesmez.
 *
 * Kaynak paket (.xlsx'in modelin tanımadığı parçaları) ana iş parçacığında kaydedilir; yazarken
 * baytları IR'deki kimlikle Worker'a gider. Worker kurulamazsa (eski ortam) işlem süreç içinde yapılır.
 */
import { createIoWorker, IoWorkerError, workerOrInProcess, type IoProgress } from "@officehere/core";
import { getSource } from "@officehere/core/opc";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { bookFromIr, bookToIr, irTransfer, runSliced, runSync, yieldToMain, type BookIr } from "./bookIr";
import { FormatError } from "./bookUtil";
import { decodeText, detectDelimiter } from "./csv";
import { openBook, writeBook, type OpenedBook, type SaveFormat } from "./formats";
import type { IoHandlers, ReadBookResult, ReadOnMain } from "./ioHandlers";
import { adoptSource, type XlsxWriteReport } from "./xlsx";

function ioClient() {
  return createIoWorker<IoHandlers>(
    workerOrInProcess(
      () => new Worker(new URL("./io.worker.ts", import.meta.url), { type: "module" }),
      () => import("./ioHandlers").then((m) => m.handlers),
    ),
    { errors: { FormatError } },
  );
}

/** Açma (iptal edilebilir); iş bitince Worker kapatılır, belleği hemen geri verilir. */
const readIo = ioClient();
/** Kaydetme ve dışa aktarma (iptal edilmez; kaynak paketi önbellekte tutar). */
const writeIo = ioClient();
/** Otomatik kayıt doğrulaması; iş bitince kapatılır. */
const verifyIo = ioClient();

export interface IoOptions {
  signal?: AbortSignal;
  onProgress?: (p: IoProgress) => void;
}

/** Açılışın Worker'daki payı; kalanı kitabın ana iş parçacığında kurulmasıdır. */
const READ_SHARE = 0.85;

/** Worker hiç kurulamadı mı (yalnız bu durumda süreç içi yola düşülür). */
function workerUnavailable(e: unknown): boolean {
  return e instanceof IoWorkerError && e.code === "load";
}

function textDelimiter(o: OpenedBook, data: ArrayBuffer): string | undefined {
  return o.format === "csv" || o.format === "txt" ? detectDelimiter(decodeText(data)) : o.format === "tsv" ? "\t" : undefined;
}

export type OpenedBookIo = OpenedBook & { delimiter?: string };

/** Süreç içi açma (HTML: Worker'da DOMParser yok; Worker kurulamazsa). */
async function openOnMain(name: string, data: ArrayBuffer): Promise<OpenedBookIo> {
  const o = await openBook(name, data);
  return { ...o, delimiter: textDelimiter(o, data) };
}

/**
 * Dosyayı açar: Worker okur ve IR'yi aktarır, kitap ana iş parçacığında dilimler hâlinde kurulur.
 * İptalde AbortError fırlatır (isAbortError). Baytlar kopyalanarak gönderilir; çağıranınki bozulmaz.
 */
export async function openBookIo(name: string, data: ArrayBuffer, opts: IoOptions = {}): Promise<OpenedBookIo> {
  const { signal, onProgress } = opts;
  let r: ReadBookResult | ReadOnMain;
  try {
    r = await readIo.call("read", { bytes: data, name }, { signal, onProgress: (p) => onProgress?.({ ratio: p.ratio * READ_SHARE, step: p.step }) });
  } catch (e) {
    if (!workerUnavailable(e)) throw e;
    console.warn("[tablo] dosya G/Ç Worker'ı kurulamadı; dosya ana iş parçacığında açılıyor", e);
    return openOnMain(name, data);
  } finally {
    if (!readIo.pending) readIo.terminate();
  }
  if (r.kind === "main") return openOnMain(name, data);
  const wb = await runSliced(bookFromIr(r.ir), { signal, onProgress: (x) => onProgress?.({ ratio: READ_SHARE + (1 - READ_SHARE) * x, step: "Çalışma kitabı kuruluyor" }) });
  // Hesapla kurulan hücreler (taşan diziler, döngüler) ve etkin sayfanın kullanılan alanı ayrı
  // dilimlerde: yüklemeden sonraki ilk çizim büyük sayfayı yeniden taramasın
  wb.syncTables();
  wb.engine.recalc();
  await yieldToMain();
  wb.usedRange(wb.active);
  await yieldToMain();
  if (signal?.aborted) throw signal.reason;
  if (r.source) await adoptSource(wb, r.source);
  return { wb, format: r.format, template: r.template, notice: r.notice, delimiter: r.delimiter };
}

/**
 * Kitabın IR anlık görüntüsü: hesap güncellenir, hücreler dilimler hâlinde okunur. Bu arada kitap
 * değişirse (motor girdisi ya da `version`) görüntü baştan alınır; sürekli değişiyorsa tek parçada.
 */
export async function snapshotIr(wb: WorkbookModel, version?: () => unknown): Promise<BookIr> {
  for (let attempt = 0; ; attempt++) {
    wb.engine.recalc();
    const iv = wb.engine.inputVersion;
    const v0 = version?.();
    if (attempt >= 3) return runSync(bookToIr(wb))!;
    const ir = await runSliced(bookToIr(wb, { stale: () => wb.engine.inputVersion !== iv || version?.() !== v0 }));
    if (ir) return ir;
  }
}

export interface WriteIoOptions {
  /** Kitabın içerik sürümü (değişirse anlık görüntü yeniden alınır). */
  version?: () => unknown;
  onProgress?: (p: IoProgress) => void;
}

/**
 * Kitabı istenen biçimde Worker'da yazar (formats.writeBook ile aynı sonuç). `report.warnings`:
 * .xlsx/.xlsm'de korunamayan öğeler.
 */
export async function writeBookIo(wb: WorkbookModel, sheet: SheetModel, format: SaveFormat, delimiter?: string, report?: XlsxWriteReport, opts: WriteIoOptions = {}): Promise<Uint8Array<ArrayBuffer> | string> {
  const ir = await snapshotIr(wb, opts.version);
  const src = getSource(wb.kaynak);
  try {
    // IR'deki tamponlar aktarılır; nesneler ileti gönderilirken (bu görevde) kopyalanır
    const r = await writeIo.call("write", { ir, format, sheet: Math.max(0, wb.sheets.indexOf(sheet)), delimiter, source: src?.bytes }, { transfer: irTransfer(ir), onProgress: opts.onProgress });
    if (report) report.warnings = r.warnings;
    return r.data;
  } catch (e) {
    if (!workerUnavailable(e)) throw e;
    console.warn("[tablo] dosya G/Ç Worker'ı kurulamadı; kitap ana iş parçacığında yazılıyor", e);
    return writeBook(wb, sheet, format, delimiter, report);
  }
}

/**
 * Otomatik kayıt doğrulaması (yalnız .xlsx): yazılan baytlar kendi okuyucumuzla açılabiliyor mu ve
 * en az `sheets` sayfa var mı. Tümüyle Worker'da (baytlar kopyalanır: Worker kurulamazsa süreç içinde denenir).
 */
export async function verifyXlsxIo(data: ArrayBuffer, sheets = 1): Promise<boolean> {
  try {
    return await verifyIo.call("verify", { bytes: data, sheets });
  } catch (e) {
    if (!workerUnavailable(e)) throw e;
    const opened = await openBook("dogrulama.xlsx", data, { kaynak: false });
    return opened.wb.sheets.length >= Math.max(1, sheets);
  } finally {
    if (!verifyIo.pending) verifyIo.terminate();
  }
}
