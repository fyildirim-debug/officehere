/**
 * Dosya G/Ç Worker'ının işleyicileri (io.worker.ts yayımlar; Worker olmayan ortamda süreç içi
 * taşıyıcı aynı işleyicileri çağırır). Açma: bayt → kitap → ara gösterim (IR, aktarılır).
 * Yazma: IR → kitap → bayt. Doğrulama (otomatik kayıt): yazılan baytlar kendi okuyucumuzla açılıyor
 * mu ve sayfaların hiçbiri kaybolmamış mı; yalnız sonuç döner, kaynak kaydedilmez.
 */
import { isCfb, withTransfer, type IoContext } from "@officehere/core";
import { OpcPackage, getSource, registerSource, releaseSource } from "@officehere/core/opc";
import { bookFromIr, bookToIr, irTransfer, runSync, type BookIr } from "./bookIr";
import { decodeText, detectDelimiter } from "./csv";
import { openBook, sniffFormat, writeBook, type BookFormat, type SaveFormat } from "./formats";

/** Açma sonucu: ana iş parçacığı kitabı IR'den kurar. */
export interface ReadBookResult {
  kind: "book";
  ir: BookIr;
  format: BookFormat;
  template: boolean;
  notice?: string;
  /** Metin dosyasının ayırıcısı (aynı biçimde kaydederken korunur). */
  delimiter?: string;
  /** .xlsx kaynak paketi (onarılmış özgün baytlar): ana iş parçacığında kaydedilir. */
  source?: Uint8Array;
}

/** Worker'da okunamayan biçim (HTML: DOMParser gerekir): ana iş parçacığında açılır. */
export interface ReadOnMain {
  kind: "main";
}

export interface WriteArgs {
  ir: BookIr;
  format: SaveFormat;
  /** Metin biçimlerinde yazılacak sayfanın sırası. */
  sheet: number;
  delimiter?: string;
  /** Kaynak paket (IR'deki kimlikle); .xlsx yazımında tanınmayan parçalar buradan taşınır. */
  source?: Uint8Array;
}

/** Bu Worker'ın kendi kaydettiği kaynaklar (süreç içi taşıyıcıda ana kayıt paylaşılır; onlara dokunulmaz). */
const owned = new Set<string>();

async function useSource(id: string | undefined, bytes: Uint8Array | undefined): Promise<void> {
  if (id && bytes && !getSource(id)) {
    registerSource(await OpcPackage.open(bytes), { id });
    owned.add(id);
  }
  // Önceki kitapların kaynakları bırakılır (Worker bellekte tek kitabın kaynağını tutar)
  for (const o of owned) {
    if (o === id) continue;
    releaseSource(o);
    owned.delete(o);
  }
}

/** Metin biçimi mi (ZIP ya da bileşik dosya değil): yalnız bunlarda HTML olabilir. */
function isText(data: ArrayBuffer): boolean {
  const b = new Uint8Array(data, 0, Math.min(4, data.byteLength));
  return !(b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04) && !isCfb(data);
}

export const handlers = {
  async read(a: { bytes: ArrayBuffer; name: string }, ctx: IoContext) {
    ctx.progress(0, "Dosya okunuyor");
    if (isText(a.bytes) && (await sniffFormat(a.name, a.bytes)) === "html") return { kind: "main" } satisfies ReadOnMain;
    const opened = await openBook(a.name, a.bytes, { progress: (r, step) => ctx.progress(r * 0.9, step) });
    const wb = opened.wb;
    // Kaynak paket Worker'da değil, kitabı kullanan ana iş parçacığında kaydedilir
    const pkg = getSource(wb.kaynak);
    const source = pkg ? pkg.bytes.slice() : undefined;
    releaseSource(wb.kaynak);
    wb.kaynak = undefined;
    ctx.progress(0.92, "Aktarılıyor");
    wb.engine.recalc();
    const ir = runSync(bookToIr(wb))!;
    const delimiter = opened.format === "csv" || opened.format === "txt" ? detectDelimiter(decodeText(a.bytes)) : opened.format === "tsv" ? "\t" : undefined;
    const out: ReadBookResult = { kind: "book", ir, format: opened.format, template: opened.template, notice: opened.notice, delimiter, source };
    return withTransfer(out, [...irTransfer(ir), ...(source ? [source.buffer as ArrayBuffer] : [])]);
  },

  async write(a: WriteArgs, ctx: IoContext) {
    ctx.progress(0, "Çalışma kitabı hazırlanıyor");
    await useSource(a.ir.kaynak, a.source);
    const wb = runSync(bookFromIr(a.ir));
    wb.kaynak = a.ir.kaynak;
    ctx.progress(0.2, "Yazılıyor");
    const report = { warnings: [] as string[] };
    const data = await writeBook(wb, wb.sheets[a.sheet] ?? wb.active, a.format, a.delimiter, report);
    ctx.progress(1);
    return withTransfer({ data, warnings: report.warnings }, typeof data === "string" ? [] : [data.buffer as ArrayBuffer]);
  },

  /** Yazılan .xlsx kendi okuyucumuzla açılabiliyor mu ve en az `sheets` sayfası var mı. */
  async verify(a: { bytes: ArrayBuffer; sheets: number }) {
    const opened = await openBook("dogrulama.xlsx", a.bytes, { kaynak: false });
    return opened.wb.sheets.length >= Math.max(1, a.sheets);
  },
};

export type IoHandlers = typeof handlers;
