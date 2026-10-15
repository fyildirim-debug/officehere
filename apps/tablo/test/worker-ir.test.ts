/**
 * Worker yolu (P4.2): dosya Worker'da okunup ara gösterimle (IR) ana iş parçacığına aktarılır ve
 * kitap orada toplu yüklemeyle kurulur. Bu test, derlemdeki her Tablo dosyasında Worker yolunun
 * kurduğu kitabın süreç içi okumayla (openBook) aynı olduğunu ve Worker'da yazılan dosyanın süreç
 * içi yazılanla aynı kitabı verdiğini denetler. Vitest'te Worker yoktur: istemci aynı işleyicileri
 * süreç içi taşıyıcıyla (yapılandırılmış kopya + aktarım) çağırır.
 *
 * Derlem: test-derlem/ + uygulama fikstürleri (+ DERLEM_CIKTI ile verilen dizinde üretilen dosyalar).
 */
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { basename } from "node:path";
import { isError, type Scalar } from "@officehere/formula";
import { getSource } from "@officehere/core/opc";
import { AxisSizes } from "../src/model/axis";
import { defaultTheme } from "../src/model/theme";
import { activateTheme } from "../src/model/themeRuntime";
import { SheetModel, WorkbookModel } from "../src/model/workbook";
import { openBook, writeBook, type SaveFormat } from "../src/io/formats";
import { decodeText, detectDelimiter } from "../src/io/csv";
import { openBookIo, snapshotIr, writeBookIo, verifyXlsxIo } from "../src/io/client";
import { bookFromIr, bookToIr, irTransfer, runSliced, runSync } from "../src/io/bookIr";
// @ts-expect-error -- derlem betiklerinin ortak modülü (JS)
import { derlemDosyalari } from "../../../scripts/derlem/ortak.mjs";

const files: { yol: string; anahtar: string }[] = derlemDosyalari({ uygulama: "tablo" });
const ab = (u: Uint8Array): ArrayBuffer => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const read = (p: string): ArrayBuffer => ab(readFileSync(p));

// Rastgele ve saate bağlı işlevler (RAND, NOW) iki okumada aynı sonucu versin
beforeAll(() => {
  vi.useFakeTimers({ toFake: ["Date"], now: new Date("2026-03-04T05:06:07Z") });
  vi.spyOn(Math, "random").mockReturnValue(0.375);
});
afterAll(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

function axisPlain(a: AxisSizes) {
  return { def: a.defaultSize, count: a.count, custom: a.customItems(), manual: a.manualItems().sort((x, y) => x - y), hidden: a.hiddenItems(), levels: a.levelItems(), collapsed: a.collapsedItems() };
}

const val = (v: Scalar) => (isError(v) ? { hata: v.code, ileti: v.message } : v);

/** Kitabın karşılaştırılabilir dökümü: modelin tüm genel alanları, hücreler, motor durumu. */
function dump(wb: WorkbookModel) {
  const e = wb.engine;
  return {
    sheets: wb.sheets.map((s) => {
      const fields: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(s)) fields[k] = v instanceof AxisSizes ? axisPlain(v) : v;
      const cells: unknown[] = [];
      e.forEachCell(s.id, (r, c, rec) => cells.push([r, c, val(rec.value), rec.formula?.text ?? null]));
      const values: unknown[] = [];
      e.forEachValue(s.id, (r, c, v) => values.push([r, c, val(v)]));
      return { fields, cells, values, extent: e.getUsedExtent(s.id), used: wb.usedRange(s) };
    }),
    styles: wb.styles.all(),
    names: e.getNames(),
    tables: e.getTables(),
    circular: e.getCircularRefs(),
    activeIndex: wb.activeIndex,
    props: wb.props,
    theme: wb.theme,
    loadWarnings: wb.loadWarnings,
    protection: wb.protection,
    passthrough: wb.passthrough,
  };
}

/** Uygulamanın açtıktan sonra yaptığı (Controller.load): tablolar ve bekleyen hesap. */
function settle(wb: WorkbookModel): WorkbookModel {
  wb.syncTables();
  wb.engine.recalc();
  return wb;
}

const sourceBytes = (wb: WorkbookModel) => getSource(wb.kaynak)?.bytes;
const XLSX = /\.(xlsx|xlsm|xltx|xltm)$/i;

describe("Worker yolu: derlemde okuma modelleri süreç içiyle eşit", () => {
  it("derlemde Tablo dosyası var", () => expect(files.length).toBeGreaterThan(20));
  for (const f of files) {
    it(f.anahtar, { timeout: 120_000 }, async () => {
      const data = read(f.yol);
      let inproc: Awaited<ReturnType<typeof openBook>> | undefined;
      let inprocError: unknown;
      // Her okuma yeni bir Worker'daki (ya da yeni açılan uygulamadaki) gibi varsayılan temayla başlar:
      // okuyucular etkin temanın yazı tipine göre normalleştirir ve süreç içinde ikisi aynı durumu paylaşır
      activateTheme(defaultTheme());
      try {
        inproc = await openBook(basename(f.yol), data);
      } catch (e) {
        inprocError = e;
      }
      activateTheme(defaultTheme());
      if (!inproc) {
        // Açılamayan dosya Worker yolunda da aynı Türkçe iletiyle açılamaz
        const err = await openBookIo(basename(f.yol), data).then(() => null, (e: unknown) => e);
        expect(err).toMatchObject({ name: "FormatError", message: (inprocError as Error).message });
        return;
      }
      const io = await openBookIo(basename(f.yol), data);
      // Çağıranın baytları kopyalanarak gönderilir (bozulmaz)
      expect(data.byteLength).toBe(readFileSync(f.yol).byteLength);
      expect({ format: io.format, template: io.template, notice: io.notice }).toEqual({ format: inproc.format, template: inproc.template, notice: inproc.notice });
      const delimiter = inproc.format === "csv" || inproc.format === "txt" ? detectDelimiter(decodeText(data)) : inproc.format === "tsv" ? "\t" : undefined;
      expect(io.delimiter).toBe(delimiter);
      expect(dump(settle(io.wb))).toEqual(dump(settle(inproc.wb)));
      // Kaynak paket ana iş parçacığında, aynı baytlarla kayıtlı
      const a = sourceBytes(inproc.wb);
      const b = sourceBytes(io.wb);
      expect(!!b).toBe(!!a);
      if (a && b) expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true);
      if (a) expect(io.wb.kaynak).not.toBe(inproc.wb.kaynak);

      // Yazma: Worker'da yazılan dosya süreç içi yazılanla aynı kitabı verir
      const formats: SaveFormat[] = XLSX.test(f.yol) ? ["xlsx"] : ["xlsx", "ods"];
      if (inproc.format === "xlsm") formats.push("xlsm");
      for (const format of formats) {
        const ra = { warnings: [] as string[] };
        const rb = { warnings: [] as string[] };
        const wa = await writeBook(inproc.wb, inproc.wb.active, format, undefined, ra);
        const wbIo = await writeBookIo(io.wb, io.wb.active, format, undefined, rb);
        expect(rb.warnings).toEqual(ra.warnings);
        const name = `yeniden.${format}`;
        const ya = await openBook(name, ab(wa as Uint8Array));
        const yb = await openBook(name, ab(wbIo as Uint8Array));
        expect(dump(settle(yb.wb))).toEqual(dump(settle(ya.wb)));
      }
    });
  }
});

describe("ara gösterim (IR)", () => {
  it("modelin her alanı taşınır (yeni alan eklenince bu test IR'nin güncellenmesini ister)", async () => {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    const ir = runSync(bookToIr(wb))!;
    const known = new Set([...Object.keys(ir.sheets[0].props), "id", "name", "rows", "cols", "styles"]);
    expect(Object.keys(s).filter((k) => !known.has(k))).toEqual([]);
    const wbKnown = new Set(["engine", "styles", "sheets", "activeIndex", "props", "theme", "loadWarnings", "protection", "kaynak", "passthrough", "nextObjectId", "usedCache", "styleWrites"]);
    expect(Object.keys(wb).filter((k) => !wbKnown.has(k))).toEqual([]);
    expect(Object.keys(new SheetModel(1, "x")).length).toBe(Object.keys(s).length);
  });

  it("boyutlar birebir: varsayılana eşit özel boyut, boyutsuz elle kaydı, anahat ve gizlilik", () => {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    s.rows.set(3, 20, true); // varsayılana eşit: kayıt yok, elle işaretli
    s.rows.set(4, 35, false);
    s.rows.setDefault(35); // 4. satır artık varsayılana eşit özel boyut
    s.rows.set(7, 12.345, true);
    s.rows.setHidden(9, true);
    s.rows.setLevel(10, 2);
    s.rows.setCollapsed(11, true);
    s.cols.set(2, 0, true);
    const back = runSync(bookFromIr(structuredClone(runSync(bookToIr(wb))!)));
    expect(axisPlain(back.sheets[0].rows)).toEqual(axisPlain(s.rows));
    expect(axisPlain(back.sheets[0].cols)).toEqual(axisPlain(s.cols));
    expect(back.sheets[0].rows.size(4)).toBe(35);
  });

  it("değerler: boş metin, yinelenen metin, iletili hata, mantıksal, formül önbelleği; sayfa kimlikleri korunur", () => {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    wb.addSheet("İkinci");
    wb.removeSheet(0);
    wb.addSheet("Üçüncü");
    const [a, b] = wb.sheets;
    wb.setInputRaw(a, 0, 0, "");
    wb.setInputRaw(a, 1, 0, "ığüşöç İĞÜŞÖÇ");
    wb.setInputRaw(a, 2, 0, "ığüşöç İĞÜŞÖÇ");
    wb.setInputRaw(a, 3, 0, true);
    wb.setInputRaw(b, 0, 0, { formula: "=İkinci!A2&\"!\"" });
    wb.setInputRaw(b, 1, 0, { formula: "=1/0" });
    wb.engine.defineName("Yerel", "=İkinci!$A$4", a.id);
    wb.engine.recalc();
    const ir = runSync(bookToIr(wb))!;
    // Yinelenen metin bir kez: "", metin, iki formül, formül sonucu, hata kodu
    expect(ir.strings.lens.length).toBe(6);
    const back = runSync(bookFromIr(structuredClone(ir, { transfer: irTransfer(ir) })));
    expect(back.sheets.map((x) => [x.id, x.name])).toEqual([
      [a.id, "İkinci"],
      [b.id, "Üçüncü"],
    ]);
    expect(back.engine.pending).toBe(false);
    expect(dump(back)).toEqual(dump(wb));
    expect(back.getValue(back.sheets[1], 0, 0)).toBe("ığüşöç İĞÜŞÖÇ!");
    expect(back.engine.getNames()).toEqual(wb.engine.getNames());
    void s;
  });

  it("parça parça çıkarılırken kitap değişirse anlık görüntü bayattır (null)", async () => {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    for (let r = 0; r < 20_000; r++) wb.setInputRaw(s, r, 0, r);
    const g = bookToIr(wb, { stale: () => wb.engine.inputVersion !== iv });
    const iv = wb.engine.inputVersion;
    g.next();
    wb.setInputRaw(s, 0, 1, "değişti");
    let r = g.next();
    while (!r.done) r = g.next();
    expect(r.value).toBeNull();
    // snapshotIr yeniden dener ve güncel kitabı verir
    const ir = await snapshotIr(wb);
    const back = runSync(bookFromIr(ir));
    expect(back.getValue(back.sheets[0], 0, 1)).toBe("değişti");
  });

  it("dilimli kurulum iptal edilir (AbortError) ve ilerleme bildirir", async () => {
    const wb = new WorkbookModel();
    const s = wb.sheets[0];
    for (let r = 0; r < 50_000; r++) wb.setInputRaw(s, r, 0, `satır ${r}`);
    const ir = runSync(bookToIr(wb))!;
    const ctrl = new AbortController();
    const seen: number[] = [];
    const p = runSliced(bookFromIr(ir), { budgetMs: 1, signal: ctrl.signal, onProgress: (x) => (seen.push(x), seen.length === 3 && ctrl.abort()) });
    await expect(p).rejects.toMatchObject({ name: "AbortError" });
    expect(seen.length).toBe(3);
    expect(seen[2]).toBeGreaterThan(0);
  });

  it("otomatik kayıt doğrulaması: geçerli dosya, eksik sayfa ve bozuk bayt", async () => {
    const wb = new WorkbookModel();
    wb.addSheet("İki");
    const bytes = (await writeBook(wb, wb.sheets[0], "xlsx")) as Uint8Array;
    expect(await verifyXlsxIo(ab(bytes), 2)).toBe(true);
    expect(await verifyXlsxIo(ab(bytes), 3)).toBe(false);
    await expect(verifyXlsxIo(new ArrayBuffer(10), 1)).rejects.toThrow();
  });
});
