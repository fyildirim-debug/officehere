/**
 * Dosya eylemleri: Aç, Kaydet, Farklı Kaydet, Yeni, Dışarı Aktar. Biçim içerik
 * imzasıyla tanınır; yazılamayan biçimler (.xls, şablon) Kaydet'te .xlsx Farklı
 * Kaydet'e yönlenir; makrolu kitap .xlsm olarak kaydedilir; .ods/.csv/.tsv ilk
 * kayıtta bir kez sorulur.
 */
import { useMemo } from "react";
import { isDirectSave, trackSave } from "@officehere/ui";
import { addRecentFile, baseName, downloadBlob, getNative, isAbortError, openFile, readPath, saveFile, saveFileAs, type FileRef, type RecentFile } from "@officehere/core";
import type { Controller } from "../controller";
import { openBookIo, writeBookIo } from "../io/client";
import { FILTERS, OPEN_FILTERS, saveFormatOf, writableAs, type SaveFormat } from "../io/formats";
import { writeHtml } from "../io/html";
import { exportPdf } from "../io/print";
import { printWithOptions } from "../io/printPages";
import type { XlsxWriteReport } from "../io/xlsx";
import { buildTemplate, TEMPLATES } from "../templates";
import type { FormatChoice } from "../dialogs/SaveFormatDialog";
import { fileFormatOf, setFileFormat } from "./fileFormat";
import { startFileTask } from "./fileProgress";

export interface FileActions {
  save: () => Promise<boolean>;
  /** Farklı Kaydet; biçim verilmezse geçerli yazılabilir biçim (yoksa .xlsx) önerilir. */
  saveAs: (format?: SaveFormat) => Promise<boolean>;
  open: () => Promise<void>;
  openPath: (p: string) => Promise<void>;
  openRecent: (f: RecentFile) => Promise<void>;
  newBook: (templateId: string) => Promise<void>;
  exportAs: (kind: string) => Promise<void>;
  print: () => void;
  /** Dosyayı yükler; başarıda true. */
  loadData: (name: string, data: ArrayBuffer, path?: string) => Promise<boolean>;
  confirmDiscard: () => Promise<boolean>;
}

const SAVE_ORDER: SaveFormat[] = ["xlsx", "ods", "csv", "tsv"];

/** Biçim sorusu (iletişim kutusu yanıtı). */
function askFormat(ctl: Controller, format: SaveFormat): Promise<FormatChoice> {
  return new Promise((resolve) => ctl.openDialog("saveFormat", { format, onChoose: resolve }));
}

/** CSV/TSV kaydında kullanıcıya bilgi. */
function textSavedNotice(ctl: Controller, format: SaveFormat): void {
  if (format !== "csv" && format !== "tsv") return;
  ctl.infoBar = { kind: "info", text: `${format === "csv" ? "CSV" : "Sekmeyle ayrılmış metin"} olarak kaydedildi: yalnız etkin sayfanın («${ctl.sheet.name}») değerleri kaydedildi. Diğer sayfalar ve biçimler için .xlsx olarak kaydedin.` };
}

/** .xlsx/.xlsm kaydında korunamayan öğeler (dosyadan gelen, yapı değişince kaydırılamayan…). */
function savedWarnings(ctl: Controller, report: XlsxWriteReport): void {
  if (!report.warnings.length) return;
  ctl.infoBar = { kind: "warning", text: `Kaydedildi. Şu öğeler bu dosyada korunamadı: ${report.warnings.join(", ")}.` };
  ctl.emit();
}

let bookCounter = 1;

export function useFileActions(ctl: Controller): FileActions {
  return useMemo(() => {
    // Yazma Worker'da; büyük kitapta durum çubuğunda ilerleme (iptal yok: dosya yarım yazılmaz)
    const write = async (format: SaveFormat, delimiter?: string, report?: XlsxWriteReport) => {
      const task = startFileTask("Kaydediliyor…");
      try {
        return await writeBookIo(ctl.wb, ctl.sheet, format, delimiter, report, { version: () => ctl.dataVersion, onProgress: (p) => task.update(p) });
      } finally {
        task.end();
      }
    };
    // Dosyanın kendi ayırıcısı yalnız aynı biçimde kaydederken korunur
    const data = (format: SaveFormat, report?: XlsxWriteReport) => {
      const st = fileFormatOf(ctl.wb);
      return write(format, writableAs(st.format) === format ? st.delimiter : undefined, report);
    };
    /** Süren açma (yenisi başlarken iptal edilir: aynı Worker'da iki açma birbirini beklemesin). */
    let opening: AbortController | null = null;
    const done = (ref: FileRef, format: SaveFormat, notice?: string) => {
      ctl.fileRef = ref;
      ctl.fileName = ref.name;
      ctl.dirty = false;
      const prev = fileFormatOf(ctl.wb);
      setFileFormat(ctl.wb, { format, confirmed: true, delimiter: format === prev.format ? prev.delimiter : undefined });
      // Uyumluluk/makro iletisi .xlsx'e kaydedince kalkar
      if (notice && ctl.infoBar?.text === notice) ctl.infoBar = null;
      textSavedNotice(ctl, format);
      void addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "tablo", openedAt: Date.now() });
      ctl.emit();
    };
    let openNotice: string | undefined;
    const actions: FileActions = {
      confirmDiscard: async () => true,
      async loadData(name, buf, path) {
        opening?.abort();
        const ctrl = new AbortController();
        opening = ctrl;
        const task = startFileTask("Açılıyor…", () => ctrl.abort());
        const version = ctl.dataVersion;
        try {
          const opened = await openBookIo(name, buf, { signal: ctrl.signal, onProgress: (p) => task.update(p) });
          task.end();
          // Açılış sürerken açık kitap düzenlendiyse değişiklikler sorulmadan atılmaz
          if (ctl.dirty && ctl.dataVersion !== version && !(await actions.confirmDiscard())) return false;
          if (opened.template) {
            // Şablon: adsız yeni çalışma kitabı (Kaydet → Farklı Kaydet .xlsx)
            ctl.load(opened.wb, `${baseName(name)}1`, null);
            setFileFormat(opened.wb, { format: "xlsx" });
          } else {
            ctl.load(opened.wb, name, { name, path });
            setFileFormat(opened.wb, { format: opened.format, delimiter: opened.delimiter });
          }
          openNotice = opened.notice;
          if (opened.notice) {
            const warn = opened.wb.loadWarnings.length ? ` Atlanan öğeler: ${opened.wb.loadWarnings.join(", ")}.` : "";
            openNotice = opened.notice + warn;
            ctl.infoBar = { kind: opened.wb.loadWarnings.length ? "warning" : "info", text: openNotice };
            ctl.emit();
          }
          void addRecentFile({ path: path ?? name, name, app: "tablo", openedAt: Date.now() });
          return true;
        } catch (e) {
          task.end();
          if (isAbortError(e)) return false;
          // openBook her hatayı anlaşılır Türkçe iletiye çevirir; kitaplığın teknik iletisi ayrıntıdadır.
          // Worker'ın kendi hataları (çöktü, kesildi) da Türkçe iletilidir.
          const err = e as Error & { detail?: string };
          const known = err?.name === "FormatError" || err?.name === "IoWorkerError";
          if (err?.detail) console.warn(`[tablo] "${name}" açılamadı: ${err.detail}`);
          ctl.alert(
            known ? `"${name}" dosyası açılamadı.\n\n${err.message}` : `"${name}" dosyası açılamadı. Dosya biçimi geçersiz ya da dosya bozuk olabilir.`,
            undefined,
            known ? err.detail : `${err?.name ?? "Error"}: ${err?.message ?? e}`,
          );
          return false;
        } finally {
          if (opening === ctrl) opening = null;
        }
      },
      async open() {
        if (!(await actions.confirmDiscard())) return;
        const f = await openFile(OPEN_FILTERS);
        if (!f) return;
        // Açılamayan dosyanın tanıtıcısı açık kitaba bağlanmaz; şablonun tanıtıcısı yoktur
        if ((await actions.loadData(f.name, f.data, f.path)) && f.handle && ctl.fileRef) ctl.fileRef = { name: f.name, handle: f.handle, path: f.path };
      },
      async openPath(p) {
        const native = getNative();
        if (native && !ctl.pristine) {
          await native.newWindow("tablo", p);
          return;
        }
        const f = await readPath(p);
        if (f) await actions.loadData(f.name, f.data, f.path);
      },
      async openRecent(f) {
        if (getNative()) await actions.openPath(f.path);
        else await actions.open();
      },
      async save() {
        if (ctl.edit && !ctl.commitEdit(null)) return false;
        const state = fileFormatOf(ctl.wb);
        const format = writableAs(state.format);
        // Yazılamayan biçim (.xls, makrolu, HTML): özgün dosyanın üzerine yazılmaz
        if (!format) return actions.saveAs("xlsx");
        if (format !== "xlsx" && format !== "xlsm" && !state.confirmed) {
          const choice = await askFormat(ctl, format);
          if (!choice) return false;
          if (choice === "xlsx") return actions.saveAs("xlsx");
          setFileFormat(ctl.wb, { ...state, confirmed: true });
        }
        const name = ctl.fileName.includes(".") ? ctl.fileName : `${ctl.fileName}.${format}`;
        const target = ctl.fileRef;
        const report: XlsxWriteReport = { warnings: [] };
        const bytes = await data(format, report);
        const ref = await trackSave(isDirectSave(target), async () => saveFile(target, name, [FILTERS[format]], bytes));
        if (!ref) return false;
        done(ref, saveFormatOf(ref.name) ?? format, openNotice);
        savedWarnings(ctl, report);
        return true;
      },
      async saveAs(preferred) {
        if (ctl.edit && !ctl.commitEdit(null)) return false;
        const format: SaveFormat = preferred ?? writableAs(fileFormatOf(ctl.wb).format) ?? "xlsx";
        const filters = [FILTERS[format], ...SAVE_ORDER.filter((f) => f !== format).map((f) => FILTERS[f])];
        const name = `${baseName(ctl.fileName)}.${format}`;
        const report: XlsxWriteReport = { warnings: [] };
        const bytes = await data(format, report);
        let finalFormat = format;
        const ref = await trackSave(false, async () => {
          const r = await saveFileAs(name, filters, bytes);
          if (!r) return null;
          // İletişim kutusunda başka biçim seçildiyse dosya o biçimde yeniden yazılır
          const chosen = saveFormatOf(r.name);
          if (chosen && chosen !== format && (r.path || r.handle)) {
            finalFormat = chosen;
            report.warnings = [];
            return saveFile(r, r.name, [FILTERS[chosen]], await write(chosen, undefined, report));
          }
          return r;
        });
        if (!ref) return false;
        done(ref, finalFormat, openNotice);
        savedWarnings(ctl, report);
        return true;
      },
      async newBook(id) {
        if (!(await actions.confirmDiscard())) return;
        const wb = buildTemplate(id);
        const n = ++bookCounter;
        const name = id === "blank" ? `Kitap${n}` : `${TEMPLATES.find((t) => t.id === id)?.label ?? "Kitap"}`;
        ctl.load(wb, name, null);
        ctl.pristine = id === "blank";
      },
      async exportAs(kind) {
        const base = baseName(ctl.fileName);
        if (kind === "csv") {
          await saveFileAs(`${base}.csv`, [FILTERS.csv], await write("csv", ";"));
        } else if (kind === "csv-comma") {
          await saveFileAs(`${base}.csv`, [{ name: "CSV (Virgülle ayrılmış)", extensions: ["csv"] }], await write("csv", ","));
        } else if (kind === "tsv") {
          await saveFileAs(`${base}.tsv`, [FILTERS.tsv], await write("tsv", "\t"));
        } else if (kind === "ods") {
          await saveFileAs(`${base}.ods`, [FILTERS.ods], await write("ods"));
        } else if (kind === "pdf") {
          await exportPdf(ctl.wb, ctl.sheet, `${base}.pdf`);
        } else if (kind === "html" || kind === "html-all") {
          const sheets = kind === "html" ? [ctl.sheet] : ctl.wb.sheets.filter((s) => !s.hidden);
          downloadBlob(new Blob([writeHtml(ctl.wb, sheets, base)], { type: "text/html" }), `${base}.html`);
        } else if (kind === "xlsx") {
          const bytes = await write("xlsx");
          downloadBlob(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${base}.xlsx`);
        }
      },
      print() {
        printWithOptions(ctl.wb, ctl.sheet, {});
      },
    };
    return actions;
  }, [ctl]);
}
