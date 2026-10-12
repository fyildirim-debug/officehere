/** Dosya görünümü (Backstage) ve dosya eylemleri: Yeni, Aç, Kaydet, Farklı Kaydet, Yazdır, Dışarı Aktar. */
import { useMemo } from "react";
import { DocumentPdf24Regular, DocumentTable24Regular, Globe24Regular, ArrowDownload24Regular } from "@fluentui/react-icons";
import { OfficeBackstage, type BackstageTemplate } from "@officehere/ui";
import { addRecentFile, baseName, downloadBlob, extensionOf, getNative, openFile, readPath, saveFile, saveFileAs, type FileFilter, type RecentFile } from "@officehere/core";
import type { Controller } from "./controller";
import { readXlsx, writeXlsx } from "./io/xlsx";
import { readCsv, writeCsv } from "./io/csv";
import { printRange, printSheet, rangeToHtml } from "./io/print";
import { buildTemplate, TEMPLATES } from "./templates";
import type { WorkbookModel } from "./model/workbook";

const XLSX_FILTER: FileFilter = { name: "Çalışma Kitabı (.xlsx)", extensions: ["xlsx"] };
const CSV_FILTER: FileFilter = { name: "CSV (Noktalı virgülle ayrılmış)", extensions: ["csv"] };
const OPEN_FILTERS: FileFilter[] = [{ name: "Tüm Çalışma Kitapları", extensions: ["xlsx", "xlsm", "csv", "txt"] }, XLSX_FILTER, CSV_FILTER];

export interface FileActions {
  save: () => Promise<boolean>;
  saveAs: () => Promise<boolean>;
  open: () => Promise<void>;
  openPath: (p: string) => Promise<void>;
  openRecent: (f: RecentFile) => Promise<void>;
  newBook: (templateId: string) => Promise<void>;
  exportAs: (kind: string) => Promise<void>;
  print: () => void;
  loadData: (name: string, data: ArrayBuffer, path?: string) => Promise<void>;
  confirmDiscard: () => Promise<boolean>;
}

async function workbookFromData(name: string, data: ArrayBuffer): Promise<WorkbookModel> {
  const ext = extensionOf(name);
  if (ext === "csv" || ext === "txt") return readCsv(data, baseName(name));
  return readXlsx(data);
}

export function useFileActions(ctl: Controller): FileActions {
  return useMemo(() => {
    const actions: FileActions = {
      confirmDiscard: async () => true,
      async loadData(name, data, path) {
        try {
          const wb = await workbookFromData(name, data);
          ctl.load(wb, name, { name, path });
          void addRecentFile({ path: path ?? name, name, app: "tablo", openedAt: Date.now() });
        } catch (e) {
          ctl.alert(`"${name}" dosyası açılamadı. Dosya biçimi geçersiz ya da dosya bozuk olabilir.\n\n${(e as Error).message ?? ""}`);
        }
      },
      async open() {
        if (!(await actions.confirmDiscard())) return;
        const f = await openFile(OPEN_FILTERS);
        if (!f) return;
        await actions.loadData(f.name, f.data, f.path);
        if (f.handle) ctl.fileRef = { name: f.name, handle: f.handle, path: f.path };
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
        const name = ctl.fileName.includes(".") ? ctl.fileName : `${ctl.fileName}.xlsx`;
        const isCsv = extensionOf(name) === "csv";
        const data = isCsv ? writeCsv(ctl.wb, ctl.sheet) : await writeXlsx(ctl.wb);
        const ref = await saveFile(ctl.fileRef, name, isCsv ? [CSV_FILTER] : [XLSX_FILTER], data);
        if (!ref) return false;
        ctl.fileRef = ref;
        ctl.fileName = ref.name;
        ctl.dirty = false;
        void addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "tablo", openedAt: Date.now() });
        ctl.emit();
        return true;
      },
      async saveAs() {
        if (ctl.edit && !ctl.commitEdit(null)) return false;
        const name = `${baseName(ctl.fileName)}.xlsx`;
        const data = await writeXlsx(ctl.wb);
        const ref = await saveFileAs(name, [XLSX_FILTER], data);
        if (!ref) return false;
        ctl.fileRef = ref;
        ctl.fileName = ref.name;
        ctl.dirty = false;
        void addRecentFile({ path: ref.path ?? ref.name, name: ref.name, app: "tablo", openedAt: Date.now() });
        ctl.emit();
        return true;
      },
      async newBook(id) {
        if (!(await actions.confirmDiscard())) return;
        const wb = buildTemplate(id);
        const n = nextBookNumber();
        const name = id === "blank" ? `Kitap${n}` : `${TEMPLATES.find((t) => t.id === id)?.label ?? "Kitap"}`;
        ctl.load(wb, name, null);
        ctl.pristine = id === "blank";
      },
      async exportAs(kind) {
        const base = baseName(ctl.fileName);
        if (kind === "csv") {
          await saveFileAs(`${base}.csv`, [CSV_FILTER], writeCsv(ctl.wb, ctl.sheet));
        } else if (kind === "csv-comma") {
          await saveFileAs(`${base}.csv`, [{ name: "CSV (Virgülle ayrılmış)", extensions: ["csv"] }], writeCsv(ctl.wb, ctl.sheet, { delimiter: "," }));
        } else if (kind === "pdf") {
          printSheet(ctl.wb, ctl.sheet);
        } else if (kind === "html") {
          const r = printRange(ctl.wb, ctl.sheet);
          const body = r ? rangeToHtml(ctl.wb, ctl.sheet, r, { gridlines: true, objects: true }) : "";
          const html = `<!doctype html><html lang="tr"><head><meta charset="utf-8"><title>${base}</title><style>table{border-collapse:collapse;table-layout:fixed;font-family:Calibri,Carlito,sans-serif}td{padding:0 4px;border:1px solid #d4d4d4;overflow:hidden;white-space:nowrap}.tb-print-sheet{position:relative;display:inline-block}.tb-print-obj{position:absolute;overflow:hidden}.tb-chart-tip{display:none}</style></head><body>${body}</body></html>`;
          downloadBlob(new Blob([html], { type: "text/html" }), `${base}.html`);
        } else if (kind === "xlsx") {
          const data = await writeXlsx(ctl.wb);
          downloadBlob(new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${base}.xlsx`);
        }
      },
      print() {
        printSheet(ctl.wb, ctl.sheet);
      },
    };
    return actions;
  }, [ctl]);
}

let bookCounter = 1;
function nextBookNumber(): number {
  return ++bookCounter;
}

function TemplatePreview({ id }: { id: string }) {
  if (id === "blank") {
    return (
      <svg viewBox="0 0 120 90" width="100%" height="100%" aria-hidden="true">
        <rect width="120" height="90" fill="#fff" />
        {Array.from({ length: 9 }, (_, i) => (
          <line key={`h${i}`} x1="0" x2="120" y1={i * 10 + 10} y2={i * 10 + 10} stroke="#E0E0E0" />
        ))}
        {Array.from({ length: 5 }, (_, i) => (
          <line key={`v${i}`} y1="0" y2="90" x1={i * 24 + 24} x2={i * 24 + 24} stroke="#E0E0E0" />
        ))}
        <rect x="0.5" y="0.5" width="24" height="10" fill="none" stroke="#107C41" strokeWidth="1.5" />
      </svg>
    );
  }
  const color = id === "budget" ? "#156082" : id === "invoice" ? "#196B24" : id === "todo" ? "#A02B93" : "#E97132";
  const title = TEMPLATES.find((t) => t.id === id)?.label ?? "";
  return (
    <svg viewBox="0 0 120 90" width="100%" height="100%" aria-hidden="true">
      <rect width="120" height="90" fill="#fff" />
      <text x="8" y="14" fontSize="9" fontWeight="700" fill={color} fontFamily="Segoe UI, sans-serif">
        {title}
      </text>
      <rect x="8" y="22" width="104" height="8" fill={color} />
      {Array.from({ length: 6 }, (_, i) => (
        <g key={i}>
          <rect x="8" y={32 + i * 8} width="104" height="7" fill={i % 2 ? "#F2F2F2" : "#fff"} />
          <rect x="10" y={34 + i * 8} width={30 + ((i * 13) % 20)} height="3" fill="#BFBFBF" />
          <rect x="84" y={34 + i * 8} width="24" height="3" fill="#8C8C8C" />
        </g>
      ))}
      {id === "sales" && [20, 32, 26, 40].map((h, i) => <rect key={i} x={70 + i * 9} y={86 - h / 2} width="6" height={h / 2} fill={color} opacity="0.85" />)}
    </svg>
  );
}

export function FileView({ ctl, files, page, onClose }: { ctl: Controller; files: FileActions; page: string; onClose: () => void }) {
  const templates: BackstageTemplate[] = TEMPLATES.map((t) => ({ id: t.id, label: t.label, preview: <TemplatePreview id={t.id} /> }));
  const wb = ctl.wb;
  const used = wb.usedRange(ctl.sheet);
  const cellCount = wb.sheets.reduce((n, s) => n + wb.engine.cellCount(s.id), 0);
  const preview = (() => {
    const r = printRange(wb, ctl.sheet);
    if (!r) return <div className="tb-print-empty">Yazdırılacak bir şey yok.</div>;
    const limited = { ...r, r2: Math.min(r.r2, r.r1 + 40), c2: Math.min(r.c2, r.c1 + 12) };
    return <div className={`tb-print-preview is-${ctl.sheet.page.orientation}`} dangerouslySetInnerHTML={{ __html: rangeToHtml(wb, ctl.sheet, limited, { gridlines: ctl.sheet.page.printGridlines, headings: ctl.sheet.page.printHeadings, objects: true }) }} />;
  })();
  return (
    <OfficeBackstage
      app="tablo"
      fileName={ctl.fileName}
      filePath={ctl.fileRef?.path}
      templates={templates}
      templateShape="landscape"
      initialPage={page as "home"}
      onNew={(id) => {
        onClose();
        void files.newBook(id);
      }}
      onOpen={() => {
        onClose();
        void files.open();
      }}
      onOpenRecent={(f) => {
        onClose();
        void files.openRecent(f);
      }}
      onSave={() => void files.save().then((ok) => ok && onClose())}
      onSaveAs={() => void files.saveAs().then((ok) => ok && onClose())}
      onPrint={() => {
        onClose();
        setTimeout(() => files.print(), 50);
      }}
      exports={[
        { id: "xlsx", title: "Çalışma kitabını indir (.xlsx)", description: "Çalışma kitabının bir kopyasını Office Open XML biçiminde indirir.", icon: <ArrowDownload24Regular /> },
        { id: "csv", title: "CSV (Noktalı virgülle ayrılmış)", description: "Etkin sayfayı UTF-8 ve ; ayırıcıyla metin dosyası olarak kaydeder.", icon: <DocumentTable24Regular /> },
        { id: "csv-comma", title: "CSV (Virgülle ayrılmış)", description: "Etkin sayfayı virgül ayırıcıyla kaydeder.", icon: <DocumentTable24Regular /> },
        { id: "pdf", title: "PDF/XPS Belgesi Oluştur", description: "Yazdırma iletişim kutusunda “PDF olarak kaydet” seçeneğini kullanın.", icon: <DocumentPdf24Regular /> },
        { id: "html", title: "Web Sayfası (.html)", description: "Kullanılan aralığı HTML tablosu olarak dışarı aktarır.", icon: <Globe24Regular /> },
      ]}
      onExport={(id) => {
        if (id === "pdf") onClose();
        void files.exportAs(id);
      }}
      onClose={async () => {
        if (!(await files.confirmDiscard())) return;
        onClose();
        await files.newBook("blank");
        ctl.dirty = false;
      }}
      onBack={onClose}
      info={[
        ["Başlık", wb.props.title ?? "Başlık ekleyin"],
        ["Yazar", wb.props.author ?? "—"],
        ["Sayfa sayısı", String(wb.sheets.length)],
        ["Dolu hücre sayısı", cellCount.toLocaleString("tr-TR")],
        ["Kullanılan aralık", used ? `${used.r2 + 1} satır × ${used.c2 + 1} sütun` : "—"],
        ["Oluşturma tarihi", wb.props.created ? wb.props.created.toLocaleString("tr-TR") : "—"],
        ["Son değiştirme", wb.props.modified ? wb.props.modified.toLocaleString("tr-TR") : "—"],
      ]}
      printPreview={preview}
    />
  );
}
