/** Dosya görünümü (Backstage): Yeni, Aç, Kaydet, Farklı Kaydet, Yazdır, Dışarı Aktar. */
import { ArrowDownload24Regular, DocumentPdf24Regular, DocumentTable24Regular, Globe24Regular } from "@fluentui/react-icons";
import { OfficeBackstage, type BackstageTemplate } from "@officehere/ui";
import { getNative } from "@officehere/core";
import type { Controller } from "./controller";
import { printRange, rangeToHtml } from "./io/print";
import { TEMPLATES } from "./templates";
import { PrintPage } from "./PrintPage";
import type { FileActions } from "./app/fileActions";

export { useFileActions, type FileActions } from "./app/fileActions";

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
        { id: "ods", title: "OpenDocument Hesap Tablosu (.ods)", description: "Tüm sayfaları açık standart ODF biçiminde kaydeder; bazı özellikler (grafikler, özet tablolar) aktarılmaz.", icon: <DocumentTable24Regular /> },
        { id: "csv", title: "CSV (Noktalı virgülle ayrılmış)", description: "Etkin sayfayı UTF-8 ve ; ayırıcıyla metin dosyası olarak kaydeder.", icon: <DocumentTable24Regular /> },
        { id: "csv-comma", title: "CSV (Virgülle ayrılmış)", description: "Etkin sayfayı virgül ayırıcıyla kaydeder.", icon: <DocumentTable24Regular /> },
        { id: "tsv", title: "Metin (Sekmeyle ayrılmış) (.tsv)", description: "Etkin sayfanın değerlerini sekmeyle ayrılmış metin olarak kaydeder.", icon: <DocumentTable24Regular /> },
        { id: "pdf", title: "PDF/XPS Belgesi Oluştur", description: getNative()?.exportPdf ? "Etkin sayfayı Sayfa Yapısı ayarlarıyla PDF dosyası olarak kaydeder." : "Yazdırma iletişim kutusunda “PDF olarak kaydet” seçeneğini kullanın.", icon: <DocumentPdf24Regular /> },
        { id: "html", title: "Web Sayfası (.html)", description: "Etkin sayfanın kullanılan aralığını biçimleriyle HTML tablosu olarak dışarı aktarır.", icon: <Globe24Regular /> },
        { id: "html-all", title: "Web Sayfası — tüm sayfalar (.html)", description: "Görünür tüm sayfaları başlıklarıyla tek bir HTML sayfasında dışarı aktarır.", icon: <Globe24Regular /> },
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
      printPage={<PrintPage ctl={ctl} onDone={onClose} />}
    />
  );
}
