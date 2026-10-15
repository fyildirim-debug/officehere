// Dosya görünümü (Yeni, Aç, Bilgi, Kaydet, Farklı Kaydet, Yazdır, Dışarı Aktar, Kapat, Seçenekler).
import { useState } from "react";
import { DocumentPdf24Regular, Code24Regular, DocumentText24Regular, Document24Regular, DocumentOnePage24Regular, DocumentTextExtract24Regular, Markdown24Regular } from "@fluentui/react-icons";
import { getNative } from "@officehere/core";
import { OfficeBackstage, getUserName, type OfficeBackstagePage } from "@officehere/ui";
import type { DocSettings } from "../model/document";
import type { PrintOptions } from "../print/printOptions";
import { PrintPage } from "../print/PrintPage";
import { usePreviewPages } from "../print/usePreviewPages";
import { useBelge } from "../state/context";
import { applyPageSetup, sectionInfo } from "../state/pageSetup";
import type { useFileOps } from "../state/useFileOps";
import { RESMI_KUME_ID, TEMPLATES } from "../templates/templates";
import { alanlariBeklet } from "../templates/resmi/alanlar";
import { ResmiSablonPenceresi } from "../templates/resmi/ResmiSablonPenceresi";

interface Props {
  page: OfficeBackstagePage;
  files: ReturnType<typeof useFileOps>;
  settings: DocSettings;
  pageCount: number;
  stats: { words: number; chars: number };
  printOptions: PrintOptions;
  setPrintOptions: (fn: (o: PrintOptions) => PrintOptions) => void;
  autosave: boolean;
  setAutosave: (v: boolean) => void;
  spellcheck: boolean;
  setSpellcheck: (v: boolean) => void;
  onPrint: (o?: PrintOptions) => void;
  onPdf: (o?: PrintOptions) => void;
  onClose: () => void;
  /** Seçenekler sayfasından açılan iletişim kutuları. */
  onDialog: (kind: "autocorrect") => void;
}

export function BelgeBackstage({ page, files, settings, pageCount, stats, printOptions, setPrintOptions, autosave, setAutosave, spellcheck, setSpellcheck, onPrint, onPdf, onClose, onDialog }: Props) {
  const nativePdf = !!getNative()?.exportPdf;
  const ctx = useBelge();
  const ed = ctx.mainEditor;
  const preview = usePreviewPages(ctx.controller, true, settings.pageColor);
  const currentPage = ctx.controller.pageOf(ed.state.selection.from);
  const setProp = (k: "title" | "subject" | "keywords" | "author", v: string) => ctx.updateSettings((s) => ({ ...s, properties: { ...s.properties, [k]: v } }));
  // Resmî yazışma kümesi: Dosya › Yeni'de tek kutucuk; şablon galerisi ve alan formu iletişim kutusunda.
  const [resmi, setResmi] = useState(false);
  const sahne = (
    <OfficeBackstage
      app="belge"
      initialPage={page}
      fileName={files.title}
      filePath={files.fileRef?.path}
      templates={TEMPLATES.filter((t) => !t.kume).map((t) => ({ id: t.id, label: t.label, preview: t.preview }))}
      onNew={async (id) => {
        if (id === RESMI_KUME_ID) return setResmi(true);
        if (await files.newDocument(id)) onClose();
      }}
      onOpen={async () => {
        onClose();
        await files.openDialog();
      }}
      onOpenRecent={async (f) => {
        onClose();
        if (f.path !== f.name) await files.openPath(f.path);
        else await files.openDialog();
      }}
      onSave={async () => {
        if (await files.save()) onClose();
      }}
      onSaveAs={async () => {
        if (await files.saveAs()) onClose();
      }}
      onPrint={() => onPrint(printOptions)}
      printPage={
        <PrintPage
          options={printOptions}
          setOptions={setPrintOptions}
          pageCount={pageCount}
          currentPage={currentPage}
          page={sectionInfo(ed, settings.page).page}
          onPage={(fn) => applyPageSetup(ed, settings, ctx.updateSettings, fn)}
          onPageSetupDialog={() => ctx.openDialog("pageSetup")}
          onPrint={() => onPrint(printOptions)}
          onPdf={() => onPdf(printOptions)}
          nativePdf={nativePdf}
          preview={preview}
        />
      }
      exports={[
        {
          id: "pdf",
          title: "PDF Belgesi Oluştur",
          description: nativePdf ? "Belgeyi sayfa düzeni korunarak PDF dosyası olarak kaydedin." : "Yazdır iletişim kutusunda hedef olarak “PDF olarak kaydet” seçin.",
          icon: <DocumentPdf24Regular />,
        },
        { id: "docx", title: "Belge (.docx)", description: "Office Open XML belgesi olarak indirin.", icon: <Document24Regular /> },
        { id: "odt", title: "OpenDocument Metin (.odt)", description: "Açık standart belge; LibreOffice ve diğer ofis paketleriyle paylaşın.", icon: <DocumentOnePage24Regular /> },
        { id: "rtf", title: "Zengin Metin (.rtf)", description: "Biçimli metni hemen her kelime işlemcinin açabildiği biçimde kaydedin.", icon: <DocumentTextExtract24Regular /> },
        { id: "md", title: "Markdown (.md)", description: "Başlıkları, listeleri, tabloları ve bağlantıları koruyan sade metin.", icon: <Markdown24Regular /> },
        { id: "html", title: "Web Sayfası (.html)", description: "Stilleriyle birlikte tek dosyalık web sayfası.", icon: <Code24Regular /> },
        { id: "txt", title: "Düz Metin (.txt)", description: "Biçimlendirme olmadan yalnızca metin (UTF-8).", icon: <DocumentText24Regular /> },
      ]}
      onExport={async (id) => {
        if (id === "pdf") return void onPdf();
        await files.exportAs(id);
        onClose();
      }}
      onClose={async () => {
        await files.closeDocument();
        onClose();
      }}
      onBack={onClose}
      info={[
        ["Sayfalar", String(pageCount)],
        ["Sözcükler", stats.words.toLocaleString("tr-TR")],
        ["Karakterler", stats.chars.toLocaleString("tr-TR")],
        ["Başlık", <PropInput key="title" label="Başlık" value={settings.properties.title} placeholder="Başlık ekleyin" onChange={(v) => setProp("title", v)} />],
        ["Konu", <PropInput key="subject" label="Konu" value={settings.properties.subject} placeholder="Konu belirtin" onChange={(v) => setProp("subject", v)} />],
        ["Etiketler", <PropInput key="keywords" label="Etiketler" value={settings.properties.keywords} placeholder="Etiket ekleyin" onChange={(v) => setProp("keywords", v)} />],
        ["Yazar", <PropInput key="author" label="Yazar" value={settings.properties.author} placeholder={getUserName()} onChange={(v) => setProp("author", v)} />],
        ["Oluşturulma", new Date(settings.properties.created).toLocaleString("tr-TR")],
        ["Son Değiştirme", new Date(settings.properties.modified).toLocaleString("tr-TR")],
      ]}
      extraOptions={
        <>
          <label className="oh-field is-inline">
            <span className="oh-field-label">Otomatik kurtarma bilgilerini her 5 saniyede bir kaydet</span>
            <input type="checkbox" checked={autosave} onChange={(e) => setAutosave(e.target.checked)} />
          </label>
          <label className="oh-field is-inline">
            <span className="oh-field-label">Yazarken yazım denetimi yap</span>
            <input type="checkbox" checked={spellcheck} onChange={(e) => setSpellcheck(e.target.checked)} />
          </label>
          <div className="oh-field is-inline">
            <span className="oh-field-label">Yazdıkça metnin nasıl düzeltileceğini ve biçimlendirileceğini değiştirin</span>
            <button type="button" className="oh-btn" onClick={() => onDialog("autocorrect")}>
              Otomatik Düzeltme Seçenekleri...
            </button>
          </div>
        </>
      }
    />
  );
  return (
    <>
      {sahne}
      {resmi && (
        <ResmiSablonPenceresi
          onClose={() => setResmi(false)}
          onOlustur={async (sablon, alanlar) => {
            alanlariBeklet(sablon, alanlar);
            setResmi(false);
            if (await files.newDocument(sablon)) onClose();
          }}
        />
      )}
    </>
  );
}

/** Bilgi sayfasında düzenlenebilir belge özelliği (docProps/core.xml). */
function PropInput({ label, value, placeholder, onChange }: { label: string; value: string; placeholder: string; onChange: (v: string) => void }) {
  return <input type="text" className="bl-prop-input" aria-label={label} value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />;
}
