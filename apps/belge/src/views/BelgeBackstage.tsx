// Dosya görünümü (Yeni, Aç, Bilgi, Kaydet, Farklı Kaydet, Yazdır, Dışarı Aktar, Kapat, Seçenekler).
import { DocumentPdf24Regular, Code24Regular, DocumentText24Regular, Document24Regular } from "@fluentui/react-icons";
import { getNative } from "@officehere/core";
import { OfficeBackstage, getUserName, type OfficeBackstagePage } from "@officehere/ui";
import type { DocSettings } from "../model/document";
import type { useFileOps } from "../state/useFileOps";
import { TEMPLATES } from "../templates/templates";

interface Props {
  page: OfficeBackstagePage;
  files: ReturnType<typeof useFileOps>;
  settings: DocSettings;
  pageCount: number;
  stats: { words: number; chars: number };
  previewHtml: string;
  autosave: boolean;
  setAutosave: (v: boolean) => void;
  spellcheck: boolean;
  setSpellcheck: (v: boolean) => void;
  onPrint: () => void;
  onPdf: () => void;
  onClose: () => void;
}

export function BelgeBackstage({ page, files, settings, pageCount, stats, previewHtml, autosave, setAutosave, spellcheck, setSpellcheck, onPrint, onPdf, onClose }: Props) {
  const nativePdf = !!getNative()?.exportPdf;
  return (
    <OfficeBackstage
      app="belge"
      initialPage={page}
      fileName={files.fileName}
      filePath={files.fileRef?.path}
      templates={TEMPLATES.map((t) => ({ id: t.id, label: t.label, preview: t.preview }))}
      onNew={async (id) => {
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
      onPrint={onPrint}
      exports={[
        {
          id: "pdf",
          title: "PDF Belgesi Oluştur",
          description: nativePdf ? "Belgeyi sayfa düzeni korunarak PDF dosyası olarak kaydedin." : "Yazdır iletişim kutusunda hedef olarak “PDF olarak kaydet” seçin.",
          icon: <DocumentPdf24Regular />,
        },
        { id: "docx", title: "Belge (.docx)", description: "Office Open XML belgesi olarak indirin.", icon: <Document24Regular /> },
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
        ["Başlık", settings.properties.title || "Başlık ekleyin"],
        ["Yazar", settings.properties.author || getUserName()],
        ["Oluşturulma", new Date(settings.properties.created).toLocaleString("tr-TR")],
        ["Son Değiştirme", new Date(settings.properties.modified).toLocaleString("tr-TR")],
      ]}
      printPreview={<div className="bl-print-preview bl-doc" dangerouslySetInnerHTML={{ __html: previewHtml }} />}
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
        </>
      }
    />
  );
}
