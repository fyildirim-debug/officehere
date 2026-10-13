/** Dosya menüsü: şablonlar, dışa aktarma ve Bilgi özellikleri. */
import { ArrowDownload24Regular, DocumentPdf24Regular, DocumentSave24Regular, Image24Regular, ImageMultiple24Regular, SlidePlay24Regular, SlideTextMultiple24Regular, TextBulletListLtr24Regular, Video24Regular } from "@fluentui/react-icons";
import { OfficeBackstage, type BackstageExport, type OfficeBackstagePage } from "@officehere/ui";
import type { RecentFile } from "@officehere/core";
import { formatDateTime, isDesktop } from "@officehere/core";
import type { Presentation } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { TEMPLATES } from "../templates/templates";
import { cm } from "../ribbon/common";
import { PrintPage } from "./PrintPage";
import type { PrintJob } from "./printLayout";

const built = new Map<string, Presentation>();
function templatePres(id: string): Presentation {
  let p = built.get(id);
  if (!p) {
    p = TEMPLATES.find((t) => t.id === id)!.build();
    built.set(id, p);
  }
  return p;
}

const EXPORTS: BackstageExport[] = [
  {
    id: "pdf",
    title: "PDF Belgesi Oluştur",
    description: isDesktop() ? "Her slayt yatay bir sayfa olarak PDF dosyasına kaydedilir." : "Her slayt yatay bir sayfa olarak yazdırılır; yazıcı olarak “PDF olarak kaydet” seçin.",
    icon: <DocumentPdf24Regular />,
  },
  { id: "png", title: "Geçerli Slaytı Resim Olarak Kaydet", description: "Geçerli slaydı PNG resmi olarak indirir.", icon: <Image24Regular /> },
  { id: "pngAll", title: "Tüm Slaytları Resim Olarak Kaydet", description: "Her slaydı PNG olarak bir ZIP arşivinde indirir.", icon: <ImageMultiple24Regular /> },
  { id: "video", title: "Video Oluştur", description: "Sunuyu geçişler, animasyon adımları ve zamanlamalarla bir video dosyasına (.webm) dönüştürün.", icon: <Video24Regular /> },
  { id: "outline", title: "Ana Hat Olarak Kaydet", description: "Slayt başlıkları ve metinleri düz metin dosyası olarak.", icon: <TextBulletListLtr24Regular /> },
  { id: "pptx", title: "Sunu (.pptx) Olarak Kaydet", description: "Açık XML sunu biçiminde farklı kaydet.", icon: <ArrowDownload24Regular /> },
  { id: "odp", title: "OpenDocument Sunu (.odp) Olarak Kaydet", description: "Açık belge biçiminde kaydedin; bazı efektler ve grafikler aktarılmayabilir.", icon: <DocumentSave24Regular /> },
  { id: "ppsx", title: "Slayt Gösterisi (.ppsx) Olarak Kaydet", description: "Açılınca doğrudan gösteri olarak oynatılacak bir dosya.", icon: <SlidePlay24Regular /> },
  { id: "potx", title: "Şablon (.potx) Olarak Kaydet", description: "Yeni sunular için başlangıç noktası olacak bir şablon.", icon: <SlideTextMultiple24Regular /> },
];

export function FileBackstage(props: {
  initialPage?: OfficeBackstagePage;
  state: AppState;
  actions: Actions;
  onBack: () => void;
  onNew: (id: string) => void;
  onOpen: () => void;
  onOpenRecent: (f: RecentFile) => void;
  onSave: () => void;
  onSaveAs: () => void;
  onPrint: () => void;
  /** Yazdır sayfasındaki ayarlarla yazdır / PDF olarak kaydet. */
  onPrintJob: (job: PrintJob) => void;
  /** Not ve dinleyici notu üst/alt bilgisi (Üst Bilgi ve Alt Bilgi iletişim kutusu). */
  onHeaderFooter: () => void;
  onExport: (id: string) => void;
  onClose: () => void;
}) {
  const { state } = props;
  const pres = state.pres;
  const templates = TEMPLATES.map((t) => {
    const p = templatePres(t.id);
    return { id: t.id, label: t.label, preview: <ScaledSlide pres={p} slide={p.slides[0]} slideNumber={1} width={194} /> };
  });
  const words = pres.slides.reduce((n, s) => n + s.shapes.reduce((m, sh) => m + (sh.kind === "sp" && sh.text ? sh.text.paragraphs.reduce((k, p) => k + p.runs.map((r) => r.text).join("").split(/\s+/).filter(Boolean).length, 0) : 0), 0), 0);
  const info: [string, React.ReactNode][] = [
    ["Slaytlar", String(pres.slides.length)],
    ["Gizli slaytlar", String(pres.slides.filter((s) => s.hidden).length)],
    ["Sözcükler", String(words)],
    ["Notlar", String(pres.slides.filter((s) => s.notes.trim()).length)],
    ["Slayt boyutu", `${cm(pres.width).toLocaleString("tr-TR")} × ${cm(pres.height).toLocaleString("tr-TR")} cm`],
    ["Tema", pres.theme.name],
    ["Başlık", pres.props.title || state.fileName],
    ["Yazar", pres.props.creator || "—"],
    ["Oluşturma tarihi", pres.props.created ? formatDateTime(Date.parse(pres.props.created)) : "—"],
  ];
  return (
    <OfficeBackstage
      app="sunum"
      initialPage={props.initialPage}
      fileName={state.fileName}
      filePath={state.fileRef?.path}
      templates={templates}
      templateShape="landscape"
      onNew={props.onNew}
      onOpen={props.onOpen}
      onOpenRecent={props.onOpenRecent}
      onSave={props.onSave}
      onSaveAs={props.onSaveAs}
      onPrint={props.onPrint}
      exports={EXPORTS}
      onExport={props.onExport}
      onClose={props.onClose}
      onBack={props.onBack}
      info={info}
      printPage={<PrintPage state={state} actions={props.actions} onPrint={props.onPrintJob} onHeaderFooter={props.onHeaderFooter} />}
      extraOptions={
        <>
          <label className="oh-field is-inline">
            <span className="oh-field-label">Yazarken yazımı denetle</span>
            <input type="checkbox" checked={state.spellcheck} onChange={(e) => props.actions.store.set({ spellcheck: e.target.checked })} />
          </label>
          <label className="oh-field is-inline">
            <span className="oh-field-label">Not bölmesini göster</span>
            <input type="checkbox" checked={state.showNotes} onChange={(e) => props.actions.store.set({ showNotes: e.target.checked })} />
          </label>
        </>
      }
    />
  );
}
