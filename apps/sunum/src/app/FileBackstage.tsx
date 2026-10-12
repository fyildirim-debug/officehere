/** Dosya menüsü: şablonlar, dışa aktarma ve Bilgi özellikleri. */
import { ArrowDownload24Regular, DocumentPdf24Regular, Image24Regular, ImageMultiple24Regular, TextBulletListLtr24Regular } from "@fluentui/react-icons";
import { useLayoutEffect, useRef, useState } from "react";
import { OfficeBackstage, type BackstageExport } from "@officehere/ui";
import type { RecentFile } from "@officehere/core";
import { formatDateTime, isDesktop } from "@officehere/core";
import type { Presentation } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { TEMPLATES } from "../templates/templates";
import { cm } from "../ribbon/common";

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
  { id: "outline", title: "Ana Hat Olarak Kaydet", description: "Slayt başlıkları ve metinleri düz metin dosyası olarak.", icon: <TextBulletListLtr24Regular /> },
  { id: "pptx", title: "Sunu (.pptx) Olarak Kaydet", description: "Açık XML sunu biçiminde farklı kaydet.", icon: <ArrowDownload24Regular /> },
];

export function FileBackstage(props: {
  state: AppState;
  actions: Actions;
  onBack: () => void;
  onNew: (id: string) => void;
  onOpen: () => void;
  onOpenRecent: (f: RecentFile) => void;
  onSave: () => void;
  onSaveAs: () => void;
  onPrint: () => void;
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
      printPreview={pres.slides[0] ? <PrintPreview pres={pres} start={state.slideIdx} /> : undefined}
    />
  );
}

/** Yazdır sayfasındaki baskı önizlemesi: bir sayfa, altta sayfa gezinmesi (gizli slaytlar basılmaz). */
function PrintPreview({ pres, start }: { pres: Presentation; start: number }) {
  const pages = pres.slides.map((s, i) => ({ s, i })).filter(({ s }) => !s.hidden);
  const ref = useRef<HTMLDivElement>(null);
  const [w, setW] = useState(480);
  const [page, setPage] = useState(() => Math.max(0, pages.findIndex((p) => p.i >= start)));
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setW(Math.max(160, Math.min(900, el.clientWidth - 2))));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  if (!pages.length) return <div className="print-prev-empty">Yazdırılacak slayt yok (tüm slaytlar gizli).</div>;
  const cur = pages[Math.min(page, pages.length - 1)];
  return (
    <div className="print-prev" ref={ref}>
      <div className="thumb-box" style={{ width: w }}>
        <ScaledSlide pres={pres} slide={cur.s} slideNumber={cur.i + 1} width={w} />
      </div>
      <div className="print-prev-nav">
        <button type="button" aria-label="Önceki Sayfa" disabled={page <= 0} onClick={() => setPage((p) => Math.max(0, p - 1))}>
          ◀
        </button>
        <span>
          {Math.min(page, pages.length - 1) + 1} / {pages.length}
        </span>
        <button type="button" aria-label="Sonraki Sayfa" disabled={page >= pages.length - 1} onClick={() => setPage((p) => Math.min(pages.length - 1, p + 1))}>
          ▶
        </button>
      </div>
    </div>
  );
}
