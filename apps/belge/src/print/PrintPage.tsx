// Dosya › Yazdır sayfası: kopya sayısı, yazıcı, ayarlar (sayfa aralığı, tek/elle iki taraflı, harmanlama,
// yönlendirme, kağıt boyutu, kenar boşlukları, yaprak başına sayfa) ve sağda gezinti ve yakınlaştırmalı
// gerçek baskı önizlemesi.
import { useRef, useState, type ReactElement, type ReactNode } from "react";
import {
  ChevronDown16Regular,
  DocumentLandscape24Regular,
  DocumentMargins24Regular,
  DocumentMultiple24Regular,
  DocumentOnePage24Regular,
  DocumentPdf24Regular,
  Orientation24Regular,
  Print24Regular,
  Print32Regular,
  DocumentCopy24Regular,
  Info16Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuProvider, Popover, BackstageHeading } from "@officehere/ui";
import { MARGIN_PRESETS, PAGE_SIZE_LIST, fmtCm, marginPresetId, pageSizeId, type PageSetup } from "../model/document";
import { setOrientation, setPaperSize } from "../ribbon/LayoutTab";
import { PER_SHEET, selectedPages, type PrintOptions } from "./printOptions";
import { PrintPreview } from "./PrintPreview";
import type { PageInfo } from "../editor/pagination";

interface Props {
  options: PrintOptions;
  setOptions: (fn: (o: PrintOptions) => PrintOptions) => void;
  pageCount: number;
  currentPage: number;
  /** İmleçteki bölümün sayfa yapısı. */
  page: PageSetup;
  onPage: (fn: (p: PageSetup) => PageSetup) => void;
  onPageSetupDialog: () => void;
  onPrint: () => void;
  onPdf: () => void;
  nativePdf: boolean;
  preview: { pages: HTMLElement[]; infos: PageInfo[] } | null;
}

interface Choice<T> {
  value: T;
  title: string;
  sub?: string;
  icon?: ReactElement;
}

/** Word'deki iki satırlı açılır ayar düğmesi. */
function PrintOption<T>({ label, value, choices, onPick, icon, extra }: { label: string; value: T; choices: Choice<T>[]; onPick: (v: T) => void; icon?: ReactElement; extra?: ReactNode }) {
  const ref = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const cur = choices.find((c) => c.value === value) ?? choices[0];
  return (
    <>
      <button ref={ref} type="button" className="bl-pr-opt" aria-label={label} aria-haspopup="menu" aria-expanded={open} onClick={() => setOpen((o) => !o)}>
        <span className="bl-pr-opt-icon">{cur?.icon ?? icon}</span>
        <span className="bl-pr-opt-text">
          <strong>{cur?.title}</strong>
          {cur?.sub && <span>{cur.sub}</span>}
        </span>
        <ChevronDown16Regular />
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={ref} role="menu">
        <MenuProvider onClose={() => setOpen(false)}>
          <Menu width={ref.current?.offsetWidth ?? 300}>
            {choices.map((c) => (
              <MenuItem key={String(c.value)} icon={c.icon} label={c.title} description={c.sub} checked={c.value === value} onClick={() => onPick(c.value)} />
            ))}
            {extra}
          </Menu>
        </MenuProvider>
      </Popover>
    </>
  );
}

export function PrintPage(p: Props) {
  const { options: o, setOptions, page } = p;
  const set = (patch: Partial<PrintOptions>) => setOptions((x) => ({ ...x, ...patch }));
  const sel = selectedPages(o, p.pageCount, p.currentPage);
  const error = typeof sel === "string" && o.range === "custom" && o.pages.trim() ? sel : null;
  const mp = marginPresetId(page);
  const sz = pageSizeId(page);
  const m = page.margin;

  return (
    <>
      <BackstageHeading>Yazdır</BackstageHeading>
      <div className="bl-pr">
        <div className="bl-pr-side">
          <div className="bl-pr-top">
            <button type="button" className="bl-pr-print" onClick={p.onPrint} disabled={!!error || (typeof sel === "string" && o.range === "custom")}>
              <Print32Regular />
              <span>Yazdır</span>
            </button>
            <label className="bl-pr-copies">
              <span>Kopyalar:</span>
              <input type="number" aria-label="Kopyalar" min={1} max={999} value={o.copies} onChange={(e) => set({ copies: Math.max(1, Math.min(999, Number(e.target.value) || 1)) })} />
            </label>
          </div>

          <h3 className="bl-pr-h">Yazıcı</h3>
          <div className="bl-pr-printer" title="Yazıcı, sistem yazdırma iletişim kutusunda seçilir.">
            <Print24Regular />
            <span className="bl-pr-opt-text">
              <strong>Sistem yazıcısı</strong>
              <span>Hazır</span>
            </span>
          </div>
          <p className="bl-pr-note">
            <Info16Regular /> Yazıcı ve yazıcı özellikleri, Yazdır'a bastıktan sonra açılan sistem iletişim kutusunda seçilir.
          </p>
          {p.nativePdf ? (
            <button type="button" className="bl-pr-link" onClick={p.onPdf}>
              <DocumentPdf24Regular /> PDF olarak kaydet
            </button>
          ) : (
            <p className="bl-pr-note">PDF dosyası için yazıcı olarak “PDF olarak kaydet” seçeneğini kullanın.</p>
          )}

          <h3 className="bl-pr-h">Ayarlar</h3>
          <PrintOption<PrintOptions["range"]>
            label="Yazdırılacak sayfalar"
            value={o.range}
            onPick={(v) => set({ range: v })}
            choices={[
              { value: "all", title: "Tüm Sayfaları Yazdır", sub: "Tüm belge", icon: <DocumentMultiple24Regular /> },
              { value: "current", title: "Geçerli Sayfayı Yazdır", sub: "Yalnızca bu sayfa", icon: <DocumentOnePage24Regular /> },
              { value: "custom", title: "Özel Yazdırma", sub: "Belirli sayfaları, bölümleri ya da aralıkları yazın", icon: <DocumentCopy24Regular /> },
            ]}
          />
          <label className="bl-pr-pages">
            <span>Sayfalar:</span>
            <input
              type="text"
              aria-label="Sayfalar"
              placeholder="ör. 1-3, 5"
              value={o.pages}
              aria-invalid={!!error}
              onChange={(e) => set({ pages: e.target.value, range: e.target.value.trim() ? "custom" : o.range === "custom" ? "all" : o.range })}
            />
          </label>
          {error && <p className="bl-pr-error">{error}</p>}
          <PrintOption<PrintOptions["duplex"]>
            label="Tek ya da iki taraflı"
            value={o.duplex}
            onPick={(v) => set({ duplex: v })}
            choices={[
              { value: "single", title: "Tek Taraflı Yazdır", sub: "Yalnızca sayfanın bir yüzüne yazdırın" },
              { value: "manual", title: "Elle İki Taraflı Yazdır", sub: "İkinci yüzü yazdırmak için istendiğinde kağıdı yeniden yerleştirin" },
            ]}
          />
          <PrintOption
            label="Harmanlama"
            value={o.collate}
            onPick={(v) => set({ collate: v })}
            choices={[
              { value: true, title: "Harmanlanmış", sub: "1,2,3 1,2,3 1,2,3" },
              { value: false, title: "Harmanlanmamış", sub: "1,1,1 2,2,2 3,3,3" },
            ]}
          />
          <PrintOption<PageSetup["orientation"]>
            label="Yönlendirme"
            value={page.orientation}
            onPick={(v) => p.onPage((x) => setOrientation(x, v))}
            choices={[
              { value: "portrait", title: "Dikey Yönlendirme", icon: <Orientation24Regular /> },
              { value: "landscape", title: "Yatay Yönlendirme", icon: <DocumentLandscape24Regular /> },
            ]}
          />
          <PrintOption
            label="Kağıt boyutu"
            value={sz}
            onPick={(v) => p.onPage((x) => setPaperSize(x, v))}
            choices={[
              ...PAGE_SIZE_LIST.map((s) => ({ value: s.id, title: s.label, sub: `${fmtCm(s.w, 1)} x ${fmtCm(s.h, 1)}` })),
              ...(sz === "custom" ? [{ value: "custom", title: "Özel boyut", sub: `${fmtCm(page.width, 1)} x ${fmtCm(page.height, 1)}` }] : []),
            ]}
          />
          <PrintOption
            label="Kenar boşlukları"
            value={mp}
            icon={<DocumentMargins24Regular />}
            onPick={(v) => {
              const preset = MARGIN_PRESETS.find((x) => x.id === v);
              if (preset) p.onPage((x) => ({ ...x, margin: { ...x.margin, ...preset.m } }));
            }}
            choices={[
              ...MARGIN_PRESETS.map((x) => ({ value: x.id, title: `${x.label} Kenar Boşlukları`, sub: `Sol: ${fmtCm(x.m.left)}   Sağ: ${fmtCm(x.m.right)}`, icon: <DocumentMargins24Regular /> })),
              ...(mp === "custom" ? [{ value: "custom", title: "Özel Kenar Boşlukları", sub: `Sol: ${fmtCm(m.left)}   Sağ: ${fmtCm(m.right)}`, icon: <DocumentMargins24Regular /> }] : []),
            ]}
          />
          <PrintOption
            label="Yaprak başına sayfa"
            value={o.perSheet}
            onPick={(v) => set({ perSheet: v })}
            choices={PER_SHEET.map((n) => ({ value: n, title: `Yaprak Başına ${n} Sayfa`, icon: <DocumentMultiple24Regular /> }))}
          />
          <button type="button" className="bl-pr-link" onClick={p.onPageSetupDialog}>
            Sayfa Yapısı
          </button>
        </div>
        <PrintPreview preview={p.preview} options={o} pageCount={p.pageCount} currentPage={p.currentPage} />
      </div>
    </>
  );
}
