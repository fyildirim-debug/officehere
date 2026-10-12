/**
 * Dosya › Yazdır sayfası: kopya sayısı, yazdırılacak kapsam (etkin sayfalar,
 * tüm çalışma kitabı, seçim), sayfa aralığı, yönlendirme, kağıt boyutu,
 * kenar boşlukları ve ölçek; sağda sayfa sayfa gezinilen baskı önizlemesi.
 */
import { useMemo, useState } from "react";
import { ChevronLeft16Regular, ChevronRight16Regular, Print24Regular } from "@fluentui/react-icons";
import { BackstageHeading } from "@officehere/ui";
import type { Controller } from "./controller";
import type { PageSetup } from "./model/types";
import { setPage } from "./dialogs/PageSetupDialog";
import { fitModeOf } from "./io/print";
import { pageHtml, paperBox, printPages, printWithOptions, type PrintScope } from "./io/printPages";

const PAPERS: { id: PageSetup["paper"]; label: string }[] = [
  { id: "A4", label: "A4 (21 cm x 29,7 cm)" },
  { id: "A3", label: "A3 (29,7 cm x 42 cm)" },
  { id: "A5", label: "A5 (14,8 cm x 21 cm)" },
  { id: "Letter", label: "Letter (21,59 cm x 27,94 cm)" },
  { id: "Legal", label: "Legal (21,59 cm x 35,56 cm)" },
];

const MARGINS: { id: string; label: string; m: PageSetup["margins"] }[] = [
  { id: "normal", label: "Normal Kenar Boşlukları", m: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  { id: "wide", label: "Geniş Kenar Boşlukları", m: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  { id: "narrow", label: "Dar Kenar Boşlukları", m: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
];

const SCALES: { id: "none" | "sheet" | "cols" | "rows"; label: string }[] = [
  { id: "none", label: "Ölçeklendirme Yok" },
  { id: "sheet", label: "Sayfayı Tek Sayfaya Sığdır" },
  { id: "cols", label: "Tüm Sütunları Tek Sayfaya Sığdır" },
  { id: "rows", label: "Tüm Satırları Tek Sayfaya Sığdır" },
];

function marginId(m: PageSetup["margins"]): string {
  return MARGINS.find((x) => x.m.top === m.top && x.m.left === m.left && x.m.right === m.right && x.m.bottom === m.bottom)?.id ?? "custom";
}

/** Önizleme kutusunun yüksekliği (px). */
const PREVIEW_H = 560;

export function PrintPage({ ctl, onDone }: { ctl: Controller; onDone: () => void }) {
  const sheet = ctl.sheet;
  const p = sheet.page;
  const [copies, setCopies] = useState(1);
  const [scope, setScope] = useState<PrintScope>("active");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [index, setIndex] = useState(0);
  const sel = ctl.range;
  const pages = useMemo(() => printPages(ctl.wb, sheet, { scope, selection: sel }), [ctl.wb, sheet, scope, sel, ctl.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const total = pages.length;
  const cur = pages[Math.min(index, Math.max(0, total - 1))];
  const num = (s: string) => (s.trim() === "" ? undefined : Math.max(1, Math.floor(Number(s)) || 1));
  const print = () => {
    onDone();
    setTimeout(() => printWithOptions(ctl.wb, sheet, { scope, selection: sel, copies, from: num(from), to: num(to) }), 50);
  };
  const fit = fitModeOf(sheet) ?? "none";

  // Önizleme: kağıt oranında kutu, içerik kenar boşluklarıyla ve ölçekle
  const box = paperBox(cur?.sheet ?? sheet);
  const k = PREVIEW_H / box.h;
  return (
    <>
      <BackstageHeading>Yazdır</BackstageHeading>
      <div className="tb-printpage">
        <div className="tb-printpage-side">
          <div className="tb-printpage-top">
            <button type="button" className="tb-printpage-btn" onClick={print} aria-label="Yazdır">
              <Print24Regular />
              <span>Yazdır</span>
            </button>
            <label className="tb-printpage-copies">
              Kopya:
              <input type="number" min={1} max={99} value={copies} onChange={(e) => setCopies(Math.max(1, Math.min(99, Number(e.target.value) || 1)))} aria-label="Kopya sayısı" />
            </label>
          </div>
          <div className="tb-printpage-label">Yazıcı</div>
          <select disabled aria-label="Yazıcı">
            <option>Sistem yazıcısı (yazdırma penceresinde seçilir)</option>
          </select>
          <div className="tb-printpage-label">Ayarlar</div>
          <select value={scope} onChange={(e) => (setScope(e.target.value as PrintScope), setIndex(0))} aria-label="Yazdırılacak">
            <option value="active">Etkin Sayfaları Yazdır</option>
            <option value="workbook">Tüm Çalışma Kitabını Yazdır</option>
            <option value="selection">Seçimi Yazdır</option>
          </select>
          <div className="tb-printpage-range">
            <span>Sayfalar:</span>
            <input type="number" min={1} value={from} onChange={(e) => setFrom(e.target.value)} aria-label="İlk sayfa" />
            <span>-</span>
            <input type="number" min={1} value={to} onChange={(e) => setTo(e.target.value)} aria-label="Son sayfa" />
          </div>
          <select value={p.orientation} onChange={(e) => setPage(ctl, { orientation: e.target.value as PageSetup["orientation"] }, "Yönlendirme")} aria-label="Yönlendirme">
            <option value="portrait">Dikey Yönlendirme</option>
            <option value="landscape">Yatay Yönlendirme</option>
          </select>
          <select value={p.paper} onChange={(e) => setPage(ctl, { paper: e.target.value as PageSetup["paper"] }, "Kağıt Boyutu")} aria-label="Kağıt boyutu">
            {PAPERS.map((x) => (
              <option key={x.id} value={x.id}>
                {x.label}
              </option>
            ))}
          </select>
          <select
            value={marginId(p.margins)}
            onChange={(e) => {
              const m = MARGINS.find((x) => x.id === e.target.value);
              if (m) setPage(ctl, { margins: { ...m.m } }, "Kenar Boşlukları");
              else ctl.openDialog("pageSetup", { tab: "margins" });
            }}
            aria-label="Kenar boşlukları"
          >
            {MARGINS.map((x) => (
              <option key={x.id} value={x.id}>
                {x.label}
              </option>
            ))}
            <option value="custom">Özel Kenar Boşlukları...</option>
          </select>
          <select value={fit} onChange={(e) => setPage(ctl, { fitToWidth: undefined, fitMode: e.target.value === "none" ? undefined : (e.target.value as "sheet") }, "Ölçek")} aria-label="Ölçek">
            {SCALES.map((x) => (
              <option key={x.id} value={x.id}>
                {x.label}
              </option>
            ))}
          </select>
          <button type="button" className="tb-printpage-link" onClick={() => ctl.openDialog("pageSetup")}>
            Sayfa Yapısı
          </button>
        </div>
        <div className="tb-printpage-preview">
          {!cur ? (
            <div className="tb-print-empty">Yazdırılacak bir şey yok.</div>
          ) : (
            <div className="tb-printpage-paper" style={{ width: box.w * k, height: box.h * k }} aria-label={`Sayfa ${index + 1} önizlemesi`}>
              <div
                className="tb-printpage-content"
                style={{ left: box.margin.left * k, top: box.margin.top * k, width: box.inner.w / (cur.scale || 1), height: box.inner.h / (cur.scale || 1), transform: `scale(${k * cur.scale})` }}
                dangerouslySetInnerHTML={{ __html: pageHtml(ctl.wb, cur) }}
              />
            </div>
          )}
          <div className="tb-printpage-nav">
            <button type="button" aria-label="Önceki sayfa" disabled={index <= 0} onClick={() => setIndex((i) => Math.max(0, i - 1))}>
              <ChevronLeft16Regular />
            </button>
            <span>
              {total ? Math.min(index + 1, total) : 0} / {total}
            </span>
            <button type="button" aria-label="Sonraki sayfa" disabled={index >= total - 1} onClick={() => setIndex((i) => Math.min(total - 1, i + 1))}>
              <ChevronRight16Regular />
            </button>
          </div>
        </div>
      </div>
    </>
  );
}
