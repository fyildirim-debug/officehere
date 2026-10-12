/**
 * Grafik öğesini biçimlendirme bölmesi (grafiğin bir öğesine çift tıklayınca
 * açılır): Grafik Alanı, Çizim Alanı, Grafik Başlığı, Gösterge, Değer/Kategori
 * Ekseni, Veri Serisi, Veri Etiketleri. Üstteki listeden öğe değiştirilebilir.
 * Değişiklikler anında uygulanır (geri alınabilir).
 */
import { useEffect, useState } from "react";
import { Dismiss16Regular } from "@fluentui/react-icons";
import { ColorPalette } from "@officehere/ui";
import type { Controller } from "../controller";
import type { ChartDef } from "../model/types";
import { updateChart } from "../actions/objects";
import { paletteOf } from "../model/theme";
import { PALETTE_LABELS } from "./chartData";
import { seriesColor } from "./ChartView";
import { extractChartData } from "./chartData";

const PX_PER_CM = 96 / 2.54;
const LEGEND_POS: { v: ChartDef["legendPos"]; label: string }[] = [
  { v: "top", label: "Üst" },
  { v: "bottom", label: "Alt" },
  { v: "left", label: "Sol" },
  { v: "right", label: "Sağ" },
];

function Section({ title, children, open = true }: { title: string; children: React.ReactNode; open?: boolean }) {
  return (
    <details className="tb-fmt-sec" open={open}>
      <summary>{title}</summary>
      <div className="tb-fmt-body">{children}</div>
    </details>
  );
}

/** Odak kaybında uygulanan metin alanı. */
function TextField({ label, value, onCommit }: { label: string; value: string; onCommit: (v: string) => void }) {
  const [v, setV] = useState(value);
  useEffect(() => setV(value), [value]);
  return (
    <label className="tb-fmt-row">
      <span>{label}</span>
      <input type="text" value={v} onChange={(e) => setV(e.target.value)} onBlur={() => v !== value && onCommit(v)} onKeyDown={(e) => e.key === "Enter" && (e.currentTarget as HTMLInputElement).blur()} />
    </label>
  );
}

function NumField({ label, value, onCommit }: { label: string; value: number; onCommit: (v: number) => void }) {
  const [v, setV] = useState(value.toLocaleString("tr-TR", { maximumFractionDigits: 2 }));
  useEffect(() => setV(value.toLocaleString("tr-TR", { maximumFractionDigits: 2 })), [value]);
  const commit = () => {
    const n = Number(v.replace(",", "."));
    if (Number.isFinite(n) && n > 0.2 && n !== value) onCommit(n);
  };
  return (
    <label className="tb-fmt-row">
      <span>{label}</span>
      <input type="text" inputMode="decimal" value={v} onChange={(e) => setV(e.target.value)} onBlur={commit} onKeyDown={(e) => e.key === "Enter" && commit()} />
      <span className="tb-fmt-unit">cm</span>
    </label>
  );
}

/** Yazılabilir sayı alanı (boşsa Otomatik). */
function OptNum({ label, value, onCommit }: { label: string; value: number | undefined; onCommit: (v: number | undefined) => void }) {
  const show = (v: number | undefined) => (v === undefined ? "" : v.toLocaleString("tr-TR", { maximumFractionDigits: 6, useGrouping: false }));
  const [v, setV] = useState(show(value));
  useEffect(() => setV(show(value)), [value]);
  const commit = () => {
    const t = v.trim();
    if (!t) return value !== undefined && onCommit(undefined);
    const n = Number(t.replace(/\./g, "").replace(",", "."));
    if (Number.isFinite(n) && n !== value) onCommit(n);
    else setV(show(value));
  };
  return (
    <label className="tb-fmt-row">
      <span>{label}</span>
      <input type="text" inputMode="decimal" placeholder="Otomatik" value={v} onChange={(e) => setV(e.target.value)} onBlur={commit} onKeyDown={(e) => e.key === "Enter" && commit()} aria-label={label} />
    </label>
  );
}

const AXIS_FORMATS: { code: string; label: string }[] = [
  { code: "", label: "Kaynağa bağlı" },
  { code: "General", label: "Genel" },
  { code: "#,##0", label: "Sayı (1.234)" },
  { code: "#,##0.00", label: "Sayı (1.234,56)" },
  { code: '#,##0 "₺"', label: "Para Birimi (₺)" },
  { code: "0%", label: "Yüzde" },
  { code: "dd.mm.yyyy", label: "Tarih" },
];

const LABEL_POS: { v: NonNullable<ChartDef["labelPos"]>; label: string }[] = [
  { v: "ctr", label: "Orta" },
  { v: "inEnd", label: "İç Uç" },
  { v: "inBase", label: "İç Taban" },
  { v: "outEnd", label: "Dış Uç" },
];

/** Biçimlendirilebilir grafik öğeleri (seçim listesi). */
export function chartElements(ctl: Controller, chart: ChartDef): { id: string; label: string; title: string }[] {
  const isPie = chart.type === "pie" || chart.type === "doughnut";
  const out = [
    { id: "area", label: "Grafik Alanı", title: "Grafik Alanını Biçimlendir" },
    { id: "plot", label: "Çizim Alanı", title: "Çizim Alanını Biçimlendir" },
    { id: "title", label: "Grafik Başlığı", title: "Grafik Başlığını Biçimlendir" },
    { id: "legend", label: "Gösterge", title: "Göstergeyi Biçimlendir" },
  ];
  if (!isPie) out.push({ id: "axis:val", label: "Dikey (Değer) Eksen", title: "Ekseni Biçimlendir" }, { id: "axis:cat", label: "Yatay (Kategori) Eksen", title: "Ekseni Biçimlendir" });
  const data = extractChartData(ctl.wb, chart);
  data.series.forEach((s, i) => out.push({ id: `series:${i}`, label: `Seri "${s.name}"`, title: "Veri Serisini Biçimlendir" }));
  out.push({ id: "labels", label: "Veri Etiketleri", title: "Veri Etiketlerini Biçimlendir" });
  return out;
}

export function ChartFormatPane({ ctl, chart }: { ctl: Controller; chart: ChartDef }) {
  const set = (patch: Partial<ChartDef>, label = "Grafik Biçimi") => updateChart(ctl, chart.id, patch, label);
  const close = () => {
    ctl.formatPane = null;
    ctl.emit();
    ctl.focusGrid();
  };
  const isPie = chart.type === "pie" || chart.type === "doughnut";
  const isLine = chart.type === "line" || chart.type === "lineMarkers" || chart.type === "scatter";
  const palette = paletteOf(ctl.wb.theme);
  const elements = chartElements(ctl, chart);
  const el = elements.find((e) => e.id === ctl.chartElement) ?? elements[0];
  const va = chart.valAxis ?? {};
  const setVa = (patch: Partial<NonNullable<ChartDef["valAxis"]>>) => set({ valAxis: { ...va, ...patch } }, "Ekseni Biçimlendir");
  const si = el.id.startsWith("series:") ? Number(el.id.slice(7)) : -1;
  const setSeries = (patch: { color?: string; border?: string; lineWidth?: number }) => {
    const list = [...(chart.seriesFmt ?? [])];
    while (list.length <= si) list.push(null);
    const next = { ...(list[si] ?? {}), ...patch };
    for (const k of Object.keys(next) as (keyof typeof next)[]) if (next[k] === undefined) delete next[k];
    list[si] = Object.keys(next).length ? next : null;
    set({ seriesFmt: list.some(Boolean) ? list : undefined }, "Veri Serisini Biçimlendir");
  };
  return (
    <aside className="tb-pane tb-fmt-pane" aria-label={el.title}>
      <div className="tb-pane-head">
        <span>{el.title}</span>
        <button type="button" className="tb-pane-close" aria-label="Kapat" onClick={close}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="tb-fmt-elsel">
        <select
          aria-label="Grafik öğesi"
          value={el.id}
          onChange={(e) => {
            ctl.chartElement = e.target.value;
            ctl.emit();
          }}
        >
          {elements.map((x) => (
            <option key={x.id} value={x.id}>
              {x.label}
            </option>
          ))}
        </select>
      </div>
      <div className="tb-fmt-scroll">
        {(el.id === "title" || el.id === "area") && (
          <Section title="Grafik Başlığı" open={el.id === "title"}>
            <label className="tb-fmt-check">
              <input type="checkbox" checked={chart.showTitle} onChange={(e) => set({ showTitle: e.target.checked }, "Grafik Başlığı")} />
              Başlığı göster
            </label>
            <TextField label="Metin" value={chart.title} onCommit={(v) => set({ title: v, showTitle: true }, "Grafik Başlığı")} />
          </Section>
        )}
        {(el.id === "legend" || el.id === "area") && (
          <Section title="Gösterge Seçenekleri" open={el.id === "legend"}>
            <label className="tb-fmt-check">
              <input type="checkbox" checked={chart.showLegend} onChange={(e) => set({ showLegend: e.target.checked }, "Gösterge")} />
              Göstergeyi göster
            </label>
            <div className="tb-fmt-radios" role="radiogroup" aria-label="Gösterge konumu">
              {LEGEND_POS.map((p) => (
                <label key={p.v}>
                  <input type="radio" name={`lg-${chart.id}`} checked={chart.legendPos === p.v} disabled={!chart.showLegend} onChange={() => set({ legendPos: p.v }, "Gösterge")} />
                  {p.label}
                </label>
              ))}
            </div>
          </Section>
        )}
        {(el.id === "plot" || el.id === "area") && (
          <Section title="Öğeler" open={el.id === "plot"}>
            <label className="tb-fmt-check">
              <input type="checkbox" checked={chart.showDataLabels} onChange={(e) => set({ showDataLabels: e.target.checked }, "Veri Etiketleri")} />
              Veri etiketleri
            </label>
            {!isPie && (
              <>
                <label className="tb-fmt-check">
                  <input type="checkbox" checked={chart.showGridlines} onChange={(e) => set({ showGridlines: e.target.checked }, "Kılavuz Çizgileri")} />
                  Ana kılavuz çizgileri
                </label>
                <label className="tb-fmt-check">
                  <input type="checkbox" checked={!!va.minorGrid} onChange={(e) => setVa({ minorGrid: e.target.checked || undefined })} />
                  İkincil kılavuz çizgileri
                </label>
              </>
            )}
            <label className="tb-fmt-check">
              <input type="checkbox" checked={!!chart.showHidden} onChange={(e) => set({ showHidden: e.target.checked || undefined }, "Gizli ve Boş Hücreler")} />
              Gizli satır ve sütunlardaki verileri göster
            </label>
          </Section>
        )}
        {el.id === "axis:val" && (
          <>
            <Section title="Eksen Seçenekleri">
              <OptNum label="En Küçük" value={va.min} onCommit={(v) => setVa({ min: v })} />
              <OptNum label="En Büyük" value={va.max} onCommit={(v) => setVa({ max: v })} />
              <OptNum label="Ana birim" value={va.major} onCommit={(v) => setVa({ major: v !== undefined && v > 0 ? v : undefined })} />
              <label className="tb-fmt-check">
                <input type="checkbox" checked={chart.showGridlines} onChange={(e) => set({ showGridlines: e.target.checked }, "Kılavuz Çizgileri")} />
                Ana kılavuz çizgileri
              </label>
              <label className="tb-fmt-check">
                <input type="checkbox" checked={!!va.minorGrid} onChange={(e) => setVa({ minorGrid: e.target.checked || undefined })} />
                İkincil kılavuz çizgileri
              </label>
            </Section>
            <Section title="Sayı">
              <select aria-label="Eksen sayı biçimi" value={va.numFmt ?? ""} onChange={(e) => setVa({ numFmt: e.target.value || undefined })}>
                {AXIS_FORMATS.map((f) => (
                  <option key={f.code} value={f.code}>
                    {f.label}
                  </option>
                ))}
              </select>
            </Section>
            <Section title="Eksen Başlığı">
              <TextField label="Dikey" value={chart.yTitle ?? ""} onCommit={(v) => set({ yTitle: v || undefined }, "Eksen Başlığı")} />
            </Section>
          </>
        )}
        {el.id === "axis:cat" && (
          <Section title="Eksen Başlığı">
            <TextField label="Yatay" value={chart.xTitle ?? ""} onCommit={(v) => set({ xTitle: v || undefined }, "Eksen Başlığı")} />
          </Section>
        )}
        {si >= 0 && (
          <>
            <Section title={isLine ? "Çizgi" : "Dolgu"}>
              <ColorPalette theme={palette} value={chart.seriesFmt?.[si]?.color ?? seriesColor(chart, si)} noneLabel="Otomatik" onSelect={(c) => setSeries({ color: c ?? undefined })} />
            </Section>
            {isLine ? (
              <Section title="Çizgi Kalınlığı">
                <OptNum label="Kalınlık (pt)" value={chart.seriesFmt?.[si]?.lineWidth} onCommit={(v) => setSeries({ lineWidth: v !== undefined && v > 0 && v <= 20 ? v : undefined })} />
              </Section>
            ) : (
              <Section title="Kenarlık">
                <ColorPalette theme={palette} value={chart.seriesFmt?.[si]?.border} noneLabel="Kenarlık Yok" onSelect={(c) => setSeries({ border: c ?? undefined })} />
              </Section>
            )}
          </>
        )}
        {el.id === "labels" && (
          <Section title="Etiket Seçenekleri">
            <label className="tb-fmt-check">
              <input type="checkbox" checked={chart.showDataLabels} onChange={(e) => set({ showDataLabels: e.target.checked }, "Veri Etiketleri")} />
              Veri etiketlerini göster
            </label>
            {(chart.type.startsWith("column") || chart.type.startsWith("bar")) && (
              <div className="tb-fmt-radios" role="radiogroup" aria-label="Etiket konumu">
                {LABEL_POS.map((p) => (
                  <label key={p.v}>
                    <input type="radio" name={`lp-${chart.id}`} checked={(chart.labelPos ?? "outEnd") === p.v} disabled={!chart.showDataLabels} onChange={() => set({ labelPos: p.v }, "Veri Etiketleri")} />
                    {p.label}
                  </label>
                ))}
              </div>
            )}
          </Section>
        )}
        {el.id === "area" && (
          <>
            {!isPie && (
              <Section title="Eksen Başlıkları" open={false}>
                <TextField label="Yatay" value={chart.xTitle ?? ""} onCommit={(v) => set({ xTitle: v || undefined }, "Eksen Başlığı")} />
                <TextField label="Dikey" value={chart.yTitle ?? ""} onCommit={(v) => set({ yTitle: v || undefined }, "Eksen Başlığı")} />
              </Section>
            )}
            <Section title="Dolgu">
              <ColorPalette theme={palette} value={chart.fill} noneLabel="Otomatik" onSelect={(c) => set({ fill: c ?? undefined }, "Dolgu")} />
            </Section>
            <Section title="Kenarlık" open={false}>
              <label className="tb-fmt-check">
                <input type="checkbox" checked={chart.border !== "none"} onChange={(e) => set({ border: e.target.checked ? "#D9D9D9" : "none" }, "Kenarlık")} />
                Kenarlık
              </label>
              <ColorPalette theme={palette} value={chart.border === "none" ? undefined : chart.border} noneLabel="Kenarlık Yok" onSelect={(c) => set({ border: c ?? "none" }, "Kenarlık")} />
            </Section>
            <Section title="Renkler" open={false}>
              <select value={chart.palette} onChange={(e) => set({ palette: Number(e.target.value), seriesFmt: undefined }, "Renkleri Değiştir")} aria-label="Renk paleti">
                {PALETTE_LABELS.map((l, i) => (
                  <option key={l} value={i}>
                    {l}
                  </option>
                ))}
              </select>
            </Section>
            <Section title="Boyut">
              <NumField label="Yükseklik" value={chart.height / PX_PER_CM} onCommit={(cm) => set({ height: Math.round(cm * PX_PER_CM) }, "Boyut")} />
              <NumField label="Genişlik" value={chart.width / PX_PER_CM} onCommit={(cm) => set({ width: Math.round(cm * PX_PER_CM) }, "Boyut")} />
            </Section>
          </>
        )}
      </div>
    </aside>
  );
}
