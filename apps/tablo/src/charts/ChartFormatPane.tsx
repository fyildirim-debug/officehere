/**
 * Grafik Alanını Biçimlendir bölmesi (grafiğe çift tıklayınca açılır):
 * başlık, gösterge, veri etiketleri, kılavuz çizgileri, eksen başlıkları,
 * dolgu ve kenarlık, boyut. Değişiklikler anında uygulanır (geri alınabilir).
 */
import { useEffect, useState } from "react";
import { Dismiss16Regular } from "@fluentui/react-icons";
import { ColorPalette } from "@officehere/ui";
import type { Controller } from "../controller";
import type { ChartDef } from "../model/types";
import { updateChart } from "../actions/objects";
import { paletteOf } from "../model/theme";
import { PALETTE_LABELS } from "./chartData";

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

export function ChartFormatPane({ ctl, chart }: { ctl: Controller; chart: ChartDef }) {
  const set = (patch: Partial<ChartDef>, label = "Grafik Biçimi") => updateChart(ctl, chart.id, patch, label);
  const close = () => {
    ctl.formatPane = null;
    ctl.emit();
    ctl.focusGrid();
  };
  const isPie = chart.type === "pie" || chart.type === "doughnut";
  const palette = paletteOf(ctl.wb.theme);
  return (
    <aside className="tb-pane tb-fmt-pane" aria-label="Grafik Alanını Biçimlendir">
      <div className="tb-pane-head">
        <span>Grafik Alanını Biçimlendir</span>
        <button type="button" className="tb-pane-close" aria-label="Kapat" onClick={close}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="tb-fmt-scroll">
        <Section title="Grafik Başlığı">
          <label className="tb-fmt-check">
            <input type="checkbox" checked={chart.showTitle} onChange={(e) => set({ showTitle: e.target.checked }, "Grafik Başlığı")} />
            Başlığı göster
          </label>
          <TextField label="Metin" value={chart.title} onCommit={(v) => set({ title: v, showTitle: true }, "Grafik Başlığı")} />
        </Section>
        <Section title="Gösterge">
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
        <Section title="Öğeler">
          <label className="tb-fmt-check">
            <input type="checkbox" checked={chart.showDataLabels} onChange={(e) => set({ showDataLabels: e.target.checked }, "Veri Etiketleri")} />
            Veri etiketleri
          </label>
          {!isPie && (
            <label className="tb-fmt-check">
              <input type="checkbox" checked={chart.showGridlines} onChange={(e) => set({ showGridlines: e.target.checked }, "Kılavuz Çizgileri")} />
              Ana kılavuz çizgileri
            </label>
          )}
        </Section>
        {!isPie && (
          <Section title="Eksen Başlıkları">
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
          <select value={chart.palette} onChange={(e) => set({ palette: Number(e.target.value) }, "Renkleri Değiştir")} aria-label="Renk paleti">
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
      </div>
    </aside>
  );
}
