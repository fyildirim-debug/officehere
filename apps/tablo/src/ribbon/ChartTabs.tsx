/** Bağlamsal grafik sekmeleri: Grafik Tasarımı ve Biçim. */
import { useEffect, useState } from "react";
import { ArrowSwap24Regular, ArrowUp20Regular, ArrowDown20Regular, Color24Regular, DataBarVerticalAdd24Regular, DataBarVerticalEdit24Regular, ArrowMove24Regular, TableEdit24Regular, Delete20Regular } from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, RibbonButton, RibbonGallery, RibbonGroup, RibbonLabel, RibbonMenuButton, RibbonSpinner, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import { reorderObject, updateChart } from "../actions/objects";
import { deleteObject } from "../charts/FloatingLayer";
import { PALETTES, PALETTE_LABELS } from "../charts/chartData";
import { CHART_STYLES } from "../charts/ChartView";
import { ChartTypeIcon } from "../dialogs/ChartDialogs";
import type { ChartDef } from "../model/types";

function StylePreview({ index, palette }: { index: number; palette: number }) {
  const st = CHART_STYLES[index];
  const c = PALETTES[palette % PALETTES.length];
  return (
    <svg width="56" height="40" viewBox="0 0 56 40" aria-hidden="true">
      <rect x="0.5" y="0.5" width="55" height="39" fill={st.bg} stroke={st.border ?? "#d0d0d0"} />
      {[10, 20, 30].map((y) => (
        <line key={y} x1="6" x2="50" y1={y} y2={y} stroke={st.grid} strokeWidth="1" />
      ))}
      <rect x="9" y="18" width="6" height="16" fill={c[0]} />
      <rect x="16" y="12" width="6" height="22" fill={c[1]} />
      <rect x="29" y="14" width="6" height="20" fill={c[0]} />
      <rect x="36" y="7" width="6" height="27" fill={c[1]} />
      {index === 2 && <text x="28" y="6" fontSize="5" textAnchor="middle" fill={st.ink}>12 18</text>}
    </svg>
  );
}

const LEGEND: { id: ChartDef["legendPos"] | "none"; label: string }[] = [
  { id: "none", label: "Yok" },
  { id: "right", label: "Sağ" },
  { id: "top", label: "Üst" },
  { id: "left", label: "Sol" },
  { id: "bottom", label: "Alt" },
];

export function ChartDesignTab({ ctl, chart }: { ctl: Controller; chart: ChartDef }) {
  const up = (patch: Partial<ChartDef>, label = "Grafik Öğesi") => updateChart(ctl, chart.id, patch, label);
  const isPie = chart.type === "pie" || chart.type === "doughnut";
  return (
    <>
      <RibbonGroup label="Grafik Düzenleri" icon={<DataBarVerticalAdd24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DataBarVerticalAdd24Regular />}
          label="Grafik Öğesi Ekle" className="tb-wide"
          menu={
            <Menu width={220}>
              <MenuItem label="Grafik Başlığı" checked={chart.showTitle} onClick={() => up({ showTitle: !chart.showTitle })} />
              <MenuItem label="Veri Etiketleri" checked={chart.showDataLabels} onClick={() => up({ showDataLabels: !chart.showDataLabels })} />
              {!isPie && <MenuItem label="Kılavuz Çizgileri" checked={chart.showGridlines} onClick={() => up({ showGridlines: !chart.showGridlines })} />}
              {!isPie && <MenuItem label="Eksen Başlıkları" checked={!!(chart.xTitle || chart.yTitle)} onClick={() => up(chart.xTitle || chart.yTitle ? { xTitle: undefined, yTitle: undefined } : { xTitle: "Eksen Başlığı", yTitle: "Eksen Başlığı" })} />}
              <MenuItem
                label="Gösterge"
                submenu={
                  <Menu width={140}>
                    {LEGEND.map((l) => (
                      <MenuItem key={l.id} label={l.label} checked={l.id === "none" ? !chart.showLegend : chart.showLegend && chart.legendPos === l.id} onClick={() => up(l.id === "none" ? { showLegend: false } : { showLegend: true, legendPos: l.id })} />
                    ))}
                  </Menu>
                }
              />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Grafik Stilleri" icon={<Color24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Color24Regular />}
          label="Renkleri Değiştir"
          menu={
            <Menu width={260}>
              <MenuHeader>Renk Paletleri</MenuHeader>
              {PALETTES.map((p, i) => (
                <MenuItem
                  key={i}
                  checked={chart.palette === i}
                  icon={
                    <svg width="48" height="12" aria-hidden="true">
                      {p.slice(0, 6).map((c, k) => (
                        <rect key={k} x={k * 8} y="0" width="7" height="12" fill={c} />
                      ))}
                    </svg>
                  }
                  label={PALETTE_LABELS[i]}
                  onClick={() => up({ palette: i }, "Renkleri Değiştir")}
                />
              ))}
            </Menu>
          }
        />
        <RibbonGallery
          items={CHART_STYLES.slice(1).map((_, i) => ({ id: String(i + 1), label: `Stil ${i + 1}`, preview: <StylePreview index={i + 1} palette={chart.palette} /> }))}
          selectedId={String(chart.style)}
          onSelect={(id) => up({ style: Number(id), showDataLabels: id === "2" ? true : chart.showDataLabels }, "Grafik Stili")}
          itemWidth={60}
          itemHeight={44}
          columns={6}
        />
      </RibbonGroup>
      <RibbonGroup label="Veri" icon={<TableEdit24Regular />}>
        <RibbonButton size="large" icon={<ArrowSwap24Regular />} label="Satır/Sütun Değiştir" disabled={chart.type === "scatter"} onClick={() => up({ seriesInRows: !chart.seriesInRows }, "Satır/Sütun Değiştir")} />
        <RibbonButton size="large" icon={<TableEdit24Regular />} label="Veri Seç" onClick={() => ctl.openDialog("chartData")} />
      </RibbonGroup>
      <RibbonGroup label="Tür" icon={<DataBarVerticalEdit24Regular />}>
        <RibbonButton size="large" icon={<ChartTypeIcon type={chart.type} size={32} />} label="Grafik Türünü Değiştir" className="tb-wide" onClick={() => ctl.openDialog("chartType")} />
      </RibbonGroup>
      <RibbonGroup label="Konum" icon={<ArrowMove24Regular />}>
        <RibbonButton size="large" icon={<ArrowMove24Regular />} label="Grafiği Taşı" description="Grafiği bu çalışma kitabındaki başka bir sayfaya taşır." onClick={() => ctl.openDialog("moveChart")} />
      </RibbonGroup>
    </>
  );
}

const CM = 96 / 2.54;

export function ChartFormatTab({ ctl, chart }: { ctl: Controller; chart: ChartDef }) {
  const up = (patch: Partial<ChartDef>, label = "Grafik Biçimi") => updateChart(ctl, chart.id, patch, label);
  const [title, setTitle] = useState(chart.title);
  const [xt, setXt] = useState(chart.xTitle ?? "");
  const [yt, setYt] = useState(chart.yTitle ?? "");
  useEffect(() => {
    setTitle(chart.title);
    setXt(chart.xTitle ?? "");
    setYt(chart.yTitle ?? "");
  }, [chart.id, chart.title, chart.xTitle, chart.yTitle]);
  const isPie = chart.type === "pie" || chart.type === "doughnut";
  const field = (label: string, value: string, set: (v: string) => void, commit: (v: string) => void, disabled = false) => (
    <div className="tb-ribbon-field">
      <RibbonLabel width={86}>{label}</RibbonLabel>
      <input
        className="tb-ribbon-input"
        value={value}
        disabled={disabled}
        onChange={(e) => set(e.target.value)}
        onBlur={() => commit(value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        }}
      />
    </div>
  );
  return (
    <>
      <RibbonGroup label="Başlıklar">
        <RibbonStack>
          {field("Grafik başlığı:", title, setTitle, (v) => v !== chart.title && up({ title: v, showTitle: true }, "Grafik Başlığı"))}
          {field("Yatay eksen:", xt, setXt, (v) => v !== (chart.xTitle ?? "") && up({ xTitle: v || undefined }, "Eksen Başlığı"), isPie)}
          {field("Dikey eksen:", yt, setYt, (v) => v !== (chart.yTitle ?? "") && up({ yTitle: v || undefined }, "Eksen Başlığı"), isPie)}
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Yerleştir">
        <RibbonStack>
          <RibbonButton icon={<ArrowUp20Regular />} label="Öne Getir" onClick={() => reorderObject(ctl, "chart", chart.id, true)} />
          <RibbonButton icon={<ArrowDown20Regular />} label="Arkaya Gönder" onClick={() => reorderObject(ctl, "chart", chart.id, false)} />
          <RibbonButton icon={<Delete20Regular />} label="Grafiği Sil" onClick={() => deleteObject(ctl, "chart", chart.id)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Boyut">
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={chart.height / CM} unit="cm" step={0.1} min={1} max={50} decimals={2} width={80} onChange={(v) => up({ height: Math.round(v * CM) }, "Grafik Boyutu")} />
          <RibbonSpinner label="Genişlik:" value={chart.width / CM} unit="cm" step={0.1} min={1} max={80} decimals={2} width={80} onChange={(v) => up({ width: Math.round(v * CM) }, "Grafik Boyutu")} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
