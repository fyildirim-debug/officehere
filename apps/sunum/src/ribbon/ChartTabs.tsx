/**
 * Bağlamsal sekmeler: Grafik Tasarımı ve Biçim (grafik seçiliyken),
 * Kayıttan Yürüt (video/ses seçiliyken).
 */
import { useMemo } from "react";
import {
  ArrowSwap24Regular,
  ChartMultiple24Regular,
  Color24Regular,
  DataBarVerticalEdit24Regular,
  LayoutCellFour24Regular,
  PaintBucket20Regular,
  Play24Regular,
  BorderAll20Regular,
  TableEdit24Regular,
  Video24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { CHART_PALETTES, CHART_STYLES, effectiveLabelPos, hasAxes, isPieType, labelPositions, QUICK_LAYOUTS, seriesColor, switchRowColumn } from "../model/chart";
import { cssColor } from "../model/color";
import type { ChartData, Color } from "../model/types";
import { ChartSvg } from "../render/ChartSvg";
import { selectedChart, setMediaOption, updateChart } from "../state/mediaActions";
import { ThemeColorSplit } from "./common";
import type { RibbonCtx } from "./types";

function useChart(ctx: RibbonCtx) {
  const ch = selectedChart(ctx.actions);
  const c = ch?.chart;
  const set = (patch: Partial<ChartData>) => ch && updateChart(ctx.actions, ch.id, (d) => ({ ...d, ...patch }));
  return { ch, c, set };
}

function PaletteSwatches({ ctx, index }: { ctx: RibbonCtx; index: number }) {
  const theme = ctx.state.pres.theme;
  return (
    <span style={{ display: "inline-flex", gap: 1 }}>
      {[0, 1, 2, 3, 4, 5].map((i) => (
        <span key={i} style={{ width: 11, height: 12, background: cssColor(seriesColor(i, undefined, index), theme), border: "1px solid rgba(0,0,0,.12)" }} />
      ))}
    </span>
  );
}

export function ChartDesignTab({ ctx }: { ctx: RibbonCtx }) {
  const { cmd, state } = ctx;
  const { c, set } = useChart(ctx);
  const pie = !!c && isPieType(c.type);
  const axes = !!c && hasAxes(c.type);
  const theme = state.pres.theme;
  // Galeri önizlemeleri: geçerli grafiğin verisiyle küçük çizimler.
  const styleItems = useMemo(
    () =>
      CHART_STYLES.map((_, i) => ({
        id: String(i + 1),
        label: `Stil ${i + 1}`,
        preview: c ? (
          <span className="chart-thumb" style={{ width: 70, height: 46 }}>
            <ChartSvg chart={{ ...c, style: i + 1, showTitle: false, legend: "none", axisTitles: undefined }} w={70} h={46} theme={theme} />
          </span>
        ) : (
          <span />
        ),
      })),
    [c, theme],
  );
  const toggleAxisTitle = (k: "cat" | "val") => {
    if (!c) return;
    const cur = { ...(c.axisTitles ?? {}) };
    if (cur[k]) delete cur[k];
    else cur[k] = "Eksen Başlığı";
    set({ axisTitles: cur.cat || cur.val ? cur : undefined });
  };
  return (
    <>
      <RibbonGroup label="Grafik Düzenleri" icon={<ChartMultiple24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DataBarVerticalEdit24Regular />}
          label="Grafik Öğesi Ekle"
          description="Eksen başlıkları, grafik başlığı, gösterge, veri etiketleri ve kılavuz çizgileri ekleyin ya da kaldırın."
          disabled={!c}
          menu={
            <Menu>
              <MenuItem
                label="Eksen Başlıkları"
                disabled={!axes}
                submenu={
                  <Menu>
                    <MenuItem label="Birincil Yatay" checked={!!c?.axisTitles?.cat} onClick={() => toggleAxisTitle("cat")} />
                    <MenuItem label="Birincil Dikey" checked={!!c?.axisTitles?.val} onClick={() => toggleAxisTitle("val")} />
                  </Menu>
                }
              />
              <MenuItem
                label="Grafik Başlığı"
                submenu={
                  <Menu>
                    <MenuItem label="Yok" checked={!c?.showTitle} onClick={() => set({ showTitle: false })} />
                    <MenuItem label="Grafiğin Üstünde" checked={!!c?.showTitle} onClick={() => set({ showTitle: true })} />
                  </Menu>
                }
              />
              <MenuItem
                label="Veri Etiketleri"
                submenu={
                  <Menu>
                    <MenuItem label="Yok" checked={!c?.dataLabels} onClick={() => set({ dataLabels: undefined, labelPos: undefined })} />
                    {c && labelPositions(c.type, !!c.grouping && c.grouping !== "clustered").length ? (
                      labelPositions(c.type, !!c.grouping && c.grouping !== "clustered").map((x) => (
                        <MenuItem key={x.pos} label={x.label} checked={!!c.dataLabels && effectiveLabelPos(c) === x.pos} onClick={() => set({ dataLabels: true, labelPos: x.pos })} />
                      ))
                    ) : (
                      <MenuItem label="Göster" checked={!!c?.dataLabels} onClick={() => set({ dataLabels: true })} />
                    )}
                  </Menu>
                }
              />
              <MenuItem
                label="Kılavuz Çizgileri"
                disabled={pie}
                submenu={
                  <Menu>
                    <MenuItem label="Birincil Ana Yatay" checked={c?.gridMajor !== false} onClick={() => set({ gridMajor: c?.gridMajor === false ? undefined : false })} />
                    <MenuItem label="Birincil İkincil Yatay" checked={!!c?.gridMinor} onClick={() => set({ gridMinor: c?.gridMinor ? undefined : true })} />
                  </Menu>
                }
              />
              <MenuItem
                label="Gösterge"
                submenu={
                  <Menu>
                    {(
                      [
                        ["none", "Yok"],
                        ["r", "Sağ"],
                        ["t", "Üst"],
                        ["l", "Sol"],
                        ["b", "Alt"],
                      ] as const
                    ).map(([v, l]) => (
                      <MenuItem key={v} label={l} checked={c?.legend === v} onClick={() => set({ legend: v })} />
                    ))}
                  </Menu>
                }
              />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<LayoutCellFour24Regular />}
          label="Hızlı Düzen"
          description="Grafiğin genel düzenini değiştirin."
          disabled={!c}
          menu={
            <Menu>
              <div className="preset-grid" style={{ gridTemplateColumns: "repeat(4, 76px)" }}>
                {QUICK_LAYOUTS.map((q) => (
                  <button key={q.name} type="button" className="preset-btn" title={q.name} aria-label={q.name} onMouseDown={(e) => e.preventDefault()} onClick={() => c && set(q.patch(c))} style={{ width: 76, height: 56 }}>
                    {c && (
                      <span className="chart-thumb" style={{ width: 72, height: 52 }}>
                        <ChartSvg chart={{ ...c, ...q.patch(c) }} w={72} h={52} theme={theme} />
                      </span>
                    )}
                  </button>
                ))}
              </div>
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Grafik Stilleri" icon={<Color24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Color24Regular />}
          label="Renkleri Değiştir"
          description="Grafiğin renk paletini seçin."
          disabled={!c}
          menu={
            <Menu className="menu-wide">
              <MenuHeader>Renkli</MenuHeader>
              {CHART_PALETTES.map((p, i) => (
                <div key={p.name}>
                  {i === 4 && <MenuHeader>Tek Renkli</MenuHeader>}
                  <MenuItem icon={<span />} checked={(c?.palette ?? 0) === i} label={<span style={{ display: "flex", gap: 8, alignItems: "center" }}><PaletteSwatches ctx={ctx} index={i} /> {p.name}</span>} onClick={() => set({ palette: i || undefined })} />
                </div>
              ))}
            </Menu>
          }
        />
        <RibbonGallery columns={4} itemWidth={74} itemHeight={50} selectedId={String(c?.style ?? 1)} items={styleItems} onSelect={(id) => c && set({ style: Number(id) === 1 ? undefined : Number(id) })} />
      </RibbonGroup>
      <RibbonGroup label="Veri" icon={<TableEdit24Regular />}>
        <RibbonButton size="large" icon={<ArrowSwap24Regular />} label="Satır/Sütun Değiştir" description="Eksendeki verileri değiştirin: seriler kategori, kategoriler seri olur." disabled={!c || pie} onClick={() => c && set(switchRowColumn(c))} />
        <RibbonButton size="large" icon={<TableEdit24Regular />} label="Verileri Düzenle" description="Grafiğin kategorilerini, serilerini ve değerlerini düzenleyin." disabled={!c} onClick={cmd.editChartData} />
      </RibbonGroup>
      <RibbonGroup label="Tür" icon={<ChartMultiple24Regular />}>
        <RibbonButton size="large" className="rb-wide" icon={<ChartMultiple24Regular />} label="Grafik Türünü Değiştir" description="Farklı bir grafik türüne geçin." disabled={!c} onClick={cmd.changeChartType} />
      </RibbonGroup>
    </>
  );
}

/** Grafik "Biçim" sekmesi: grafik alanı dolgusu ve anahattı. */
export function ChartFormatTab({ ctx }: { ctx: RibbonCtx }) {
  const { c, set } = useChart(ctx);
  const theme = ctx.state.pres.theme;
  return (
    <>
      <RibbonGroup label="Geçerli Seçim" icon={<ChartMultiple24Regular />}>
        <RibbonStack>
          <span className="rb-note">Grafik Alanı</span>
          <RibbonButton icon={<BorderAll20Regular />} label="Stile Uyması İçin Sıfırla" disabled={!c} onClick={() => set({ fill: undefined, line: undefined, style: undefined, palette: undefined })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Şekil Stilleri" icon={<PaintBucket20Regular />}>
        <RibbonStack>
          <ThemeColorSplit icon={<PaintBucket20Regular />} label="Şekil Dolgusu" tip="Şekil Dolgusu" size="medium" theme={theme} color={c?.fill ?? null} disabled={!c} noneLabel="Dolgu Yok" onApply={(col: Color | null) => set({ fill: col ?? undefined })} onColorChange={(col: Color | null) => set({ fill: col ?? undefined })} />
          <ThemeColorSplit icon={<BorderAll20Regular />} label="Şekil Anahattı" tip="Şekil Anahattı" size="medium" theme={theme} color={c?.line ?? null} disabled={!c} noneLabel="Anahat Yok" onApply={(col: Color | null) => set({ line: col ?? undefined })} onColorChange={(col: Color | null) => set({ line: col ?? undefined })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Seçenekler" icon={<ChartMultiple24Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Veri etiketleri" checked={!!c?.dataLabels} disabled={!c} onChange={(v) => set({ dataLabels: v || undefined })} />
          <RibbonCheckbox label="Ana kılavuz çizgileri" checked={c?.gridMajor !== false} disabled={!c || !hasAxes(c.type)} onChange={(v) => set({ gridMajor: v ? undefined : false })} />
          <RibbonCheckbox label="Gösterge" checked={!!c && c.legend !== "none"} disabled={!c} onChange={(v) => set({ legend: v ? "b" : "none" })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

export function PlaybackTab({ ctx }: { ctx: RibbonCtx }) {
  const { actions: a, cmd } = ctx;
  const s = a.selectedShapes();
  const m = s.length === 1 && s[0].kind === "pic" ? s[0].media : undefined;
  return (
    <>
      <RibbonGroup label="Önizleme" icon={<Play24Regular />}>
        <RibbonButton size="large" icon={<Play24Regular />} label="Yürüt" description={m?.kind === "audio" ? "Sesi önizleyin." : "Videoyu önizleyin."} disabled={!m} onClick={cmd.playMedia} />
      </RibbonGroup>
      <RibbonGroup label={m?.kind === "audio" ? "Ses Seçenekleri" : "Video Seçenekleri"} icon={<Video24Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Otomatik başlat" checked={!!m?.autoPlay} disabled={!m} onChange={(v) => setMediaOption(a, { autoPlay: v || undefined })} />
          <RibbonCheckbox label="Durdurulana Kadar Döngü" checked={!!m?.loop} disabled={!m} onChange={(v) => setMediaOption(a, { loop: v || undefined })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
