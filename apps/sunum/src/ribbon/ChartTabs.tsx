/** Bağlamsal sekmeler: Grafik Tasarımı (grafik seçiliyken) ve Kayıttan Yürüt (video/ses seçiliyken). */
import { ChartMultiple24Regular, DataBarVerticalEdit24Regular, Play24Regular, TableEdit24Regular, Video24Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { selectedChart, setMediaOption, updateChart } from "../state/mediaActions";
import type { RibbonCtx } from "./types";

export function ChartDesignTab({ ctx }: { ctx: RibbonCtx }) {
  const { actions: a, cmd } = ctx;
  const ch = selectedChart(a);
  const c = ch?.chart;
  const set = (patch: Partial<NonNullable<typeof c>>) => ch && updateChart(a, ch.id, (d) => ({ ...d, ...patch }));
  const pie = c?.type === "pie" || c?.type === "doughnut";
  return (
    <>
      <RibbonGroup label="Grafik Düzenleri" icon={<ChartMultiple24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DataBarVerticalEdit24Regular />}
          label="Grafik Öğesi Ekle"
          description="Grafik başlığı, gösterge ve veri etiketleri gibi öğeleri ekleyin ya da kaldırın."
          disabled={!c}
          menu={
            <Menu>
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
              <MenuItem
                label="Veri Etiketleri"
                submenu={
                  <Menu>
                    <MenuItem label="Yok" checked={!c?.dataLabels} onClick={() => set({ dataLabels: undefined })} />
                    <MenuItem label={pie ? "Ortala" : "Göster"} checked={!!c?.dataLabels} onClick={() => set({ dataLabels: true })} />
                  </Menu>
                }
              />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Veri" icon={<TableEdit24Regular />}>
        <RibbonButton size="large" icon={<TableEdit24Regular />} label="Verileri Düzenle" description="Grafiğin kategorilerini, serilerini ve değerlerini düzenleyin." disabled={!c} onClick={cmd.editChartData} />
      </RibbonGroup>
      <RibbonGroup label="Tür" icon={<ChartMultiple24Regular />}>
        <RibbonButton size="large" className="rb-wide" icon={<ChartMultiple24Regular />} label="Grafik Türünü Değiştir" description="Farklı bir grafik türüne geçin." disabled={!c} onClick={cmd.changeChartType} />
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
