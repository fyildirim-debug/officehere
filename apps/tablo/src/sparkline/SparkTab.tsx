/** Bağlamsal "Mini Grafik" sekmesi: Mini Grafik, Tür, Göster, Stil, Grupla. */
import { DataUsageEdit24Regular, Eraser20Regular, Group20Regular, GroupDismiss20Regular, Color20Regular } from "@fluentui/react-icons";
import { themeShades } from "@officehere/core";
import { ColorPalette, Menu, MenuHeader, MenuItem, MenuSeparator, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonButton, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import type { SparklineGroup } from "../model/types";
import { SparkGlyph } from "../ribbon/glyphs";
import { useThemePalette } from "../themeContext";
import { POINT_RED } from "./geometry";
import { clearSparklines, groupSparklines, ungroupSparklines, updateGroup } from "./sparkOps";

const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

interface StylePreset {
  id: string;
  label: string;
  color: string;
  points: string;
}

/** Stil galerisi: tema vurgularından türetilen 12 stil. */
function stylePresets(accents: readonly string[]): StylePreset[] {
  const out: StylePreset[] = [];
  accents.forEach((a, i) => out.push({ id: `d${i}`, label: `Mini Grafik Stili Vurgu ${i + 1}, %25 Daha Koyu`, color: themeShades(a)[4].toUpperCase(), points: POINT_RED }));
  accents.forEach((a, i) => out.push({ id: `l${i}`, label: `Mini Grafik Stili Vurgu ${i + 1}`, color: a.toUpperCase(), points: themeShades(a)[5].toUpperCase() }));
  return out;
}

function StylePreview({ p }: { p: StylePreset }) {
  return (
    <svg width="46" height="30" viewBox="0 0 46 30" aria-hidden="true">
      <rect x="0.5" y="0.5" width="45" height="29" fill="#fff" stroke="#d0d0d0" />
      <polyline points="5,20 13,12 21,17 29,8 37,14 41,10" fill="none" stroke={p.color} strokeWidth="1.6" strokeLinejoin="round" />
      <circle cx="29" cy="8" r="2.2" fill={p.points} />
      <circle cx="5" cy="20" r="2.2" fill={p.points} />
    </svg>
  );
}

/** İşaretçi Rengi simgesi: noktalı çizgi. */
function MarkerGlyph() {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <path d="M2.5 14l5-6 4 3 6-6" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinejoin="round" />
      <circle cx="7.5" cy="8" r="2" fill="#D00000" />
      <circle cx="17.5" cy="5" r="2" fill="#D00000" />
    </svg>
  );
}

/** Eksen düğmesi simgesi. */
function AxisGlyph() {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <path d="M3 2.5v15h14.5" fill="none" stroke="currentColor" strokeWidth="1.2" />
      <path d="M5 13l3-4 3 2 4-6" fill="none" stroke="#107C41" strokeWidth="1.4" strokeLinejoin="round" />
    </svg>
  );
}

export function SparklineTab({ ctl, group }: { ctl: Controller; group: SparklineGroup }) {
  const palette = useThemePalette();
  const set = (patch: Partial<SparklineGroup>, label: string) => updateGroup(ctl, group.id, patch, label);
  const presets = stylePresets(ctl.wb.theme.colors.slice(4));
  const current = presets.find((p) => p.color === group.color.toUpperCase())?.id;
  const pointMenu = (label: string, key: "negativeColor" | "markerColor" | "highColor" | "lowColor" | "firstColor" | "lastColor", show: keyof SparklineGroup) => (
    <MenuItem
      label={label}
      submenu={
        <ColorPalette
          theme={palette}
          value={(group[key] as string | undefined) ?? (key === "markerColor" ? group.color : POINT_RED)}
          onSelect={(c) => c && set({ [key]: c, [show]: true } as Partial<SparklineGroup>, "İşaretçi Rengi")}
        />
      }
    />
  );
  const type = (t: SparklineGroup["type"], label: string) => (
    <RibbonButton size="large" icon={<SparkGlyph kind={t} />} label={label} checked={group.type === t} onClick={() => group.type !== t && set({ type: t, negative: t === "winLoss" ? true : group.negative }, "Mini Grafik Türü")} />
  );
  return (
    <>
      <RibbonGroup label="Mini Grafik" icon={<DataUsageEdit24Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<DataUsageEdit24Regular />}
          label="Veri Düzenle"
          className="tb-split-nowrap"
          description="Mini grafik grubunun konumunu ve veri kaynağını düzenler."
          onClick={() => ctl.openDialog("sparkline", { edit: group.id })}
          menu={
            <Menu width={290}>
              <MenuItem label="Grup Konumunu ve Verilerini Düzenle..." onClick={() => ctl.openDialog("sparkline", { edit: group.id })} />
              <MenuItem label="Tek Mini Grafiğin Verilerini Düzenle..." onClick={() => ctl.openDialog("sparklineItem")} />
              <MenuSeparator />
              <MenuItem label="Gizli ve Boş Hücreler..." onClick={() => ctl.openDialog("sparklineHidden", { group: group.id })} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Tür" icon={<SparkGlyph kind="line" />}>
        {type("line", "Çizgi")}
        {type("column", "Sütun")}
        {type("winLoss", "Kazanç/Kayıp")}
      </RibbonGroup>
      <RibbonGroup label="Göster" icon={<MarkerGlyph />}>
        <RibbonStack>
          <RibbonCheckbox label="Yüksek Nokta" checked={!!group.highPoint} onChange={(v) => set({ highPoint: v || undefined }, "Yüksek Nokta")} />
          <RibbonCheckbox label="Düşük Nokta" checked={!!group.lowPoint} onChange={(v) => set({ lowPoint: v || undefined }, "Düşük Nokta")} />
          <RibbonCheckbox label="Negatif Noktalar" checked={!!group.negative} onChange={(v) => set({ negative: v || undefined }, "Negatif Noktalar")} />
        </RibbonStack>
        <RibbonStack>
          <RibbonCheckbox label="İlk Nokta" checked={!!group.firstPoint} onChange={(v) => set({ firstPoint: v || undefined }, "İlk Nokta")} />
          <RibbonCheckbox label="Son Nokta" checked={!!group.lastPoint} onChange={(v) => set({ lastPoint: v || undefined }, "Son Nokta")} />
          <RibbonCheckbox label="İşaretçiler" checked={!!group.markers} disabled={group.type !== "line"} onChange={(v) => set({ markers: v || undefined }, "İşaretçiler")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Stil" icon={<Color20Regular />}>
        <RibbonGallery
          items={presets.map((p) => ({ id: p.id, label: p.label, preview: <StylePreview p={p} /> }))}
          selectedId={current}
          onSelect={(id) => {
            const p = presets.find((x) => x.id === id)!;
            set({ color: p.color, markerColor: p.color, negativeColor: p.points, highColor: p.points, lowColor: p.points, firstColor: p.points, lastColor: p.points }, "Mini Grafik Stili");
          }}
          itemWidth={50}
          itemHeight={34}
          columns={4}
          expandedColumns={6}
        />
        <RibbonStack>
          <RibbonMenuButton
            icon={<Color20Regular />}
            label="Mini Grafik Rengi"
            menu={
              <Menu width={250}>
                <ColorPalette theme={palette} value={group.color} onSelect={(c) => c && set({ color: c }, "Mini Grafik Rengi")} />
                <MenuSeparator />
                <MenuItem
                  label="Kalınlık"
                  submenu={
                    <Menu width={140}>
                      {WEIGHTS.map((w) => (
                        <MenuItem key={w} label={`${String(w).replace(".", ",")} nk`} checked={(group.weight ?? 0.75) === w} onClick={() => set({ weight: w }, "Kalınlık")} />
                      ))}
                    </Menu>
                  }
                />
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<MarkerGlyph />}
            label="İşaretçi Rengi"
            menu={
              <Menu width={220}>
                {pointMenu("Negatif Noktalar", "negativeColor", "negative")}
                {pointMenu("İşaretçiler", "markerColor", "markers")}
                {pointMenu("Yüksek Nokta", "highColor", "highPoint")}
                {pointMenu("Düşük Nokta", "lowColor", "lowPoint")}
                {pointMenu("İlk Nokta", "firstColor", "firstPoint")}
                {pointMenu("Son Nokta", "lastColor", "lastPoint")}
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Grupla" icon={<Group20Regular />}>
        <RibbonStack>
          <RibbonMenuButton
            icon={<AxisGlyph />}
            label="Eksen"
            menu={
              <Menu width={300}>
                <MenuHeader>Yatay Eksen Seçenekleri</MenuHeader>
                <MenuItem label="Ekseni Göster" checked={!!group.showAxis} onClick={() => set({ showAxis: !group.showAxis || undefined }, "Ekseni Göster")} />
                <MenuItem label="Sağdan Sola Çizim Verileri" checked={!!group.rightToLeft} onClick={() => set({ rightToLeft: !group.rightToLeft || undefined }, "Sağdan Sola Çizim Verileri")} />
                <MenuHeader>Dikey Eksen En Küçük Değer Seçenekleri</MenuHeader>
                <MenuItem label="Her Mini Grafik İçin Otomatik" checked={!group.minAxis || group.minAxis === "individual"} onClick={() => set({ minAxis: undefined }, "Dikey Eksen")} />
                <MenuItem label="Tüm Mini Grafikler İçin Aynı" checked={group.minAxis === "group"} onClick={() => set({ minAxis: "group" }, "Dikey Eksen")} />
                <MenuItem label="Özel Değer..." checked={group.minAxis === "custom"} onClick={() => ctl.openDialog("sparklineAxis", { group: group.id, which: "min" })} />
                <MenuHeader>Dikey Eksen En Büyük Değer Seçenekleri</MenuHeader>
                <MenuItem label="Her Mini Grafik İçin Otomatik" checked={!group.maxAxis || group.maxAxis === "individual"} onClick={() => set({ maxAxis: undefined }, "Dikey Eksen")} />
                <MenuItem label="Tüm Mini Grafikler İçin Aynı" checked={group.maxAxis === "group"} onClick={() => set({ maxAxis: "group" }, "Dikey Eksen")} />
                <MenuItem label="Özel Değer..." checked={group.maxAxis === "custom"} onClick={() => ctl.openDialog("sparklineAxis", { group: group.id, which: "max" })} />
              </Menu>
            }
          />
          <RibbonButton icon={<Group20Regular />} label="Gruplandır" description="Seçili mini grafikleri tek grupta birleştirir; grup ayarları etkin hücredekinden alınır." onClick={() => groupSparklines(ctl)} />
          <RibbonButton icon={<GroupDismiss20Regular />} label="Grubu Çöz" description="Seçili mini grafik grubunu ayrı mini grafiklere böler." disabled={group.items.length < 2} onClick={() => ungroupSparklines(ctl)} />
        </RibbonStack>
        <RibbonSplitButton
          size="large"
          icon={<Eraser20Regular />}
          label="Temizle"
          description="Seçili mini grafikleri kaldırır."
          onClick={() => clearSparklines(ctl, false)}
          menu={
            <Menu width={290}>
              <MenuItem label="Seçili Mini Grafikleri Temizle" onClick={() => clearSparklines(ctl, false)} />
              <MenuItem label="Seçili Mini Grafik Gruplarını Temizle" onClick={() => clearSparklines(ctl, true)} />
            </Menu>
          }
        />
      </RibbonGroup>
    </>
  );
}
