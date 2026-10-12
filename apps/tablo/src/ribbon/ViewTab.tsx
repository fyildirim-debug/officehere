/** Görünüm sekmesi: Çalışma Kitabı Görünümleri, Göster, Yakınlaştır, Pencere. */
import { DocumentPageBreak24Regular, Grid20Regular, PageFit24Regular, ZoomFit24Regular, ZoomIn24Regular, TableFreezeColumnAndRow20Regular, TableFreezeRow20Regular, TableFreezeColumn20Regular, TableFreezeColumnAndRowDismiss20Regular, Window24Regular, WindowNew24Regular, Eye24Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import { freeze, isFrozen, setViewMode, setZoom, zoomToSelection } from "../actions/view";
import { FreezeGlyph } from "./glyphs";
import { openNewWindow } from "../app/windowSync";

export function ViewTab({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const frozen = isFrozen(ctl);
  const toggle = (k: "showGridlines" | "showHeadings", v: boolean) => {
    sheet[k] = v;
    ctl.emit();
  };
  return (
    <>
      <RibbonGroup label="Çalışma Kitabı Görünümleri" icon={<Window24Regular />}>
        <RibbonButton size="large" icon={<Grid20Regular />} label="Normal" checked={sheet.viewMode === "normal"} onClick={() => setViewMode(ctl, "normal")} />
        <RibbonButton size="large" icon={<DocumentPageBreak24Regular />} label="Sayfa Sonu Önizleme" checked={sheet.viewMode === "pageBreak"} onClick={() => setViewMode(ctl, "pageBreak")} />
        <RibbonButton size="large" icon={<PageFit24Regular />} label="Sayfa Düzeni" checked={sheet.viewMode === "pageLayout"} onClick={() => setViewMode(ctl, "pageLayout")} />
      </RibbonGroup>
      <RibbonGroup label="Göster" icon={<Eye24Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Cetvel" checked={ctl.showRuler} disabled={sheet.viewMode !== "pageLayout"} onChange={(v) => ((ctl.showRuler = v), ctl.emit())} />
          <RibbonCheckbox label="Kılavuz Çizgileri" checked={sheet.showGridlines} onChange={(v) => toggle("showGridlines", v)} />
        </RibbonStack>
        <RibbonStack>
          <RibbonCheckbox label="Formül Çubuğu" checked={ctl.showFormulaBar} onChange={(v) => ((ctl.showFormulaBar = v), ctl.emit())} />
          <RibbonCheckbox label="Başlıklar" checked={sheet.showHeadings} onChange={(v) => toggle("showHeadings", v)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Yakınlaştır" icon={<ZoomIn24Regular />}>
        <RibbonButton size="large" icon={<ZoomIn24Regular />} label="Yakınlaştır" description="Yakınlaştırma düzeyini seçer." onClick={() => ctl.openDialog("zoom")} />
        <RibbonButton size="large" icon={<span className="tb-glyph-100">100</span>} label="%100" onClick={() => setZoom(ctl, 100)} />
        <RibbonButton size="large" icon={<ZoomFit24Regular />} label="Seçime Yakınlaştır" description="Seçili hücre aralığı pencereyi dolduracak şekilde yakınlaştırır." onClick={() => zoomToSelection(ctl)} />
      </RibbonGroup>
      <RibbonGroup label="Pencere" icon={<TableFreezeColumnAndRow20Regular />}>
        <RibbonButton size="large" icon={<WindowNew24Regular />} label="Yeni Pencere" description="Çalışma kitabınızı ikinci bir pencerede açar; iki penceredeki değişiklikler eşzamanlı yansır." onClick={() => openNewWindow(ctl)} />
        <RibbonMenuButton
          size="large"
          icon={<FreezeGlyph />}
          label="Bölmeleri Dondur"
          menu={
            <Menu width={300}>
              {frozen ? (
                <MenuItem icon={<TableFreezeColumnAndRowDismiss20Regular />} label="Bölmeleri Çöz" description="Tüm satır ve sütunların kilidini açarak tüm çalışma sayfasında kaydırma yapın." onClick={() => freeze(ctl, "none")} />
              ) : (
                <MenuItem icon={<TableFreezeColumnAndRow20Regular />} label="Bölmeleri Dondur" description="Geçerli seçime göre sayfanın kalanında kaydırırken satır ve sütunları görünür tutun." onClick={() => freeze(ctl, "panes")} />
              )}
              <MenuItem icon={<TableFreezeRow20Regular />} label="Üst Satırı Dondur" description="Sayfanın kalanında kaydırırken üst satırı görünür tutun." onClick={() => freeze(ctl, "topRow")} />
              <MenuItem icon={<TableFreezeColumn20Regular />} label="İlk Sütunu Dondur" description="Sayfanın kalanında kaydırırken ilk sütunu görünür tutun." onClick={() => freeze(ctl, "firstCol")} />
            </Menu>
          }
        />
      </RibbonGroup>
    </>
  );
}
