/** Sayfa Düzeni sekmesi: Sayfa Yapısı, Sığdırmak İçin Ölçeklendir, Sayfa Seçenekleri, Yerleştir. */
import { ArrowStepIn24Regular, DocumentLandscape24Regular, DocumentMargins24Regular, DocumentPageBreak24Regular, OrientationPortrait20Regular, DocumentLandscape20Regular, Print24Regular, PageFit24Regular, Layer24Regular, Layer20Regular, ArrowUp20Regular, ArrowDown20Regular, Grid24Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonStack, RibbonLabel } from "@officehere/ui";
import type { Controller } from "../controller";
import { MARGIN_PRESETS, PAPERS, setPage } from "../dialogs/PageSetupDialog";
import { reorderObject } from "../actions/objects";
import { ThemesGroup } from "./ThemesGroup";

const cmText = (inch: number) => (inch * 2.54).toLocaleString("tr-TR", { maximumFractionDigits: 2 });

export function PageLayoutTab({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const p = sheet.page;
  const sel = ctl.selObj;
  const viewToggle = (k: "showGridlines" | "showHeadings", v: boolean) => {
    ctl.ops.sheetOp(sheet, "Sayfa Seçenekleri", () => {
      sheet[k] = v;
    });
  };
  return (
    <>
      <ThemesGroup ctl={ctl} />
      <RibbonGroup label="Sayfa Yapısı" icon={<DocumentMargins24Regular />} onLauncher={() => ctl.openDialog("pageSetup")} launcherTip="Sayfa Yapısı">
        <RibbonMenuButton
          size="large"
          icon={<DocumentMargins24Regular />}
          label="Kenar Boşlukları"
          menu={
            <Menu width={280}>
              {MARGIN_PRESETS.map((m) => (
                <MenuItem
                  key={m.id}
                  checked={JSON.stringify(m.m) === JSON.stringify(p.margins)}
                  label={m.label}
                  description={`Üst: ${cmText(m.m.top)} cm  Alt: ${cmText(m.m.bottom)} cm  Sol: ${cmText(m.m.left)} cm  Sağ: ${cmText(m.m.right)} cm`}
                  onClick={() => setPage(ctl, { margins: m.m }, "Kenar Boşlukları")}
                />
              ))}
              <MenuSeparator />
              <MenuItem label="Özel Kenar Boşlukları..." onClick={() => ctl.openDialog("pageSetup", { tab: "margins" })} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentLandscape24Regular />}
          label="Yönlendirme"
          menu={
            <Menu width={160}>
              <MenuItem icon={<OrientationPortrait20Regular />} label="Dikey" checked={p.orientation === "portrait"} onClick={() => setPage(ctl, { orientation: "portrait" }, "Yönlendirme")} />
              <MenuItem icon={<DocumentLandscape20Regular />} label="Yatay" checked={p.orientation === "landscape"} onClick={() => setPage(ctl, { orientation: "landscape" }, "Yönlendirme")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<PageFit24Regular />}
          label="Boyut"
          menu={
            <Menu width={240}>
              {PAPERS.map((x) => (
                <MenuItem key={x.id} label={x.label} checked={p.paper === x.id} onClick={() => setPage(ctl, { paper: x.id }, "Kağıt Boyutu")} />
              ))}
              <MenuSeparator />
              <MenuItem label="Diğer Kağıt Boyutları..." onClick={() => ctl.openDialog("pageSetup", { tab: "page" })} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<Print24Regular />}
          label="Yazdırma Alanı"
          menu={
            <Menu width={220}>
              <MenuItem label="Yazdırma Alanını Belirle" onClick={() => setPage(ctl, { printArea: { ...ctl.range } }, "Yazdırma Alanı")} />
              <MenuItem label="Yazdırma Alanını Temizle" disabled={!p.printArea} onClick={() => setPage(ctl, { printArea: undefined }, "Yazdırma Alanı")} />
            </Menu>
          }
        />
        <RibbonButton
          size="large"
          icon={<DocumentPageBreak24Regular />}
          label="Sayfa Sonu Önizleme"
          checked={sheet.viewMode === "pageBreak"}
          description="Yazdırıldığında sayfa sonlarının nereye geleceğini gösterir."
          onClick={() => {
            sheet.viewMode = sheet.viewMode === "pageBreak" ? "normal" : "pageBreak";
            ctl.emit();
          }}
        />
      </RibbonGroup>
      <RibbonGroup label="Sığdırmak İçin Ölçeklendir" icon={<ArrowStepIn24Regular />} onLauncher={() => ctl.openDialog("pageSetup", { tab: "page" })}>
        <RibbonStack>
          <RibbonLabel>Genişlik:</RibbonLabel>
          <RibbonCheckbox label="Tek sayfaya sığdır" checked={!!p.fitToWidth} onChange={(v) => setPage(ctl, { fitToWidth: v || undefined }, "Ölçek")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Sayfa Seçenekleri" icon={<Grid24Regular />} onLauncher={() => ctl.openDialog("pageSetup", { tab: "sheet" })}>
        <RibbonStack>
          <RibbonLabel>Kılavuz Çizgileri</RibbonLabel>
          <RibbonCheckbox label="Görüntüle" checked={sheet.showGridlines} onChange={(v) => viewToggle("showGridlines", v)} />
          <RibbonCheckbox label="Yazdır" checked={p.printGridlines} onChange={(v) => setPage(ctl, { printGridlines: v }, "Kılavuz Çizgilerini Yazdır")} />
        </RibbonStack>
        <RibbonStack>
          <RibbonLabel>Başlıklar</RibbonLabel>
          <RibbonCheckbox label="Görüntüle" checked={sheet.showHeadings} onChange={(v) => viewToggle("showHeadings", v)} />
          <RibbonCheckbox label="Yazdır" checked={p.printHeadings} onChange={(v) => setPage(ctl, { printHeadings: v }, "Başlıkları Yazdır")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Yerleştir" icon={<Layer24Regular />}>
        <RibbonStack>
          <RibbonButton icon={<ArrowUp20Regular />} label="Öne Getir" disabled={!sel} onClick={() => sel && reorderObject(ctl, sel.kind, sel.id, true)} />
          <RibbonButton icon={<ArrowDown20Regular />} label="Arkaya Gönder" disabled={!sel} onClick={() => sel && reorderObject(ctl, sel.kind, sel.id, false)} />
          <RibbonButton icon={<Layer20Regular />} label="Seçim Bölmesi" checked={ctl.formatPane === "selection"} onClick={() => ((ctl.formatPane = ctl.formatPane === "selection" ? null : "selection"), ctl.emit())} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
