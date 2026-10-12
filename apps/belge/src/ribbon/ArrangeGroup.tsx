// "Yerleştir" grubu (Düzen sekmesi ve Resim Biçimi): Konum, Metni Kaydır, Öne Getir, Arkaya Gönder.
// Komutlar seçili resme uygulanır; resim seçili değilse devre dışıdır (Word'deki gibi).
// Seçenekler ne yaptığını gösteren kendi çizimlerimizle sunulur (packages/ui ArrangeIcons).
import { LayerIcon, Menu, MenuItem, MenuSeparator, PositionIcon, RibbonGroup, RibbonMenuButton, RibbonSplitButton, WrapIcon } from "@officehere/ui";
import { useBelge } from "../state/context";
import { WRAP_LABEL, arrange, selectedImage, setWrap, type WrapMode } from "../editor/imageWrap";
import { PositionMenu } from "./PositionMenu";

/** Menü çizimlerinin boyutu (px). */
const ICON = 32;

export function ArrangeGroup({ extra }: { extra?: React.ReactNode }) {
  const ctx = useBelge();
  const { editor } = ctx;
  const img = selectedImage(editor);
  const wrap = (img?.node.attrs.wrap as WrapMode | undefined) ?? "inline";
  return (
    <RibbonGroup label="Yerleştir" icon={<PositionIcon size={32} />}>
      <RibbonMenuButton
        size="large"
        icon={<PositionIcon size={32} />}
        label="Konum"
        description="Resmin sayfadaki yerini seçin; metin çevresinden akar."
        disabled={!img}
        menu={<PositionMenu />}
      />
      <RibbonMenuButton
        size="large"
        icon={<WrapIcon mode="square" size={32} />}
        label="Metni Kaydır"
        description="Metnin resmin çevresinden nasıl akacağını seçin."
        disabled={!img}
        menu={
          <Menu width={220} label="Metni Kaydır">
            {(Object.keys(WRAP_LABEL) as WrapMode[]).map((w) => (
              <MenuItem key={w} icon={<WrapIcon mode={w} size={ICON} />} label={WRAP_LABEL[w]} checked={wrap === w} onClick={() => setWrap(editor, w)} />
            ))}
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<LayerIcon kind="forward" size={32} />}
        label="Öne Getir"
        disabled={!img}
        onClick={() => arrange(editor, "forward")}
        menu={
          <Menu label="Öne Getir">
            <MenuItem icon={<LayerIcon kind="forward" size={ICON} />} label="Öne Getir" onClick={() => arrange(editor, "forward")} />
            <MenuItem icon={<LayerIcon kind="front" size={ICON} />} label="En Öne Getir" onClick={() => arrange(editor, "front")} />
            <MenuSeparator />
            <MenuItem icon={<WrapIcon mode="front" size={ICON} />} label="Metnin Önüne Getir" onClick={() => setWrap(editor, "front")} />
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<LayerIcon kind="backward" size={32} />}
        label="Arkaya Gönder"
        disabled={!img}
        onClick={() => arrange(editor, "backward")}
        menu={
          <Menu label="Arkaya Gönder">
            <MenuItem icon={<LayerIcon kind="backward" size={ICON} />} label="Arkaya Gönder" onClick={() => arrange(editor, "backward")} />
            <MenuItem icon={<LayerIcon kind="back" size={ICON} />} label="En Arkaya Gönder" onClick={() => arrange(editor, "back")} />
            <MenuSeparator />
            <MenuItem icon={<WrapIcon mode="behind" size={ICON} />} label="Metnin Arkasına Gönder" onClick={() => setWrap(editor, "behind")} />
          </Menu>
        }
      />
      {extra}
    </RibbonGroup>
  );
}
