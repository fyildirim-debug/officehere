// "Yerleştir" grubu (Düzen sekmesi ve Resim Biçimi): Konum, Metni Kaydır, Öne Getir, Arkaya Gönder.
// Komutlar seçili resme uygulanır; resim seçili değilse devre dışıdır (Word'deki gibi).
import { PositionBackward20Regular, PositionForward20Regular, PositionToBack24Regular, PositionToFront24Regular, TextWrap24Regular, TextboxAlignMiddle24Regular } from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonGroup, RibbonMenuButton, RibbonSplitButton } from "@officehere/ui";
import { useBelge } from "../state/context";
import { POSITION_PRESETS, WRAP_LABEL, arrange, selectedImage, setPosition, setWrap, type WrapMode } from "../editor/imageWrap";

export function ArrangeGroup({ extra }: { extra?: React.ReactNode }) {
  const ctx = useBelge();
  const { editor } = ctx;
  const img = selectedImage(editor);
  const wrap = (img?.node.attrs.wrap as WrapMode | undefined) ?? "inline";
  const toPosition = (id: string) => {
    const preset = POSITION_PRESETS.find((p) => p.id === id);
    if (!preset || !img) return;
    const info = ctx.controller.info;
    const page = info.pages[ctx.controller.pageOf(img.pos)] ?? info.pages[0];
    const pagesEl = editor.view.dom.closest(".bl-pages") as HTMLElement | null;
    setPosition(editor, preset, page, pagesEl?.getBoundingClientRect().top ?? 0, ctx.view.zoom / 100);
  };
  return (
    <RibbonGroup label="Yerleştir" icon={<TextWrap24Regular />}>
      <RibbonMenuButton
        size="large"
        icon={<TextboxAlignMiddle24Regular />}
        label="Konum"
        disabled={!img}
        menu={
          <Menu width={300}>
            <MenuItem label="Metinle Aynı Hizada" checked={wrap === "inline"} onClick={() => setWrap(editor, "inline")} />
            <MenuHeader>Metin Kaydırma ile</MenuHeader>
            {POSITION_PRESETS.map((p) => (
              <MenuItem key={p.id} label={p.label} onClick={() => toPosition(p.id)} />
            ))}
          </Menu>
        }
      />
      <RibbonMenuButton
        size="large"
        icon={<TextWrap24Regular />}
        label="Metni Kaydır"
        disabled={!img}
        menu={
          <Menu width={240}>
            {(Object.keys(WRAP_LABEL) as WrapMode[]).map((w) => (
              <MenuItem key={w} label={WRAP_LABEL[w]} checked={wrap === w} onClick={() => setWrap(editor, w)} />
            ))}
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<PositionToFront24Regular />}
        label="Öne Getir"
        disabled={!img}
        onClick={() => arrange(editor, "forward")}
        menu={
          <Menu>
            <MenuItem icon={<PositionForward20Regular />} label="Öne Getir" onClick={() => arrange(editor, "forward")} />
            <MenuItem icon={<PositionToFront24Regular />} label="En Öne Getir" onClick={() => arrange(editor, "front")} />
            <MenuSeparator />
            <MenuItem label="Metnin Önüne Getir" onClick={() => setWrap(editor, "front")} />
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<PositionToBack24Regular />}
        label="Arkaya Gönder"
        disabled={!img}
        onClick={() => arrange(editor, "backward")}
        menu={
          <Menu>
            <MenuItem icon={<PositionBackward20Regular />} label="Arkaya Gönder" onClick={() => arrange(editor, "backward")} />
            <MenuItem icon={<PositionToBack24Regular />} label="En Arkaya Gönder" onClick={() => arrange(editor, "back")} />
            <MenuSeparator />
            <MenuItem label="Metnin Arkasına Gönder" onClick={() => setWrap(editor, "behind")} />
          </Menu>
        }
      />
      {extra}
    </RibbonGroup>
  );
}
