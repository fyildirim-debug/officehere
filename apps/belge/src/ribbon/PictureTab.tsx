// Bağlamsal "Resim Biçimi" sekmesi.
import { NodeSelection } from "@tiptap/pm/state";
import {
  ArrowRotateClockwise20Regular,
  ImageAltText20Regular,
  ImageBorder32Regular,
  TextAlignCenter20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  RotateLeft20Regular,
  FlipHorizontal20Regular,
  FlipVertical20Regular,
} from "@fluentui/react-icons";
import { ColorPalette, Menu, MenuItem, MenuSeparator, RibbonButton, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { IMAGE_STYLES } from "../editor/image";

const CM = 96 / 2.54;

export function PictureTab() {
  const ctx = useBelge();
  const { editor } = ctx;
  const sel = editor.state.selection;
  const node = sel instanceof NodeSelection && sel.node.type.name === "image" ? sel.node : null;
  if (!node) return null;
  const a = node.attrs;
  const pos = sel.from;
  const update = (attrs: Record<string, unknown>) => {
    const tr = editor.state.tr.setNodeMarkup(pos, undefined, { ...a, ...attrs });
    tr.setSelection(NodeSelection.create(tr.doc, pos));
    editor.view.dispatch(tr);
  };
  const w = (a.width as number | null) ?? 200;
  const h = (a.height as number | null) ?? 150;
  const align = (v: "left" | "center" | "right") => editor.chain().focus().setTextAlign(v).setNodeSelection(pos).run();

  return (
    <>
      <RibbonGroup label="Resim Stilleri" icon={<ImageBorder32Regular />}>
        <RibbonGallery
          columns={6}
          itemWidth={56}
          itemHeight={50}
          selectedId={(a.imgStyle as string) ?? "none"}
          onSelect={(id) => update({ imgStyle: id === "none" ? null : id })}
          items={IMAGE_STYLES.map((s) => ({
            id: s.id,
            label: s.label,
            preview: (
              <span className="bl-imgstyle-prev" data-img-style={s.id}>
                <i />
              </span>
            ),
          }))}
        />
        <RibbonMenuButton
          icon={<ImageBorder32Regular />}
          label="Resim Kenarlığı"
          menu={
            <div className="oh-menu oh-menu-palette">
              <ColorPalette noneLabel="Anahat Yok" onSelect={(c) => update({ border: c ? `2px solid ${c}` : null })} />
            </div>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Erişilebilirlik" icon={<ImageAltText20Regular />}>
        <RibbonButton size="large" icon={<ImageAltText20Regular />} label="Alternatif Metin" onClick={() => ctx.openDialog("altText")} />
      </RibbonGroup>
      <RibbonGroup label="Yerleştir" icon={<ArrowRotateClockwise20Regular />}>
        <RibbonStack>
          <RibbonMenuButton
            icon={<ArrowRotateClockwise20Regular />}
            label="Döndür"
            menu={
              <Menu>
                <MenuItem icon={<ArrowRotateClockwise20Regular />} label="Sağa 90° Döndür" onClick={() => update({ rotate: (((a.rotate as number) ?? 0) + 90) % 360 })} />
                <MenuItem icon={<RotateLeft20Regular />} label="Sola 90° Döndür" onClick={() => update({ rotate: (((a.rotate as number) ?? 0) + 270) % 360 })} />
                <MenuItem icon={<FlipVertical20Regular />} label="Dikey Çevir" onClick={() => update({ flipV: !a.flipV })} />
                <MenuItem icon={<FlipHorizontal20Regular />} label="Yatay Çevir" onClick={() => update({ flipH: !a.flipH })} />
                <MenuSeparator />
                <MenuItem label="Döndürmeyi Sıfırla" onClick={() => update({ rotate: 0, flipH: false, flipV: false })} />
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<TextAlignLeft20Regular />}
            label="Hizala"
            menu={
              <Menu>
                <MenuItem icon={<TextAlignLeft20Regular />} label="Sola Hizala" onClick={() => align("left")} />
                <MenuItem icon={<TextAlignCenter20Regular />} label="Ortala" onClick={() => align("center")} />
                <MenuItem icon={<TextAlignRight20Regular />} label="Sağa Hizala" onClick={() => align("right")} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Boyut" onLauncher={() => ctx.openDialog("imageSize")}>
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={h / CM} step={0.1} min={0.1} max={60} unit="cm" decimals={2} width={78} onChange={(v) => update({ height: Math.round(v * CM), width: Math.round((v * CM * w) / h) })} />
          <RibbonSpinner label="Genişlik:" value={w / CM} step={0.1} min={0.1} max={60} unit="cm" decimals={2} width={78} onChange={(v) => update({ width: Math.round(v * CM), height: Math.round((v * CM * h) / w) })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
