// Resim bağlam menüsü: resmi değiştir/kaydet, resim yazısı, Metni Kaydır, katman sırası, boyut, alternatif
// metin, Resim Biçimi, bağlantı/açıklama ve sil. Komutlar seçili resme (NodeSelection) uygulanır.
import { saveFileAs } from "@officehere/core";
import { ImageAltText20Regular, ImageArrowCounterclockwise20Regular, ImageEdit20Regular, ImageTable20Regular, ResizeImage20Regular, SaveImage20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { selectedImage, type WrapMode } from "../../editor/imageWrap";
import { ClipboardItems, DeleteItem, LayerItems, LinkCommentItems, WrapItem, type MenuDeps } from "./common";

const EXT: Record<string, string> = { "image/jpeg": "jpg", "image/svg+xml": "svg", "image/x-icon": "ico" };

/** Resmi özgün verisiyle dosyaya kaydeder (masaüstünde kaydetme penceresi, tarayıcıda indirme). */
export async function saveImageAs(src: string, alt: string | null): Promise<void> {
  const blob = await (await fetch(src)).blob();
  const ext = EXT[blob.type] ?? (/^image\/([a-z0-9]+)$/.exec(blob.type)?.[1] || "png");
  const base = (alt ?? "").replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 60) || "Resim";
  await saveFileAs(`${base}.${ext}`, [{ name: `${ext.toLocaleUpperCase("tr-TR")} Resmi`, extensions: [ext] }], blob);
}

export function ImageMenu({ deps }: { deps: MenuDeps }) {
  const { openDialog, showTab, notify, inBody } = useBelge();
  const { editor } = deps;
  const img = selectedImage(editor);
  const wrap = (img?.node.attrs.wrap as WrapMode | undefined) ?? "inline";
  const src = (img?.node.attrs.src as string | undefined) ?? "";
  const save = () => saveImageAs(src, (img?.node.attrs.alt as string | null) ?? null).catch((e: Error) => notify(`Resim kaydedilemedi: ${e.message}`, "error"));
  return (
    <Menu width={250} label="Resim">
      <ClipboardItems editor={editor} />
      <MenuSeparator />
      <MenuItem icon={<ImageArrowCounterclockwise20Regular />} label="Resmi Değiştir..." onClick={deps.replaceImage} />
      <MenuItem icon={<SaveImage20Regular />} label="Resim Olarak Kaydet..." disabled={!src} onClick={() => void save()} />
      {inBody && <MenuItem icon={<ImageTable20Regular />} label="Resim Yazısı Ekle..." onClick={() => openDialog("caption")} />}
      <MenuSeparator />
      <WrapItem editor={editor} wrap={wrap} />
      <LayerItems editor={editor} wrap={wrap} />
      <MenuSeparator />
      <MenuItem icon={<ResizeImage20Regular />} label="Boyut ve Konum..." onClick={() => openDialog("imageSize")} />
      <MenuItem icon={<ImageAltText20Regular />} label="Alternatif Metni Görüntüle..." onClick={() => openDialog("altText")} />
      <MenuItem icon={<ImageEdit20Regular />} label="Resmi Biçimlendir..." onClick={() => showTab("picture")} />
      <MenuSeparator />
      <LinkCommentItems editor={editor} openLink={deps.openLink} />
      <MenuSeparator />
      <DeleteItem editor={editor} />
    </Menu>
  );
}
