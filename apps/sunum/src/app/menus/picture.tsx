/** Resim ve ortam (video/ses) sağ tık menüleri. */
import { Crop20Regular, ImageAltText20Regular, ImageEdit20Regular, Play20Regular, Save20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { PicShape } from "../../model/types";
import { replacePicture, savePictureAs } from "../pictureCmds";
import { ClipItems, CommentItem, DeleteItem, FormatItem, GroupItem, LinkItems, OrderItems, SizePosItem, type MenuDeps } from "./common";

export function pictureMenu({ a, cmd, openDialog }: MenuDeps, pic: PicShape) {
  const cropping = a.s.cropId === pic.id;
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <MenuItem
        icon={<ImageEdit20Regular />}
        label="Resmi Değiştir"
        submenu={
          <Menu>
            <MenuItem label="Dosyadan..." onClick={() => replacePicture(a)} />
          </Menu>
        }
      />
      <MenuItem icon={<Crop20Regular />} label="Kırp" checked={cropping} onClick={() => a.store.set({ cropId: cropping ? null : pic.id, editing: null })} />
      {pic.zoom && (
        <MenuItem
          label="Yakınlaştırmaya Dön"
          checked={!!pic.zoom.ret}
          onClick={() => a.updateShapes((s) => (s.kind === "pic" && s.zoom ? { ...s, zoom: { ...s.zoom, ret: s.zoom.ret ? undefined : true } } : s), [pic.id])}
        />
      )}
      <MenuSeparator />
      <GroupItem a={a} />
      <OrderItems a={a} />
      <MenuSeparator />
      <LinkItems a={a} cmd={cmd} />
      <MenuItem icon={<Save20Regular />} label="Resim Olarak Kaydet..." onClick={() => void savePictureAs(a, pic)} />
      <MenuItem icon={<ImageAltText20Regular />} label="Alternatif Metni Görüntüle..." onClick={cmd.altText} />
      <MenuSeparator />
      <SizePosItem openDialog={openDialog} />
      <FormatItem a={a} label="Resmi Biçimlendir..." />
      <CommentItem a={a} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}

/** Video/ses: önizleme ve kapak resmiyle ilgili komutlar. */
export function mediaMenu({ a, cmd, openDialog }: MenuDeps, pic: PicShape) {
  const audio = pic.media?.kind === "audio";
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <MenuItem icon={<Play20Regular />} label="Yürüt" onClick={cmd.playMedia} />
      <MenuSeparator />
      <OrderItems a={a} />
      <MenuSeparator />
      <MenuItem icon={<ImageAltText20Regular />} label="Alternatif Metni Görüntüle..." onClick={cmd.altText} />
      <SizePosItem openDialog={openDialog} />
      <FormatItem a={a} label={audio ? "Sesi Biçimlendir..." : "Videoyu Biçimlendir..."} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}
