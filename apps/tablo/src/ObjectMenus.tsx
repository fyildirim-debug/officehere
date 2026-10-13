/** Yüzen nesne sağ tık menüleri: resim, şekil/metin kutusu ve grafik için ayrı içerik. */
import {
  ArrowMove20Regular,
  ChartMultiple20Regular,
  ClipboardPaste20Regular,
  Copy20Regular,
  Cut20Regular,
  DataBarVertical20Regular,
  Delete20Regular,
  Edit20Regular,
  ImageEdit20Regular,
  PaintBucket20Regular,
  Save20Regular,
} from "@fluentui/react-icons";
import { LayerIcon, Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { Controller } from "./controller";
import { deleteObject } from "./charts/FloatingLayer";
import { copyObject, hasObjectClip, pasteObject, startShapeEdit, stepObject } from "./actions/objectClip";
import { reorderObject } from "./actions/objects";
import { replaceImage, saveImageAs } from "./actions/imageFile";

type ObjKind = "chart" | "image" | "shape";

interface P {
  ctl: Controller;
  kind: ObjKind;
  id: string;
}

const select = (ctl: Controller, kind: ObjKind, id: string) => {
  ctl.selObj = { kind, id };
};

function ClipItems({ ctl, kind, id }: P) {
  return (
    <>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={() => (select(ctl, kind, id), copyObject(ctl, true))} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={() => (select(ctl, kind, id), copyObject(ctl, false))} />
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" disabled={!hasObjectClip()} onClick={() => pasteObject(ctl)} />
    </>
  );
}

/** Öne Getir › / Arkaya Gönder ›: en öne/arkaya ya da bir adım (tüm nesne türleri arasında). */
function OrderItems({ ctl, kind, id }: P) {
  return (
    <>
      <MenuItem
        icon={<LayerIcon kind="forward" />}
        label="Öne Getir"
        submenu={
          <Menu>
            <MenuItem icon={<LayerIcon kind="front" />} label="En Öne Getir" onClick={() => reorderObject(ctl, kind, id, true)} />
            <MenuItem icon={<LayerIcon kind="forward" />} label="Öne Getir" onClick={() => stepObject(ctl, kind, id, true)} />
          </Menu>
        }
      />
      <MenuItem
        icon={<LayerIcon kind="backward" />}
        label="Arkaya Gönder"
        submenu={
          <Menu>
            <MenuItem icon={<LayerIcon kind="back" />} label="En Arkaya Gönder" onClick={() => reorderObject(ctl, kind, id, false)} />
            <MenuItem icon={<LayerIcon kind="backward" />} label="Arkaya Gönder" onClick={() => stepObject(ctl, kind, id, false)} />
          </Menu>
        }
      />
    </>
  );
}

function DeleteItem({ ctl, kind, id }: P) {
  return <MenuItem icon={<Delete20Regular />} label="Sil" shortcut="Delete" onClick={() => deleteObject(ctl, kind, id)} />;
}

function ImageMenu(p: P) {
  const { ctl, id } = p;
  return (
    <Menu width={240}>
      <ClipItems {...p} />
      <MenuSeparator />
      <MenuItem icon={<ImageEdit20Regular />} label="Resmi Değiştir..." onClick={() => replaceImage(ctl, id)} />
      <MenuSeparator />
      <OrderItems {...p} />
      <MenuSeparator />
      <MenuItem icon={<Save20Regular />} label="Resim Olarak Kaydet..." onClick={() => void saveImageAs(ctl, id)} />
      <MenuSeparator />
      <DeleteItem {...p} />
    </Menu>
  );
}

function ShapeMenu(p: P) {
  const { ctl, id } = p;
  const shape = ctl.sheet.shapes.find((s) => s.id === id);
  const line = shape?.kind === "line";
  return (
    <Menu width={240}>
      <ClipItems {...p} />
      <MenuSeparator />
      {!line && (
        <>
          <MenuItem icon={<Edit20Regular />} label="Metni Düzenle" onClick={() => startShapeEdit(ctl, id, "")} />
          <MenuSeparator />
        </>
      )}
      <OrderItems {...p} />
      {!line && (
        <>
          <MenuSeparator />
          <MenuItem icon={<PaintBucket20Regular />} label="Şekli Biçimlendir..." onClick={() => ctl.openDialog("shapeText", { id })} />
        </>
      )}
      <MenuSeparator />
      <DeleteItem {...p} />
    </Menu>
  );
}

function ChartMenu(p: P) {
  const { ctl, kind, id } = p;
  return (
    <Menu width={240}>
      <ClipItems {...p} />
      <MenuSeparator />
      <MenuItem icon={<ChartMultiple20Regular />} label="Grafik Türünü Değiştir..." onClick={() => ctl.openDialog("chartType")} />
      <MenuItem icon={<DataBarVertical20Regular />} label="Veri Seç..." onClick={() => ctl.openDialog("chartData")} />
      <MenuItem icon={<ArrowMove20Regular />} label="Grafiği Taşı..." onClick={() => ctl.openDialog("moveChart")} />
      <MenuSeparator />
      <OrderItems {...p} />
      <MenuSeparator />
      <MenuItem icon={<PaintBucket20Regular />} label="Grafik Alanını Biçimlendir..." onClick={() => (select(ctl, kind, id), ctl.openDialog("chartFormat"))} />
      <MenuSeparator />
      <DeleteItem {...p} />
    </Menu>
  );
}

export function ObjectMenu(p: P) {
  if (p.kind === "image") return <ImageMenu {...p} />;
  if (p.kind === "shape") return <ShapeMenu {...p} />;
  return <ChartMenu {...p} />;
}
