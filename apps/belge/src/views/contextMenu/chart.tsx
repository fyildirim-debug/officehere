// Grafik bağlam menüsü: Verileri Düzenle, Grafik Türünü Değiştir, resim yazısı ve sil.
import { DataBarHorizontal20Regular, DataBarVertical20Regular, DataLine20Regular, DataPie20Regular, ImageTable20Regular, TableEdit20Regular } from "@fluentui/react-icons";
import type { ReactElement } from "react";
import { NodeSelection } from "@tiptap/pm/state";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { CHART_TYPE_LABEL, type ChartData, type ChartType } from "../../editor/chart";
import { ClipboardItems, DeleteItem, type MenuDeps } from "./common";

const TYPE_ICON: Record<ChartType, ReactElement> = {
  column: <DataBarVertical20Regular />,
  bar: <DataBarHorizontal20Regular />,
  line: <DataLine20Regular />,
  pie: <DataPie20Regular />,
};

export function ChartMenu({ deps }: { deps: MenuDeps }) {
  const { openDialog, inBody } = useBelge();
  const { editor } = deps;
  const sel = editor.state.selection;
  const data = sel instanceof NodeSelection && sel.node.type.name === "chart" ? (sel.node.attrs.data as ChartData | null) : null;
  return (
    <Menu width={240} label="Grafik">
      <ClipboardItems editor={editor} />
      <MenuSeparator />
      <MenuItem icon={<TableEdit20Regular />} label="Verileri Düzenle..." onClick={() => openDialog("chart")} />
      <MenuItem
        icon={TYPE_ICON[data?.type ?? "column"]}
        label="Grafik Türünü Değiştir"
        disabled={!data}
        submenu={
          <Menu label="Grafik Türünü Değiştir">
            {(Object.keys(CHART_TYPE_LABEL) as ChartType[]).map((t) => (
              <MenuItem key={t} icon={TYPE_ICON[t]} label={CHART_TYPE_LABEL[t]} checked={data?.type === t} onClick={() => data && editor.chain().focus().insertChart({ ...data, type: t }).run()} />
            ))}
          </Menu>
        }
      />
      {inBody && <MenuItem icon={<ImageTable20Regular />} label="Resim Yazısı Ekle..." onClick={() => openDialog("caption")} />}
      <MenuSeparator />
      <DeleteItem editor={editor} />
    </Menu>
  );
}
