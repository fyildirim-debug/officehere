// Tablo hücresi bağlam menüsü: ekle/sil, birleştir/böl, hücre hizalaması, otomatik sığdır, kenarlıklar,
// Tablo Özellikleri ve resim yazısı. Yazı tipi/paragraf komutları bu menüde yer almaz (Word'deki gibi).
import { CellSelection } from "@tiptap/pm/tables";
import {
  BorderAll20Regular,
  BorderNone20Regular,
  Delete20Regular,
  ImageTable20Regular,
  TableCellsMerge20Regular,
  TableCellsSplit20Regular,
  TableDeleteColumn20Regular,
  TableDeleteRow20Regular,
  TableInsertColumn20Regular,
  TableInsertRow20Regular,
  TableResizeColumn20Regular,
  TableSettings20Regular,
  AlignCenterVertical20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuGallery, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { columnWidthAt } from "../../state/pageSetup";
import { alignCells, autoFitContents, distributeColumns, tableInfo } from "../../ribbon/tableUtil";
import { ClipboardItems, LinkCommentItems, type MenuDeps } from "./common";

type VAlign = "top" | "middle" | "bottom";
type HAlign = "left" | "center" | "right";

const ALIGN_LABEL: Record<`${VAlign}-${HAlign}`, string> = {
  "top-left": "Üste Sola Hizala",
  "top-center": "Üste Ortala",
  "top-right": "Üste Sağa Hizala",
  "middle-left": "Ortada Sola Hizala",
  "middle-center": "Ortala",
  "middle-right": "Ortada Sağa Hizala",
  "bottom-left": "Alta Sola Hizala",
  "bottom-center": "Alta Ortala",
  "bottom-right": "Alta Sağa Hizala",
};

/** İmleçteki hücrenin hizası (dikey: hücre özniteliği, yatay: paragraf). */
function currentAlign(deps: MenuDeps): `${VAlign}-${HAlign}` {
  const { $from } = deps.editor.state.selection;
  let v: VAlign = "top";
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === "tableCell" || n.type.name === "tableHeader") {
      v = (n.attrs.verticalAlign as VAlign | null) ?? "top";
      break;
    }
  }
  const h = ($from.parent.attrs.textAlign as HAlign | null) ?? "left";
  return `${v}-${h === "center" || h === "right" ? h : "left"}`;
}

export function TableMenu({ deps }: { deps: MenuDeps }) {
  const ctx = useBelge();
  const { openDialog, settings, inBody } = ctx;
  const { editor } = deps;
  const chain = () => editor.chain().focus();
  const cellSel = editor.state.selection instanceof CellSelection;
  const info = tableInfo(editor);
  const style = (info?.node.attrs.tableStyle as string | undefined) ?? "grid";
  const contentW = () => columnWidthAt(ctx.mainEditor, settings.page);
  const setTable = (attrs: Record<string, unknown>) => chain().updateAttributes("table", attrs).run();
  return (
    <Menu width={240} label="Tablo">
      <ClipboardItems editor={editor} />
      <MenuSeparator />
      <MenuItem
        label="Ekle"
        icon={<TableInsertRow20Regular />}
        submenu={
          <Menu label="Ekle">
            <MenuItem icon={<TableInsertColumn20Regular />} label="Sola Sütun Ekle" onClick={() => chain().addColumnBefore().run()} />
            <MenuItem icon={<TableInsertColumn20Regular />} label="Sağa Sütun Ekle" onClick={() => chain().addColumnAfter().run()} />
            <MenuItem icon={<TableInsertRow20Regular />} label="Üste Satır Ekle" onClick={() => chain().addRowBefore().run()} />
            <MenuItem icon={<TableInsertRow20Regular />} label="Alta Satır Ekle" onClick={() => chain().addRowAfter().run()} />
          </Menu>
        }
      />
      <MenuItem
        label="Sil"
        icon={<Delete20Regular />}
        submenu={
          <Menu label="Sil">
            <MenuItem icon={<TableDeleteColumn20Regular />} label="Sütunu Sil" onClick={() => chain().deleteColumn().run()} />
            <MenuItem icon={<TableDeleteRow20Regular />} label="Satırı Sil" onClick={() => chain().deleteRow().run()} />
            <MenuItem icon={<Delete20Regular />} label="Tabloyu Sil" onClick={() => chain().deleteTable().run()} />
          </Menu>
        }
      />
      <MenuItem icon={<TableCellsMerge20Regular />} label="Hücreleri Birleştir" disabled={!cellSel} onClick={() => chain().mergeCells().run()} />
      <MenuItem icon={<TableCellsSplit20Regular />} label="Hücreleri Böl..." onClick={() => openDialog("splitCells")} />
      <MenuSeparator />
      <MenuItem
        label="Hücre Hizalaması"
        icon={<AlignCenterVertical20Regular />}
        submenu={
          <Menu label="Hücre Hizalaması">
            <MenuGallery
              label="Hücre Hizalaması"
              columns={3}
              itemWidth={34}
              itemHeight={30}
              selectedId={currentAlign(deps)}
              onSelect={(id) => {
                const [v, h] = id.split("-") as [VAlign, HAlign];
                alignCells(editor, v, h);
              }}
              items={(Object.keys(ALIGN_LABEL) as (keyof typeof ALIGN_LABEL)[]).map((id) => {
                const [v, h] = id.split("-");
                return {
                  id,
                  label: ALIGN_LABEL[id],
                  preview: (
                    <span className={`bl-cellalign v-${v} h-${h}`}>
                      <i />
                      <i />
                      <i />
                    </span>
                  ),
                };
              })}
            />
          </Menu>
        }
      />
      <MenuItem
        label="Otomatik Sığdır"
        icon={<TableResizeColumn20Regular />}
        submenu={
          <Menu label="Otomatik Sığdır">
            <MenuItem label="İçeriği Otomatik Sığdır" onClick={() => autoFitContents(editor)} />
            <MenuItem label="Pencereyi Otomatik Sığdır" onClick={() => distributeColumns(editor, contentW())} />
            <MenuItem label="Sabit Sütun Genişliği" onClick={() => distributeColumns(editor, contentW())} />
          </Menu>
        }
      />
      <MenuItem
        label="Kenarlıklar"
        icon={<BorderAll20Regular />}
        submenu={
          <Menu label="Kenarlıklar">
            <MenuItem icon={<BorderAll20Regular />} label="Tüm Kenarlıklar" checked={style === "grid"} onClick={() => setTable({ tableStyle: "grid" })} />
            <MenuItem icon={<BorderNone20Regular />} label="Kenarlık Yok" checked={style === "none"} onClick={() => setTable({ tableStyle: "none" })} />
          </Menu>
        }
      />
      <MenuItem icon={<TableSettings20Regular />} label="Tablo Özellikleri..." onClick={() => openDialog("tableProps")} />
      {inBody && <MenuItem icon={<ImageTable20Regular />} label="Resim Yazısı Ekle..." onClick={() => openDialog("caption")} />}
      <MenuSeparator />
      <LinkCommentItems editor={editor} openLink={deps.openLink} />
    </Menu>
  );
}
