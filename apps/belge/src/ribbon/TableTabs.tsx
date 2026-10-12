// Bağlamsal tablo sekmeleri: Tablo Tasarımı ve Düzen.
import {
  BorderAll20Regular,
  BorderNone20Regular,
  Delete32Regular,
  PaintBucket20Regular,
  TableCellsMerge20Regular,
  TableCellsSplit20Regular,
  TableDeleteColumn20Regular,
  TableDeleteRow20Regular,
  TableInsertColumn20Regular,
  TableInsertRow20Regular,
  TableResizeColumn20Regular,
  TableSettings20Regular,
  TableSimple20Regular,
  Grid20Regular,
  Cursor20Regular,
  Delete20Regular,
  TableArrowUp20Regular,
} from "@fluentui/react-icons";
import { ColorSplitButton, Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonRow, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import { useState } from "react";
import { useBelge } from "../state/context";
import { TABLE_STYLES } from "../editor/table";
import { alignCells, autoFitContents, distributeColumns, selectTablePart, tableInfo } from "./tableUtil";
import { columnWidthAt } from "../state/pageSetup";

function TablePreview({ id }: { id: string }) {
  const s = TABLE_STYLES.find((x) => x.id === id)!;
  return (
    <span className="bl-tsprev">
      {[0, 1, 2, 3].map((r) => (
        <span key={r} className="bl-tsprev-row">
          {[0, 1, 2, 3].map((c) => (
            <i
              key={c}
              style={{
                background: r === 0 && s.header ? s.header : r % 2 === 1 && s.band ? s.band : "#fff",
                borderColor: s.noBorders ? "#ddd" : s.border,
                borderStyle: s.noBorders ? "dashed" : "solid",
              }}
            />
          ))}
        </span>
      ))}
    </span>
  );
}

export function TableDesignTab() {
  const ctx = useBelge();
  const { editor } = ctx;
  const info = tableInfo(editor);
  const a = info?.node.attrs ?? {};
  const [shade, setShade] = useState<string | null>("#D9D9D9");
  const set = (attrs: Record<string, unknown>) => editor.chain().focus().updateAttributes("table", attrs).run();
  return (
    <>
      <RibbonGroup label="Tablo Stili Seçenekleri" icon={<TableSimple20Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Üst Bilgi Satırı" checked={!!a.headerRow} onChange={(v) => set({ headerRow: v })} />
          <RibbonCheckbox label="Toplam Satırı" checked={!!a.totalRow} onChange={(v) => set({ totalRow: v })} />
          <RibbonCheckbox label="Bantlı Satırlar" checked={!!a.bandedRows} onChange={(v) => set({ bandedRows: v })} />
        </RibbonStack>
        <RibbonStack>
          <RibbonCheckbox label="İlk Sütun" checked={!!a.firstColumn} onChange={(v) => set({ firstColumn: v })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Stilleri" icon={<TableSimple20Regular />}>
        <RibbonGallery
          columns={5}
          itemWidth={66}
          itemHeight={50}
          selectedId={(a.tableStyle as string) ?? "grid"}
          onSelect={(id) => set({ tableStyle: id })}
          items={TABLE_STYLES.map((s) => ({ id: s.id, label: s.label, preview: <TablePreview id={s.id} /> }))}
        />
        <RibbonStack>
          <ColorSplitButton size="large" icon={<PaintBucket20Regular />} label="Gölgelendirme" tip="Gölgelendirme" color={shade} noneLabel="Renk Yok" onColorChange={setShade} onApply={(c) => editor.chain().focus().setCellAttribute("background", c).run()} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Kenarlıklar" icon={<BorderAll20Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<BorderAll20Regular />}
          label="Kenarlıklar"
          onClick={() => set({ tableStyle: "grid" })}
          menu={
            <Menu>
              <MenuItem icon={<BorderAll20Regular />} label="Tüm Kenarlıklar" onClick={() => set({ tableStyle: "grid" })} />
              <MenuItem icon={<BorderNone20Regular />} label="Kenarlık Yok" onClick={() => set({ tableStyle: "none" })} />
            </Menu>
          }
        />
      </RibbonGroup>
    </>
  );
}

export function TableLayoutTab() {
  const ctx = useBelge();
  const { editor, settings } = ctx;
  const chain = () => editor.chain().focus();
  const m = { contentW: columnWidthAt(ctx.mainEditor, settings.page) };
  const alignBtn = (v: "top" | "middle" | "bottom", h: "left" | "center" | "right", tip: string) => (
    <RibbonButton size="small" icon={<span className={`bl-cellalign v-${v} h-${h}`}><i /><i /><i /></span>} tip={tip} onClick={() => alignCells(editor, v, h)} />
  );
  return (
    <>
      <RibbonGroup label="Tablo" icon={<TableSimple20Regular />}>
        <RibbonStack>
          <RibbonMenuButton
            icon={<Cursor20Regular />}
            label="Seç"
            menu={
              <Menu>
                <MenuItem label="Hücreyi Seç" onClick={() => selectTablePart(editor, "cell")} />
                <MenuItem label="Sütunu Seç" onClick={() => selectTablePart(editor, "col")} />
                <MenuItem label="Satırı Seç" onClick={() => selectTablePart(editor, "row")} />
                <MenuItem label="Tabloyu Seç" onClick={() => selectTablePart(editor, "table")} />
              </Menu>
            }
          />
          <RibbonButton icon={<Grid20Regular />} label="Kılavuz Çizgilerini Göster" checked={ctx.view.gridlines} onClick={() => ctx.setView({ gridlines: !ctx.view.gridlines })} />
          <RibbonButton icon={<TableSettings20Regular />} label="Özellikler" onClick={() => ctx.openDialog("tableProps")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Satırlar ve Sütunlar" icon={<TableInsertRow20Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Delete32Regular />}
          label="Sil"
          menu={
            <Menu>
              <MenuItem icon={<TableDeleteColumn20Regular />} label="Sütunları Sil" onClick={() => chain().deleteColumn().run()} />
              <MenuItem icon={<TableDeleteRow20Regular />} label="Satırları Sil" onClick={() => chain().deleteRow().run()} />
              <MenuItem icon={<Delete20Regular />} label="Tabloyu Sil" onClick={() => chain().deleteTable().run()} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<TableArrowUp20Regular />} label="Üste Ekle" onClick={() => chain().addRowBefore().run()} />
        <RibbonStack>
          <RibbonButton icon={<TableInsertRow20Regular />} label="Alta Ekle" onClick={() => chain().addRowAfter().run()} />
          <RibbonButton icon={<TableInsertColumn20Regular />} label="Sola Ekle" onClick={() => chain().addColumnBefore().run()} />
          <RibbonButton icon={<TableInsertColumn20Regular />} label="Sağa Ekle" onClick={() => chain().addColumnAfter().run()} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Birleştir" icon={<TableCellsMerge20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<TableCellsMerge20Regular />} label="Hücreleri Birleştir" disabled={!ctx.fmt.cellSelection} onClick={() => chain().mergeCells().run()} />
          <RibbonButton icon={<TableCellsSplit20Regular />} label="Hücreleri Böl" onClick={() => ctx.openDialog("splitCells")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Hücre Boyutu" icon={<TableResizeColumn20Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<TableResizeColumn20Regular />}
          label="Otomatik Sığdır"
          menu={
            <Menu>
              <MenuItem label="İçeriği Otomatik Sığdır" onClick={() => autoFitContents(editor)} />
              <MenuItem label="Pencereyi Otomatik Sığdır" onClick={() => distributeColumns(editor, m.contentW)} />
              <MenuItem label="Sabit Sütun Genişliği" onClick={() => distributeColumns(editor, m.contentW)} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonButton icon={<TableResizeColumn20Regular />} label="Sütunları Dağıt" onClick={() => distributeColumns(editor, m.contentW)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Hizalama" icon={<TableSimple20Regular />}>
        <RibbonStack>
          <RibbonRow>
            {alignBtn("top", "left", "Üste Sola Hizala")}
            {alignBtn("top", "center", "Üste Ortala")}
            {alignBtn("top", "right", "Üste Sağa Hizala")}
          </RibbonRow>
          <RibbonRow>
            {alignBtn("middle", "left", "Ortada Sola Hizala")}
            {alignBtn("middle", "center", "Ortala")}
            {alignBtn("middle", "right", "Ortada Sağa Hizala")}
          </RibbonRow>
          <RibbonRow>
            {alignBtn("bottom", "left", "Alta Sola Hizala")}
            {alignBtn("bottom", "center", "Alta Ortala")}
            {alignBtn("bottom", "right", "Alta Sağa Hizala")}
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
