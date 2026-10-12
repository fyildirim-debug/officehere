/** Sağ tık menüleri: hücre, satır/sütun başlığı, sayfa sekmesi, yüzen nesne. */
import {
  ArrowSortDownLines20Regular,
  ArrowSortUpLines20Regular,
  ClipboardPaste20Regular,
  Copy20Regular,
  Cut20Regular,
  Delete20Regular,
  Filter20Regular,
  Link20Regular,
  NoteAdd20Regular,
  TableDeleteColumn20Regular,
  TableDeleteRow20Regular,
  TableEdit20Regular,
  TableInsertColumn20Regular,
  TableInsertRow20Regular,
} from "@fluentui/react-icons";
import { ColorPalette, Menu, MenuHeader, MenuItem, MenuSeparator } from "@officehere/ui";
import { paletteOf } from "./model/theme";
import type { Controller } from "./controller";
import * as E from "./actions/edit";
import * as D from "./actions/data";
import { displayText } from "./model/display";
import { deleteObject } from "./charts/FloatingLayer";
import { copyObject, hasObjectClip, pasteObject, startShapeEdit } from "./actions/objectClip";
import { reorderObject } from "./actions/objects";
import { requestDeleteSheet } from "./dialogs/MiscDialogs";
import { openLink } from "./grid/Grid";
import type { TableDef } from "./model/types";
import { insertPageBreak, resetPageBreaks } from "./grid/pageBreakDrag";
import { setPage } from "./dialogs/PageSetupDialog";
import { deleteTableCols, deleteTableRows, insertTableCols, insertTableRows, selectTablePart, tableSpan } from "./actions/tableEdit";
import { ContextMiniBar } from "./ContextMiniBar";

function copy(ctl: Controller, cut: boolean) {
  const d = E.copySelection(ctl, cut);
  if (d) void E.writeSystemClipboard(d);
}

function PasteItems({ ctl }: { ctl: Controller }) {
  return (
    <>
      <MenuHeader>Yapıştırma Seçenekleri:</MenuHeader>
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" onClick={() => E.pasteFromButton(ctl)} />
      <MenuItem label="Değerler" onClick={() => E.pasteFromButton(ctl, { what: "values" })} />
      <MenuItem label="Formüller" onClick={() => E.pasteFromButton(ctl, { what: "formulas" })} />
      <MenuItem label="Devrik Dönüşüm" onClick={() => E.pasteFromButton(ctl, { what: "all", transpose: true })} />
      <MenuItem label="Biçimlendirme" onClick={() => E.pasteFromButton(ctl, { what: "formats" })} />
      <MenuItem label="Özel Yapıştır..." onClick={() => ctl.openDialog("pasteSpecial")} />
    </>
  );
}

/** Tablo hücresinde: Ekle / Sil / Seç alt menüleri ve Tablo komutları. */
function TableItems({ ctl, table }: { ctl: Controller; table: TableDef }) {
  const span = tableSpan(ctl, table);
  const lastDataRow = table.ref.r2 - (table.totalRow ? 1 : 0);
  return (
    <>
      <MenuItem
        label="Ekle"
        submenu={
          <Menu width={230}>
            <MenuItem icon={<TableInsertColumn20Regular />} label="Tablo Sütunları Sola" onClick={() => insertTableCols(ctl, table)} />
            {span.c2 === table.ref.c2 && <MenuItem label="Tablo Sütunları Sağa" onClick={() => insertTableCols(ctl, table, true)} />}
            <MenuItem icon={<TableInsertRow20Regular />} label="Tablo Satırları Yukarı" onClick={() => insertTableRows(ctl, table)} />
            {span.r2 === lastDataRow && <MenuItem label="Tablo Satırları Aşağı" onClick={() => insertTableRows(ctl, table, true)} />}
          </Menu>
        }
      />
      <MenuItem
        label="Sil"
        submenu={
          <Menu width={200}>
            <MenuItem icon={<TableDeleteColumn20Regular />} label="Tablo Sütunları" onClick={() => deleteTableCols(ctl, table)} />
            <MenuItem icon={<TableDeleteRow20Regular />} label="Tablo Satırları" onClick={() => deleteTableRows(ctl, table)} />
          </Menu>
        }
      />
      <MenuItem
        label="Seç"
        submenu={
          <Menu width={210}>
            <MenuItem label="Tablo Sütunu Verileri" onClick={() => selectTablePart(ctl, table, "colData")} />
            <MenuItem label="Tüm Tablo Sütunu" onClick={() => selectTablePart(ctl, table, "col")} />
            <MenuItem label="Tablo Satırı" onClick={() => selectTablePart(ctl, table, "row")} />
          </Menu>
        }
      />
      <MenuItem
        label="Tablo"
        submenu={
          <Menu width={210}>
            <MenuItem label="Toplam Satırı" checked={table.totalRow} onClick={() => D.updateTable(ctl, table, { totalRow: !table.totalRow })} />
            <MenuItem label="Aralığa Dönüştür" onClick={() => ctl.openDialog("confirm", { message: "Tabloyu normal aralığa dönüştürmek istiyor musunuz?", onOk: () => D.convertTableToRange(ctl, table), okLabel: "Evet" })} />
          </Menu>
        }
      />
    </>
  );
}

export function CellMenu({ ctl }: { ctl: Controller }) {
  const { row, col } = ctl.sel.active;
  const sheet = ctl.sheet;
  const note = sheet.noteAt(row, col);
  const link = sheet.linkAt(row, col);
  const hasFilter = D.filtersOf(sheet).length > 0;
  const table = sheet.tableAt(row, col);
  return (
    <Menu width={230}>
      {/* Yaygın tablolama uygulamalarındaki gibi menünün üstünde mini araç çubuğu. */}
      <ContextMiniBar ctl={ctl} />
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={() => copy(ctl, true)} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={() => copy(ctl, false)} />
      <PasteItems ctl={ctl} />
      <MenuSeparator />
      {table ? (
        <TableItems ctl={ctl} table={table} />
      ) : (
        <>
          <MenuItem label="Ekle..." onClick={() => E.insertCellsSmart(ctl)} />
          <MenuItem label="Sil..." onClick={() => E.deleteCellsSmart(ctl)} />
        </>
      )}
      <MenuItem label="İçeriği Temizle" onClick={() => E.clearContents(ctl)} />
      <MenuSeparator />
      <MenuItem
        icon={<Filter20Regular />}
        label="Filtre"
        submenu={
          <Menu width={260}>
            <MenuItem label="Filtreyi Temizle" disabled={!hasFilter} onClick={() => D.clearAllFilters(ctl)} />
            <MenuItem label="Yeniden Uygula" disabled={!hasFilter} onClick={() => D.reapplyFilters(ctl)} />
            <MenuSeparator />
            <MenuItem
              label="Seçili Hücrenin Değerine Göre Filtre Uygula"
              onClick={() => {
                if (!sheet.autoFilter) D.toggleAutoFilter(ctl);
                const f = sheet.autoFilter;
                if (f) D.setColumnFilter(ctl, f, col, { kind: "values", values: [displayText(ctl.wb, sheet, row, col)], blanks: false });
              }}
            />
            <MenuItem
              label="Seçili Hücrenin Rengine Göre Filtre Uygula"
              onClick={() => {
                if (!sheet.autoFilter) D.toggleAutoFilter(ctl);
                const f = sheet.autoFilter;
                if (f) D.setColumnFilter(ctl, f, col, { kind: "color", color: ctl.wb.styleAt(sheet, row, col).fill ?? "" });
              }}
            />
          </Menu>
        }
      />
      <MenuItem
        label="Sırala"
        submenu={
          <Menu width={240}>
            <MenuItem icon={<ArrowSortDownLines20Regular />} label="A'dan Z'ye Sırala" onClick={() => D.quickSort(ctl, false)} />
            <MenuItem icon={<ArrowSortUpLines20Regular />} label="Z'den A'ya Sırala" onClick={() => D.quickSort(ctl, true)} />
            <MenuItem
              label="Seçili Hücre Rengini Üste Koy"
              onClick={() => {
                const r = D.sortRegion(ctl);
                const fill = ctl.wb.styleAt(sheet, row, col).fill ?? null;
                D.sortRange(ctl, r, [{ col, desc: false, by: "fill", order: fill ? [fill] : [] }], D.guessHeader(ctl.wb, sheet, r));
              }}
            />
            <MenuItem label="Özel Sıralama..." onClick={() => ctl.openDialog("sort")} />
          </Menu>
        }
      />
      <MenuSeparator />
      {note ? (
        <>
          <MenuItem icon={<NoteAdd20Regular />} label="Notu Düzenle" onClick={() => ctl.openDialog("note")} />
          <MenuItem label="Notu Sil" onClick={() => ctl.ops.setNote(sheet, row, col, null)} />
        </>
      ) : (
        <MenuItem icon={<NoteAdd20Regular />} label="Not Ekle" shortcut="Shift+F2" onClick={() => ctl.openDialog("note")} />
      )}
      <MenuSeparator />
      <MenuItem icon={<TableEdit20Regular />} label="Hücreleri Biçimlendir..." shortcut="Ctrl+1" onClick={() => ctl.openDialog("formatCells")} />
      <MenuItem label="Açılır Listeden Seç..." onClick={() => ctl.openDialog("pickList")} />
      <MenuItem label="Ad Tanımla..." onClick={() => ctl.openDialog("defineName")} />
      {link ? (
        <>
          <MenuItem icon={<Link20Regular />} label="Köprüyü Düzenle..." onClick={() => ctl.openDialog("hyperlink")} />
          <MenuItem label="Köprüyü Aç" onClick={() => openLink(ctl, link)} />
          <MenuItem label="Köprüyü Kaldır" onClick={() => ctl.ops.setLink(sheet, row, col, null)} />
        </>
      ) : (
        <MenuItem icon={<Link20Regular />} label="Bağlantı" shortcut="Ctrl+K" onClick={() => ctl.openDialog("hyperlink")} />
      )}
      {sheet.viewMode === "pageBreak" && <PageBreakItems ctl={ctl} />}
    </Menu>
  );
}

/** Sayfa Sonu Önizleme görünümünde sağ tık komutları. */
function PageBreakItems({ ctl }: { ctl: Controller }) {
  const p = ctl.sheet.page;
  const { row, col } = ctl.sel.active;
  const onBreak = (p.rowBreaks ?? []).includes(row) || (p.colBreaks ?? []).includes(col);
  const any = !!(p.rowBreaks?.length || p.colBreaks?.length);
  return (
    <>
      <MenuSeparator />
      <MenuItem label="Sayfa Sonu Ekle" onClick={() => insertPageBreak(ctl, false)} />
      {onBreak && <MenuItem label="Sayfa Sonunu Kaldır" onClick={() => insertPageBreak(ctl, true)} />}
      <MenuItem label="Tüm Sayfa Sonlarını Sıfırla" disabled={!any} onClick={() => resetPageBreaks(ctl)} />
      <MenuItem label="Yazdırma Alanını Belirle" onClick={() => setPage(ctl, { printArea: { ...ctl.range } }, "Yazdırma Alanı")} />
      <MenuItem label="Yazdırma Alanını Sıfırla" disabled={!p.printArea} onClick={() => setPage(ctl, { printArea: undefined }, "Yazdırma Alanı")} />
      <MenuItem label="Sayfa Yapısı..." onClick={() => ctl.openDialog("pageSetup")} />
    </>
  );
}

export function HeaderMenu({ ctl, axis }: { ctl: Controller; axis: "row" | "col" }) {
  const r = ctl.range;
  const items: number[] = [];
  const [a, b] = axis === "row" ? [r.r1, Math.min(r.r2, r.r1 + 100000)] : [r.c1, r.c2];
  for (let i = a; i <= b; i++) items.push(i);
  return (
    <Menu width={220}>
      <ContextMiniBar ctl={ctl} />
      <MenuItem icon={<Cut20Regular />} label="Kes" onClick={() => copy(ctl, true)} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" onClick={() => copy(ctl, false)} />
      <PasteItems ctl={ctl} />
      <MenuSeparator />
      <MenuItem label="Ekle" onClick={() => (axis === "row" ? E.insertRows(ctl) : E.insertCols(ctl))} />
      <MenuItem label="Sil" onClick={() => (axis === "row" ? E.deleteRows(ctl) : E.deleteCols(ctl))} />
      <MenuItem label="İçeriği Temizle" onClick={() => E.clearContents(ctl)} />
      <MenuSeparator />
      <MenuItem icon={<TableEdit20Regular />} label="Hücreleri Biçimlendir..." onClick={() => ctl.openDialog("formatCells")} />
      <MenuItem label={axis === "row" ? "Satır Yüksekliği..." : "Sütun Genişliği..."} onClick={() => ctl.openDialog(axis === "row" ? "rowHeight" : "colWidth")} />
      <MenuItem label="Gizle" onClick={() => ctl.ops.setHidden(ctl.sheet, axis, items, true)} />
      <MenuItem
        label="Göster"
        onClick={() => {
          const ax = axis === "row" ? ctl.sheet.rows : ctl.sheet.cols;
          ctl.ops.setHidden(ctl.sheet, axis, ax.hiddenItems().filter((i) => i >= a - 1 && i <= b + 1), false);
        }}
      />
    </Menu>
  );
}

export function TabMenu({ ctl, index }: { ctl: Controller; index: number }) {
  const sheet = ctl.wb.sheets[index];
  // Çalışma kitabı yapısı korunurken yapı komutları devre dışı
  const locked = !!ctl.wb.protection?.structure;
  return (
    <Menu width={220}>
      <MenuItem label="Ekle..." disabled={locked} onClick={() => ctl.ops.addSheet(index)} />
      <MenuItem icon={<Delete20Regular />} label="Sil" disabled={locked || ctl.wb.sheets.filter((s) => !s.hidden).length <= 1} onClick={() => requestDeleteSheet(ctl, index)} />
      <MenuItem label="Yeniden Adlandır" disabled={locked} onClick={() => window.dispatchEvent(new CustomEvent("tb-rename-sheet", { detail: index }))} />
      <MenuItem label="Taşı veya Kopyala..." disabled={locked} onClick={() => ctl.openDialog("moveSheet", { index })} />
      <MenuItem label={sheet.protected ? "Sayfa Korumasını Kaldır..." : "Sayfayı Koru..."} onClick={() => ctl.openDialog("protect")} />
      <MenuItem
        label="Sekme Rengi"
        submenu={
          <div className="oh-menu oh-menu-palette">
            <ColorPalette value={sheet.tabColor} noneLabel="Renk Yok" theme={paletteOf(ctl.wb.theme)} onSelect={(c) => ctl.ops.setTabColor(sheet, c ?? undefined)} />
          </div>
        }
      />
      <MenuSeparator />
      <MenuItem label="Gizle" disabled={locked || ctl.wb.sheets.filter((s) => !s.hidden).length <= 1} onClick={() => ctl.ops.setSheetHidden(index, true)} />
      <MenuItem label="Göster..." disabled={locked || !ctl.wb.sheets.some((s) => s.hidden)} onClick={() => ctl.openDialog("unhideSheet")} />
    </Menu>
  );
}

export function ObjectMenu({ ctl, kind, id }: { ctl: Controller; kind: "chart" | "image" | "shape"; id: string }) {
  const select = () => {
    ctl.selObj = { kind, id };
  };
  return (
    <Menu width={240}>
      <MenuItem icon={<Cut20Regular />} label="Kes" onClick={() => (select(), copyObject(ctl, true))} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" onClick={() => (select(), copyObject(ctl, false))} />
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" disabled={!hasObjectClip()} onClick={() => pasteObject(ctl)} />
      <MenuSeparator />
      {kind === "chart" && (
        <>
          <MenuItem label="Grafik Türünü Değiştir..." onClick={() => ctl.openDialog("chartType")} />
          <MenuItem label="Veri Seç..." onClick={() => ctl.openDialog("chartData")} />
          <MenuItem label="Grafiği Taşı..." onClick={() => ctl.openDialog("moveChart")} />
          <MenuSeparator />
          <MenuItem label="Grafik Alanını Biçimlendir..." onClick={() => (select(), ctl.openDialog("chartFormat"))} />
          <MenuSeparator />
        </>
      )}
      {kind === "shape" && <MenuItem label="Metni Düzenle" onClick={() => startShapeEdit(ctl, id, "")} />}
      <MenuItem label="Öne Getir" onClick={() => reorderObject(ctl, kind, id, true)} />
      <MenuItem label="Arkaya Gönder" onClick={() => reorderObject(ctl, kind, id, false)} />
      <MenuSeparator />
      <MenuItem icon={<Delete20Regular />} label="Sil" onClick={() => deleteObject(ctl, kind, id)} />
    </Menu>
  );
}
