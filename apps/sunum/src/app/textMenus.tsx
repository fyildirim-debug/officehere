/**
 * Metin düzenlerken ve tablo hücresinde sağ tık menüleri: Kes/Kopyala,
 * Yapıştırma Seçenekleri, Yazı Tipi, Paragraf, Madde İşaretleri,
 * Numaralandırma, Köprü; tabloda ayrıca Ekle / Sil / Seç / Birleştir / Böl.
 */
import {
  ClipboardPaste20Regular,
  ClipboardTextLtr20Regular,
  CommentAdd20Regular,
  Copy20Regular,
  Cut20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  LinkEdit20Regular,
  Open20Regular,
  PaintBucket20Regular,
  TableCellsMerge20Regular,
  TableCellsSplit20Regular,
  TableDeleteRow20Regular,
  TableInsertRow20Regular,
  TextBulletListLtr20Regular,
  TextFont20Regular,
  TextNumberListLtr20Regular,
  TextParagraph20Regular,
} from "@fluentui/react-icons";
import { getUserName, Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { deleteCols, deleteRows, insertCol, insertRow, mergeCells, normRange, splitCell, type CellRange } from "../model/tableOps";
import type { TableShape } from "../model/types";
import { BulletGrid } from "../ribbon/HomeTab";
import type { AppCommands, DialogKind } from "../ribbon/types";
import type { Actions } from "../state/actions";
import { currentLink } from "../state/linkActions";
import { setListStyle, textInfo } from "../state/textActions";

/** Menüden çağrılan pano komutu: önce düzenleyiciye odak ve seçim geri verilir. */
function clip(a: Actions, op: "cut" | "copy") {
  a.textHandle?.focus();
  document.execCommand(op);
}

function pasteItems(a: Actions, cmd: AppCommands) {
  return (
    <MenuItem
      icon={<ClipboardPaste20Regular />}
      label="Yapıştırma Seçenekleri:"
      submenu={
        <Menu>
          <MenuItem icon={<ClipboardPaste20Regular />} label="Kaynak Biçimlendirmesini Koru" shortcut="Ctrl+V" onClick={() => (a.textHandle?.focus(), cmd.paste("keep"))} />
          <MenuItem icon={<ClipboardTextLtr20Regular />} label="Yalnızca Metni Koru" onClick={() => (a.textHandle?.focus(), cmd.paste("text"))} />
        </Menu>
      }
    />
  );
}

function linkItems(a: Actions, cmd: AppCommands) {
  const link = currentLink(a);
  return link ? (
    <>
      <MenuItem icon={<LinkEdit20Regular />} label="Bağlantıyı Düzenle..." onClick={cmd.insertLink} />
      <MenuItem icon={<Open20Regular />} label="Bağlantıyı Aç" onClick={cmd.openLink} />
      <MenuItem icon={<LinkDismiss20Regular />} label="Bağlantıyı Kaldır" onClick={cmd.removeLink} />
    </>
  ) : (
    <MenuItem icon={<Link20Regular />} label="Bağlantı..." shortcut="Ctrl+K" onClick={cmd.insertLink} />
  );
}

function formatItems(a: Actions, openDialog: (d: DialogKind) => void) {
  const info = textInfo(a);
  const marL = info?.para?.marL;
  return (
    <>
      <MenuItem icon={<TextFont20Regular />} label="Yazı Tipi..." shortcut="Ctrl+Shift+F" onClick={() => openDialog("font")} />
      <MenuItem icon={<TextParagraph20Regular />} label="Paragraf..." onClick={() => openDialog("paragraph")} />
      <MenuItem icon={<TextBulletListLtr20Regular />} label="Madde İşaretleri" submenu={<BulletGrid kind="bullet" onPick={(b) => setListStyle(a, b, !marL)} />} />
      <MenuItem icon={<TextNumberListLtr20Regular />} label="Numaralandırma" submenu={<BulletGrid kind="number" onPick={(b) => setListStyle(a, b)} />} />
    </>
  );
}

/** Şekil/metin kutusunda metin düzenlenirken. */
export function textEditMenu(a: Actions, cmd: AppCommands, openDialog: (d: DialogKind) => void) {
  const hasSel = !!a.textHandle && a.textHandle.selectedText().length > 0;
  return (
    <Menu>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" disabled={!hasSel} onClick={() => clip(a, "cut")} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" disabled={!hasSel} onClick={() => clip(a, "copy")} />
      {pasteItems(a, cmd)}
      <MenuSeparator />
      {formatItems(a, openDialog)}
      <MenuSeparator />
      {linkItems(a, cmd)}
      <MenuSeparator />
      <MenuItem icon={<PaintBucket20Regular />} label="Şekli Biçimlendir..." onClick={() => cmd.togglePane("shape")} />
      <MenuItem icon={<CommentAdd20Regular />} label="Yeni Açıklama" onClick={() => a.addComment(getUserName())} />
    </Menu>
  );
}

function tableRange(a: Actions, t: TableShape): CellRange {
  const st = a.s;
  if (st.cellSel) return normRange(st.cellSel);
  const c = st.editing?.shapeId === t.id ? st.editing.cell : undefined;
  if (c) return { r1: c[0], c1: c[1], r2: c[0], c2: c[1] };
  return { r1: 0, c1: 0, r2: 0, c2: 0 };
}

function updateTable(a: Actions, t: TableShape, fn: (t: TableShape) => TableShape | null) {
  const res = fn(t);
  if (!res) {
    a.store.set({ editing: null, cellSel: null });
    a.deleteSelection();
    return;
  }
  a.updateShapes((s) => (s.id === t.id ? res : s), [t.id]);
}

/** Tablo hücresinde (düzenlerken ya da hücre seçiliyken). */
export function tableCellMenu(a: Actions, cmd: AppCommands, openDialog: (d: DialogKind) => void, t: TableShape) {
  const r = tableRange(a, t);
  const editing = !!a.s.editing?.cell;
  const hasSel = editing && !!a.textHandle && a.textHandle.selectedText().length > 0;
  const cell = t.rows[r.r1]?.cells[r.c1];
  const multi = r.r1 !== r.r2 || r.c1 !== r.c2;
  const leave = () => a.store.set({ editing: null, cellSel: null });
  return (
    <Menu>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" disabled={editing && !hasSel} onClick={() => (editing ? clip(a, "cut") : cmd.cut())} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" disabled={editing && !hasSel} onClick={() => (editing ? clip(a, "copy") : cmd.copy())} />
      {editing ? pasteItems(a, cmd) : <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => cmd.paste()} />}
      <MenuSeparator />
      <MenuItem
        icon={<TableInsertRow20Regular />}
        label="Ekle"
        submenu={
          <Menu>
            <MenuItem label="Sütunları Sola Ekle" onClick={() => updateTable(a, t, (x) => insertCol(x, r.c1, false))} />
            <MenuItem label="Sütunları Sağa Ekle" onClick={() => updateTable(a, t, (x) => insertCol(x, r.c2, true))} />
            <MenuItem label="Satırları Üste Ekle" onClick={() => updateTable(a, t, (x) => insertRow(x, r.r1, false))} />
            <MenuItem label="Satırları Alta Ekle" onClick={() => updateTable(a, t, (x) => insertRow(x, r.r2, true))} />
          </Menu>
        }
      />
      <MenuItem
        icon={<TableDeleteRow20Regular />}
        label="Sil"
        submenu={
          <Menu>
            <MenuItem label="Sütunları Sil" onClick={() => (leave(), updateTable(a, t, (x) => deleteCols(x, r.c1, r.c2)))} />
            <MenuItem label="Satırları Sil" onClick={() => (leave(), updateTable(a, t, (x) => deleteRows(x, r.r1, r.r2)))} />
            <MenuItem label="Tabloyu Sil" onClick={() => (leave(), a.deleteSelection())} />
          </Menu>
        }
      />
      <MenuItem
        label="Seç"
        submenu={
          <Menu>
            <MenuItem label="Tabloyu Seç" onClick={() => a.store.set({ cellSel: { r1: 0, c1: 0, r2: t.rows.length - 1, c2: t.cols.length - 1 }, editing: null })} />
            <MenuItem label="Sütunu Seç" onClick={() => a.store.set({ cellSel: { r1: 0, c1: r.c1, r2: t.rows.length - 1, c2: r.c2 }, editing: null })} />
            <MenuItem label="Satırı Seç" onClick={() => a.store.set({ cellSel: { r1: r.r1, c1: 0, r2: r.r2, c2: t.cols.length - 1 }, editing: null })} />
          </Menu>
        }
      />
      <MenuItem icon={<TableCellsMerge20Regular />} label="Hücreleri Birleştir" disabled={!multi} onClick={() => (leave(), updateTable(a, t, (x) => mergeCells(x, r)))} />
      <MenuItem icon={<TableCellsSplit20Regular />} label="Hücreleri Böl" disabled={!(cell?.gridSpan || cell?.rowSpan)} onClick={() => updateTable(a, t, (x) => splitCell(x, r.r1, r.c1))} />
      <MenuSeparator />
      {formatItems(a, openDialog)}
      {editing && (
        <>
          <MenuSeparator />
          {linkItems(a, cmd)}
        </>
      )}
      <MenuSeparator />
      <MenuItem icon={<PaintBucket20Regular />} label="Şekli Biçimlendir..." onClick={() => cmd.togglePane("shape")} />
      <MenuItem icon={<CommentAdd20Regular />} label="Yeni Açıklama" onClick={() => a.addComment(getUserName())} />
    </Menu>
  );
}
