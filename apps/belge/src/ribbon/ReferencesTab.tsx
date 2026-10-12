// Başvurular sekmesi: İçindekiler.
import { ArrowSync20Regular, TextBulletListSquare32Regular, TextAdd20Regular, Delete20Regular, DocumentEndnote20Regular, ArrowNext20Regular } from "@fluentui/react-icons";
import { TextSelection } from "@tiptap/pm/state";
import { collectNotes, type NoteKind } from "../editor/notes";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyStyle } from "../editor/commands";

export function ReferencesTab() {
  const ctx = useBelge();
  const { editor, styles } = ctx;
  const hasToc = (() => {
    let found = false;
    editor.state.doc.descendants((n) => {
      if (n.type.name === "toc") found = true;
      return !found;
    });
    return found;
  })();

  const insert = (title: string) => {
    editor.chain().focus().insertContent({ type: "toc", attrs: { entries: [], title } }).run();
    // Sayfa numaraları sayfalamadan sonra hazır olur.
    requestAnimationFrame(() => requestAnimationFrame(ctx.updateToc));
  };

  const removeToc = () => {
    const { state } = editor;
    const tr = state.tr;
    const ranges: [number, number][] = [];
    state.doc.descendants((n, pos) => {
      if (n.type.name === "toc") ranges.push([pos, pos + n.nodeSize]);
      return true;
    });
    for (const [a, b] of ranges.reverse()) tr.delete(a, b);
    editor.view.dispatch(tr);
  };

  const notes = collectNotes(editor.state.doc);
  const gotoNote = (kind: NoteKind, dir: 1 | -1) => {
    const list = notes.filter((n) => n.kind === kind);
    if (!list.length) return;
    const pos = editor.state.selection.from;
    const n = dir > 0 ? (list.find((x) => x.pos >= pos) ?? list[0]) : ([...list].reverse().find((x) => x.pos < pos - 1) ?? list[list.length - 1]);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, n.pos + 1)).scrollIntoView());
    editor.commands.focus();
  };

  return (
    <>
      <RibbonGroup label="İçindekiler" icon={<TextBulletListSquare32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<TextBulletListSquare32Regular />}
          label="İçindekiler"
          disabled={!ctx.inBody}
          menu={
            <Menu width={260}>
              <MenuHeader>Yerleşik</MenuHeader>
              <MenuItem label="Otomatik Tablo 1" description="Başlık: İçindekiler" onClick={() => insert("İçindekiler")} />
              <MenuItem label="Otomatik Tablo 2" description="Başlık: İçerik" onClick={() => insert("İçerik")} />
              <MenuSeparator />
              <MenuItem icon={<Delete20Regular />} label="İçindekiler Tablosunu Kaldır" disabled={!hasToc} onClick={removeToc} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonMenuButton
            icon={<TextAdd20Regular />}
            label="Metin Ekle"
            menu={
              <Menu>
                <MenuItem label="İçindekiler Tablosunda Gösterme" checked={!ctx.fmt.styleId.startsWith("Heading")} onClick={() => applyStyle(editor, styles, "Normal")} />
                <MenuItem label="Düzey 1" checked={ctx.fmt.styleId === "Heading1"} onClick={() => applyStyle(editor, styles, "Heading1")} />
                <MenuItem label="Düzey 2" checked={ctx.fmt.styleId === "Heading2"} onClick={() => applyStyle(editor, styles, "Heading2")} />
                <MenuItem label="Düzey 3" checked={ctx.fmt.styleId === "Heading3"} onClick={() => applyStyle(editor, styles, "Heading3")} />
              </Menu>
            }
          />
          <RibbonButton icon={<ArrowSync20Regular />} label="Tabloyu Güncelleştir" disabled={!hasToc} onClick={ctx.updateToc} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Dipnotlar" icon={<DocumentEndnote20Regular />}>
        <RibbonButton size="large" icon={<span className="bl-fn-icon">ab<sup>1</sup></span>} label="Dipnot Ekle" shortcut="Ctrl+Alt+F" onClick={() => ctx.insertNote("footnote")} />
        <RibbonStack>
          <RibbonButton icon={<DocumentEndnote20Regular />} label="Son Not Ekle" shortcut="Ctrl+Alt+D" onClick={() => ctx.insertNote("endnote")} />
          <RibbonMenuButton
            icon={<ArrowNext20Regular />}
            label="Sonraki Dipnot"
            disabled={!notes.length}
            menu={
              <Menu>
                <MenuItem label="Sonraki Dipnot" onClick={() => gotoNote("footnote", 1)} />
                <MenuItem label="Önceki Dipnot" onClick={() => gotoNote("footnote", -1)} />
                <MenuItem label="Sonraki Son Not" onClick={() => gotoNote("endnote", 1)} />
                <MenuItem label="Önceki Son Not" onClick={() => gotoNote("endnote", -1)} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
