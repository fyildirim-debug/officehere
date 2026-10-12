// Başlık çubuğu: Hızlı Erişim Araç Çubuğu (Kaydet, Geri Al, Yinele, Baskı Önizleme ve Yazdır) ve komut araması.
import { ArrowRedo20Regular, ArrowUndo20Regular, Print20Regular, Save20Regular } from "@fluentui/react-icons";
import { CommandSearch, QatButton, TitleBar, fileCommands, type OfficeBackstagePage, type RibbonTabDef } from "@officehere/ui";
import type { Editor } from "@tiptap/core";
import { repeatOrRedo } from "../editor/repeat";
import { useBelge } from "../state/context";

/** Komut aramasının şerit ve Dosya menüsü bağlantıları. */
export interface TitleBarSearch {
  tabs: RibbonTabDef[];
  collapsed: boolean;
  onExpand: () => void;
  /** Sekmeye geçer (Dosya görünümü açıksa kapanır). */
  onTab: (id: string) => void;
  onBackstage: (page: OfficeBackstagePage) => void;
  onNew: () => void;
  onOpen: () => void;
  onSaveAs: () => void;
  onPdf: () => void;
}

interface Props {
  title: string;
  dirty: boolean;
  active: Editor;
  canUndo: boolean;
  canRedo: boolean;
  onSave: () => void;
  onPrint: () => void;
  onFind: () => void;
  search: TitleBarSearch;
}

export function AppTitleBar({ title, dirty, active, canUndo, canRedo, onSave, onPrint, onFind, search }: Props) {
  const { notify } = useBelge();
  // Gezinti bölmesinde arar (Alt+Q kutusunun eski "Bul" davranışı).
  const findInDocument = (q: string) => {
    onFind();
    requestAnimationFrame(() => {
      const input = document.querySelector<HTMLInputElement>(".bl-nav-search input");
      if (input) {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
        setter?.call(input, q);
        input.dispatchEvent(new Event("input", { bubbles: true }));
      }
    });
  };
  const files = fileCommands(
    { newDoc: search.onNew, open: search.onOpen, save: onSave, saveAs: search.onSaveAs, exportPdf: search.onPdf, print: onPrint, backstage: search.onBackstage },
    "Belge",
  );
  return (
    <>
      <TitleBar
        app="belge"
        title={title}
        dirty={dirty}
        quickAccess={
          <>
            <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={onSave} />
            <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" disabled={!canUndo} onClick={() => active.chain().focus().undo().run()} />
            <QatButton
              icon={<ArrowRedo20Regular />}
              tip="Yinele"
              shortcut="Ctrl+Y"
              disabled={!canRedo}
              onClick={() => {
                active.commands.focus();
                repeatOrRedo(active);
              }}
            />
            <QatButton icon={<Print20Regular />} tip="Baskı Önizleme ve Yazdır" shortcut="Ctrl+P" onClick={onPrint} />
          </>
        }
      />
      {/* Ctrl+K "Bağlantı ekle" olduğundan arama yalnız Alt+Q ile açılır. */}
      <CommandSearch
        app="belge"
        tabs={search.tabs}
        collapsed={search.collapsed}
        onExpand={search.onExpand}
        onTabChange={search.onTab}
        extra={files}
        find={{ label: (q) => `Belgede bul: ${q}`, path: "Gezinti bölmesi", shortcut: "Ctrl+F", run: findInDocument }}
        findOnEnter
        onMessage={(t) => notify(t, "info")}
        onDismiss={() => active.commands.focus()}
      />
    </>
  );
}
