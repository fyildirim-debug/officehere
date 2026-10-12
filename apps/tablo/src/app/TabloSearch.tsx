/** Başlık çubuğundaki komut araması (Alt+Q): şerit komutları, Dosya eylemleri ve "Sayfada bul". */
import { CommandSearch, fileCommands, type RibbonTabDef } from "@officehere/ui";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";

interface Props {
  ctl: Controller;
  files: FileActions;
  tabs: RibbonTabDef[];
  collapsed: boolean;
  onExpand: () => void;
  /** Sekmeye geçer (Dosya görünümü açıksa kapanır). */
  onTab: (id: string) => void;
  onBackstage: (page: string) => void;
}

export function TabloSearch({ ctl, files, tabs, collapsed, onExpand, onTab, onBackstage }: Props) {
  // Ctrl+O ve Ctrl+P gibi Aç ve Yazdır da Dosya görünümünün ilgili sayfasını açar.
  const extra = fileCommands(
    {
      newDoc: () => void files.newBook("blank"),
      open: () => onBackstage("open"),
      save: () => void files.save(),
      saveAs: () => void files.saveAs(),
      exportPdf: () => void files.exportAs("pdf"),
      print: () => onBackstage("print"),
      backstage: onBackstage,
    },
    "Çalışma Kitabı",
  );
  // Ctrl+K "Bağlantı" olduğundan arama yalnız Alt+Q ile açılır.
  return (
    <CommandSearch
      app="tablo"
      tabs={tabs}
      collapsed={collapsed}
      onExpand={onExpand}
      onTabChange={onTab}
      extra={extra}
      find={{
        label: (q) => `Sayfada bul: ${q}`,
        path: "Bul ve Değiştir",
        shortcut: "Ctrl+F",
        run: (q) => {
          if (q) ctl.openDialog("find", { query: q });
        },
      }}
      findOnEnter
      onMessage={(text) => {
        ctl.infoBar = { kind: "info", text };
        ctl.emit();
      }}
      onDismiss={() => ctl.focusGrid()}
    />
  );
}
