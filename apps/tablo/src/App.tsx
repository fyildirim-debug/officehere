/** OfficeHere Tablo uygulama kökü: kabuk, şerit, formül çubuğu, ızgara, durum çubuğu. */
import { useEffect, useMemo, useState, useSyncExternalStore } from "react";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular } from "@fluentui/react-icons";
import { AppShell, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, useHotkeys, type RibbonTabDef } from "@officehere/ui";
import { Controller } from "./controller";
import { Grid } from "./grid/Grid";
import { FormulaBar } from "./grid/FormulaBar";
import { HomeTab } from "./ribbon/HomeTab";
import { StatusArea } from "./StatusArea";
import * as E from "./actions/edit";
import { CellMenu, HeaderMenu, TabMenu, ObjectMenu } from "./ContextMenus";
import { DialogHost, type DialogState } from "./dialogs/DialogHost";
import { InsertTab } from "./ribbon/InsertTab";
import { PageLayoutTab } from "./ribbon/PageLayoutTab";
import { FormulasTab } from "./ribbon/FormulasTab";
import { DataTab } from "./ribbon/DataTab";
import { ReviewTab } from "./ribbon/ReviewTab";
import { ViewTab } from "./ribbon/ViewTab";
import { HelpTab } from "./ribbon/HelpTab";
import { TableDesignTab } from "./ribbon/TableDesignTab";
import { ChartDesignTab, ChartFormatTab } from "./ribbon/ChartTabs";
import { FileView, useFileActions } from "./FileView";
import { FilterMenu } from "./grid/FilterMenu";
import { PageLayoutView } from "./grid/PageLayoutView";

export function App() {
  const ctl = useMemo(() => new Controller(), []);
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const [tab, setTab] = useState("home");
  const [collapsed, setCollapsed] = useState(false);
  const [dialog, setDialog] = useState<DialogState | null>(null);
  const [backstage, setBackstage] = useState<string | null>(null);
  const [filterMenu, setFilterMenu] = useState<{ row: number; col: number; x: number; y: number } | null>(null);
  const ctx = useContextMenu();
  const files = useFileActions(ctl);

  const lifecycle = useDocumentLifecycle({
    title: ctl.fileName,
    dirty: ctl.dirty,
    onOpenPath: (p) => void files.openPath(p),
    onSave: () => files.save(),
  });
  files.confirmDiscard = lifecycle.confirmDiscard;

  useEffect(() => {
    ctl.openDialog = (id, arg) => {
      if (id === "objectMenu") {
        const a = arg as { x: number; y: number; kind: "chart" | "image" | "shape"; id: string };
        ctx.open({ clientX: a.x, clientY: a.y, preventDefault: () => {} }, <ObjectMenu ctl={ctl} kind={a.kind} id={a.id} />);
        return;
      }
      setDialog({ id, arg });
    };
    ctl.alert = (message, title) => setDialog({ id: "alert", arg: { message, title } });
  });

  // Pano olayları (Ctrl+C / X / V)
  useEffect(() => {
    const inGrid = () => {
      const a = document.activeElement;
      return !!a && a.classList.contains("tb-editor-input") && !ctl.edit && !backstage;
    };
    const onCopy = (e: ClipboardEvent) => {
      if (!inGrid()) return;
      const d = E.copySelection(ctl, e.type === "cut");
      if (!d) return;
      e.preventDefault();
      e.clipboardData?.setData("text/plain", d.tsv);
      e.clipboardData?.setData("text/html", d.html);
    };
    const onPaste = (e: ClipboardEvent) => {
      if (!inGrid()) return;
      e.preventDefault();
      E.pasteData(ctl, e.clipboardData?.getData("text/plain") ?? "", e.clipboardData?.getData("text/html") || undefined);
    };
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCopy);
    document.addEventListener("paste", onPaste);
    return () => {
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("cut", onCopy);
      document.removeEventListener("paste", onPaste);
    };
  }, [ctl, backstage]);

  useHotkeys({
    "Mod+S": () => void files.save(),
    F12: () => void files.saveAs(),
    "Mod+O": () => setBackstage("open"),
    "Mod+N": () => void files.newBook("blank"),
    "Mod+P": () => setBackstage("print"),
    "Mod+Alt+V": () => !ctl.edit && setDialog({ id: "pasteSpecial" }),
    "Mod+F1": () => setCollapsed((c) => !c),
  });

  // Bağlamsal sekmeler
  const sheet = ctl.sheet;
  const activeTable = sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col);
  const activeChart = ctl.selObj?.kind === "chart" ? sheet.charts.find((c) => c.id === ctl.selObj!.id) : undefined;
  useEffect(() => {
    if ((tab === "tableDesign" && !activeTable) || ((tab === "chartDesign" || tab === "chartFormat") && !activeChart)) setTab("home");
  }, [tab, activeTable, activeChart]);
  useEffect(() => {
    if (activeChart && tab !== "chartFormat") setTab("chartDesign");
  }, [activeChart?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const tabs: RibbonTabDef[] = [
    { id: "home", label: "Giriş", content: <HomeTab ctl={ctl} /> },
    { id: "insert", label: "Ekle", content: <InsertTab ctl={ctl} /> },
    { id: "pageLayout", label: "Sayfa Düzeni", content: <PageLayoutTab ctl={ctl} /> },
    { id: "formulas", label: "Formüller", content: <FormulasTab ctl={ctl} /> },
    { id: "data", label: "Veri", content: <DataTab ctl={ctl} /> },
    { id: "review", label: "Gözden Geçir", content: <ReviewTab ctl={ctl} /> },
    { id: "view", label: "Görünüm", content: <ViewTab ctl={ctl} /> },
    { id: "help", label: "Yardım", content: <HelpTab ctl={ctl} /> },
    { id: "tableDesign", label: "Tablo Tasarımı", contextual: true, contextualGroup: "Tablo Araçları", hidden: !activeTable, content: activeTable ? <TableDesignTab ctl={ctl} table={activeTable} /> : null },
    { id: "chartDesign", label: "Grafik Tasarımı", contextual: true, contextualGroup: "Grafik Araçları", hidden: !activeChart, content: activeChart ? <ChartDesignTab ctl={ctl} chart={activeChart} /> : null },
    { id: "chartFormat", label: "Biçim", contextual: true, contextualGroup: "Grafik Araçları", hidden: !activeChart, content: activeChart ? <ChartFormatTab ctl={ctl} chart={activeChart} /> : null },
  ];

  const openContext = (e: { clientX: number; clientY: number; preventDefault: () => void }, kind: "cell" | "col" | "row" | "tab", index?: number) => {
    if (ctl.edit) return;
    if (kind === "tab") ctx.open(e, <TabMenu ctl={ctl} index={index ?? ctl.wb.activeIndex} />);
    else if (kind === "cell") ctx.open(e, <CellMenu ctl={ctl} />);
    else ctx.open(e, <HeaderMenu ctl={ctl} axis={kind === "col" ? "col" : "row"} />);
  };

  return (
    <AppShell
      app="tablo"
      titleBar={
        <TitleBar
          app="tablo"
          title={ctl.fileName}
          dirty={ctl.dirty}
          onSearch={(q) => {
            if (q.trim()) setDialog({ id: "find", arg: { query: q } });
          }}
          searchPlaceholder="Ara (Alt+Q)"
          quickAccess={
            <>
              <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={() => void files.save()} />
              <QatButton icon={<ArrowUndo20Regular />} tip={ctl.ops.history.undoLabel ? `Geri Al: ${ctl.ops.history.undoLabel}` : "Geri Al"} shortcut="Ctrl+Z" disabled={!ctl.ops.history.canUndo && !ctl.edit} onClick={() => E.undo(ctl)} />
              <QatButton icon={<ArrowRedo20Regular />} tip={ctl.ops.history.redoLabel ? `Yinele: ${ctl.ops.history.redoLabel}` : "Yinele"} shortcut="Ctrl+Y" disabled={!ctl.ops.history.canRedo} onClick={() => E.redo(ctl)} />
            </>
          }
        />
      }
      ribbon={<Ribbon tabs={tabs} activeTab={tab} onTabChange={setTab} onFileClick={() => setBackstage("home")} collapsed={collapsed} onCollapsedChange={setCollapsed} />}
      belowRibbon={ctl.showFormulaBar ? <FormulaBar ctl={ctl} /> : undefined}
      statusBar={<StatusArea ctl={ctl} />}
      backstage={backstage ? <FileView ctl={ctl} files={files} page={backstage} onClose={() => (setBackstage(null), requestAnimationFrame(() => ctl.focusGrid()))} /> : undefined}
    >
      <div className="tb-root">
        {sheet.viewMode === "pageLayout" ? (
          <PageLayoutView ctl={ctl} />
        ) : (
          <Grid ctl={ctl} onContextMenu={openContext} onFilterButton={(row, col, x, y) => setFilterMenu({ row, col, x, y })} />
        )}
      </div>
      {filterMenu && <FilterMenu ctl={ctl} {...filterMenu} onClose={() => (setFilterMenu(null), ctl.focusGrid())} />}
      {ctx.element}
      <DialogHost ctl={ctl} state={dialog} onClose={() => (setDialog(null), requestAnimationFrame(() => ctl.focusGrid()))} files={files} />
      {lifecycle.dialog}
    </AppShell>
  );
}
