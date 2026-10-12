/** OfficeHere Tablo uygulama kökü: kabuk, şerit, formül çubuğu, ızgara, durum çubuğu. */
import { useEffect, useMemo, useState, useSyncExternalStore } from "react";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular } from "@fluentui/react-icons";
import { AppShell, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, useHotkeys, type RibbonTabDef } from "@officehere/ui";
import { baseName } from "@officehere/core";
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
import { installTestHook } from "./testHook";
import { pivotAt, pivotButtons } from "./pivot/pivotOps";
import { PivotAnalyzeTab, PivotDesignTab } from "./pivot/PivotTabs";
import { PivotButtonMenu, PivotFieldsPane } from "./pivot/PivotFieldsPane";
import { PivotCellMenu } from "./pivot/PivotMenu";
import { WatchWindow } from "./panes/WatchWindow";
import { ThemePaletteContext } from "./themeContext";
import { paletteOf } from "./model/theme";

let dialogSeq = 0;

export function App() {
  const ctl = useMemo(() => new Controller(), []);
  useEffect(() => installTestHook(ctl), [ctl]);
  // Geliştirme kipinde tarayıcı testleri için erişim
  if (import.meta.env.DEV) (window as unknown as { __tablo?: Controller }).__tablo = ctl;
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const [tab, setTab] = useState("home");
  const [collapsed, setCollapsed] = useState(false);
  const [dialogs, setDialogs] = useState<DialogState[]>([]);
  // İletişim kutuları yığın olarak tutulur: uyarı/onay, açık iletişim kutusunun üstünde açılır
  const pushDialog = (d: DialogState) =>
    setDialogs((st) => {
      // Modsuz Bul ve Değiştir zaten açıksa yenisi açılmaz, sekmesi değişir
      const isFind = (x: string) => x === "find" || x === "replace";
      if (isFind(d.id) && st.some((x) => isFind(x.id))) return st.map((x) => (isFind(x.id) ? { ...d, key: ++dialogSeq } : x));
      return [...st, { ...d, key: ++dialogSeq }];
    });
  const [backstage, setBackstage] = useState<string | null>(null);
  const [filterMenu, setFilterMenu] = useState<{ row: number; col: number; x: number; y: number } | null>(null);
  const ctx = useContextMenu();
  const files = useFileActions(ctl);

  const lifecycle = useDocumentLifecycle({
    title: baseName(ctl.fileName),
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
      // İletişim kutuları hücre içeriğiyle çalışır: süren düzenleme önce onaylanır
      if (ctl.edit && id !== "insertFunction" && id !== "alert" && id !== "confirm") ctl.commitEdit(null);
      pushDialog({ id, arg });
    };
    ctl.alert = (message, title) => pushDialog({ id: "alert", arg: { message, title } });
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
    "Mod+Alt+V": () => !ctl.edit && pushDialog({ id: "pasteSpecial" }),
    "Mod+F1": () => setCollapsed((c) => !c),
  });

  // Bağlamsal sekmeler
  const sheet = ctl.sheet;
  const activeTable = sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col);
  const activeChart = ctl.selObj?.kind === "chart" ? sheet.charts.find((c) => c.id === ctl.selObj!.id) : undefined;
  const activePivot = !ctl.selObj ? pivotAt(sheet, ctl.sel.active.row, ctl.sel.active.col) : undefined;
  useEffect(() => {
    if ((tab === "tableDesign" && !activeTable) || ((tab === "chartDesign" || tab === "chartFormat") && !activeChart) || ((tab === "pivotAnalyze" || tab === "pivotDesign") && !activePivot)) setTab("home");
  }, [tab, activeTable, activeChart, activePivot]);
  useEffect(() => {
    if (activeChart && tab !== "chartFormat") setTab("chartDesign");
  }, [activeChart?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (ctl.requestTab) {
      setTab(ctl.requestTab);
      ctl.requestTab = null;
    }
  });

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
    { id: "pivotAnalyze", label: "PivotTable Analiz", contextual: true, contextualGroup: "PivotTable Araçları", hidden: !activePivot, content: activePivot ? <PivotAnalyzeTab ctl={ctl} def={activePivot} /> : null },
    { id: "pivotDesign", label: "Tasarım", contextual: true, contextualGroup: "PivotTable Araçları", hidden: !activePivot, content: activePivot ? <PivotDesignTab ctl={ctl} def={activePivot} /> : null },
  ];

  const palette = useMemo(() => paletteOf(ctl.wb.theme), [ctl.wb.theme]);

  const openContext = (e: { clientX: number; clientY: number; preventDefault: () => void }, kind: "cell" | "col" | "row" | "tab", index?: number) => {
    if (ctl.edit) return;
    if (kind === "tab") ctx.open(e, <TabMenu ctl={ctl} index={index ?? ctl.wb.activeIndex} />);
    else if (kind === "cell") {
      const pv = pivotAt(ctl.sheet, ctl.sel.active.row, ctl.sel.active.col);
      ctx.open(e, pv && pv.output && pv.rows.length + pv.cols.length + pv.values.length > 0 ? <PivotCellMenu ctl={ctl} def={pv} /> : <CellMenu ctl={ctl} />);
    }
    else ctx.open(e, <HeaderMenu ctl={ctl} axis={kind === "col" ? "col" : "row"} />);
  };

  return (
    <ThemePaletteContext.Provider value={palette}>
    <AppShell
      app="tablo"
      titleBar={
        <TitleBar
          app="tablo"
          title={baseName(ctl.fileName)}
          dirty={ctl.dirty}
          onSearch={(q) => {
            if (q.trim()) pushDialog({ id: "find", arg: { query: q } });
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
        <div className="tb-workarea">
          {sheet.viewMode === "pageLayout" ? (
            <PageLayoutView ctl={ctl} />
          ) : (
            <Grid ctl={ctl} onContextMenu={openContext} onFilterButton={(row, col, x, y) => setFilterMenu({ row, col, x, y })} />
          )}
          {activePivot && ctl.showFieldList && <PivotFieldsPane ctl={ctl} sheet={sheet} def={activePivot} />}
        </div>
        {ctl.showWatch && <WatchWindow ctl={ctl} />}
      </div>
      {filterMenu &&
        (() => {
          const close = () => (setFilterMenu(null), ctl.focusGrid());
          const pb = pivotButtons(sheet).find((b) => b.row === filterMenu.row && b.col === filterMenu.col);
          if (pb) return <PivotButtonMenu ctl={ctl} sheet={sheet} pivotId={pb.pivotId} field={pb.field} x={filterMenu.x} y={filterMenu.y} onClose={close} />;
          return <FilterMenu ctl={ctl} {...filterMenu} onClose={close} />;
        })()}
      {ctx.element}
      <DialogHost
        ctl={ctl}
        stack={dialogs}
        onClose={(key) => {
          setDialogs((st) => {
            const next = st.filter((d) => d.key !== key);
            if (!next.length) requestAnimationFrame(() => ctl.focusGrid());
            return next;
          });
        }}
        files={files}
        onBackstage={setBackstage}
      />
      {lifecycle.dialog}
    </AppShell>
    </ThemePaletteContext.Provider>
  );
}
