/** OfficeHere Tablo uygulama kökü: kabuk, şerit, formül çubuğu, ızgara, durum çubuğu. */
import { useEffect, useMemo, useState, useSyncExternalStore } from "react";
import { ArrowRedo20Regular, ArrowUndo20Regular, Save20Regular } from "@fluentui/react-icons";
import { AppShell, MessageBar, QatButton, Ribbon, TitleBar, useContextMenu, useDocumentLifecycle, useHotkeys, type RibbonTabDef } from "@officehere/ui";
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
import { sparkAt } from "./sparkline/geometry";
import { SparklineTab } from "./sparkline/SparkTab";
import { ThemePaletteContext } from "./themeContext";
import { ChartFormatPane } from "./charts/ChartFormatPane";
import { SelectionPane } from "./charts/SelectionPane";
import { ErrorBoundary } from "./app/ErrorBoundary";
import { recoveryHandlers } from "./app/recovery";
import { useAutoRecover } from "./app/useAutoRecover";
import { copyObject, objectClipMarker, pasteObject } from "./actions/objectClip";
import { paletteOf } from "./model/theme";

let dialogSeq = 0;

export function App({ ctl: given }: { ctl?: Controller } = {}) {
  const ctl = useMemo(() => given ?? new Controller(), [given]);
  const recovery = useMemo(() => recoveryHandlers(ctl), [ctl]);
  const autoRecover = useAutoRecover(ctl);
  useEffect(() => installTestHook(ctl), [ctl]);
  // Şeritteki onay kutusu değişince odak çalışma sayfasına döner (Ctrl+Z vb. çalışsın)
  useEffect(() => {
    const onChange = (e: Event) => {
      const t = e.target;
      if (t instanceof HTMLInputElement && t.type === "checkbox" && t.closest(".oh-ribbon")) requestAnimationFrame(() => ctl.focusGrid());
    };
    document.addEventListener("change", onChange, true);
    // Odak şeritte/sekme düğmesindeyken de Ctrl+Z / Ctrl+Y çalışır
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented || e.isComposing || !(e.ctrlKey || e.metaKey) || e.altKey || (e.code !== "KeyZ" && e.code !== "KeyY")) return;
      const t = e.target as HTMLElement | null;
      if (t?.closest?.("textarea, [contenteditable='true'], .oh-dialog, .oh-backstage")) return;
      if (t instanceof HTMLInputElement && t.type !== "checkbox" && t.type !== "radio") return;
      if (ctl.edit) return;
      e.preventDefault();
      if (e.code === "KeyZ" && !e.shiftKey) E.undo(ctl);
      else E.redo(ctl);
      ctl.focusGrid();
    };
    window.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("change", onChange, true);
      window.removeEventListener("keydown", onKey);
    };
  }, [ctl]);
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
  ctl.saveHook = () => files.save();

  useEffect(() => {
    ctl.openDialog = (id, arg) => {
      if (id === "chartFormat") {
        ctl.formatPane = "chart";
        ctl.emit();
        return;
      }
      if (id === "objectMenu") {
        const a = arg as { x: number; y: number; kind: "chart" | "image" | "shape"; id: string };
        ctx.open({ clientX: a.x, clientY: a.y, preventDefault: () => {} }, <ObjectMenu ctl={ctl} kind={a.kind} id={a.id} />);
        return;
      }
      // İletişim kutuları hücre içeriğiyle çalışır: süren düzenleme önce onaylanır
      if (ctl.edit && id !== "insertFunction" && id !== "alert" && id !== "confirm" && id !== "validationError") ctl.commitEdit(null);
      pushDialog({ id, arg });
    };
    ctl.alert = (message, title) => pushDialog({ id: "alert", arg: { message, title } });
  });

  // Pano olayları (Ctrl+C / X / V)
  useEffect(() => {
    const inGrid = () => {
      const a = document.activeElement;
      if (ctl.selObj && a?.closest(".tb-float") && !backstage) return true;
      return !!a && a.classList.contains("tb-editor-input") && !ctl.edit && !backstage;
    };
    const onCopy = (e: ClipboardEvent) => {
      if (!inGrid()) return;
      // Seçili grafik/resim/şekil kopyalanır (uygulama içi nesne panosu)
      if (ctl.selObj) {
        if (copyObject(ctl, e.type === "cut")) {
          e.preventDefault();
          e.clipboardData?.setData("text/plain", "");
          e.clipboardData?.setData("text/html", objectClipMarker() ?? "");
        }
        return;
      }
      const d = E.copySelection(ctl, e.type === "cut");
      if (!d) return;
      e.preventDefault();
      e.clipboardData?.setData("text/plain", d.tsv);
      e.clipboardData?.setData("text/html", d.html);
    };
    const onPaste = (e: ClipboardEvent) => {
      if (!inGrid()) return;
      e.preventDefault();
      const marker = objectClipMarker();
      if (marker && (e.clipboardData?.getData("text/html") ?? "").includes(marker)) {
        pasteObject(ctl);
        return;
      }
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
  const activeSpark = !ctl.selObj && sheet.sparklines.length ? sparkAt(sheet, ctl.sel.active.row, ctl.sel.active.col)?.group : undefined;
  useEffect(() => {
    if ((tab === "tableDesign" && !activeTable) || ((tab === "chartDesign" || tab === "chartFormat") && !activeChart) || ((tab === "pivotAnalyze" || tab === "pivotDesign") && !activePivot) || (tab === "sparkline" && !activeSpark)) setTab("home");
  }, [tab, activeTable, activeChart, activePivot, activeSpark]);
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
    { id: "pivotAnalyze", label: "Özet Tablo Analizi", contextual: true, contextualGroup: "Özet Tablo Araçları", hidden: !activePivot, content: activePivot ? <PivotAnalyzeTab ctl={ctl} def={activePivot} /> : null },
    { id: "pivotDesign", label: "Tasarım", contextual: true, contextualGroup: "Özet Tablo Araçları", hidden: !activePivot, content: activePivot ? <PivotDesignTab ctl={ctl} def={activePivot} /> : null },
    { id: "sparkline", label: "Mini Grafik", contextual: true, contextualGroup: "Mini Grafik Araçları", hidden: !activeSpark, content: activeSpark ? <SparklineTab ctl={ctl} group={activeSpark} /> : null },
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
      ribbon={
        <ErrorBoundary region="Şerit" compact recovery={recovery}>
          <Ribbon tabs={tabs} activeTab={tab} onTabChange={setTab} onFileClick={() => setBackstage("home")} collapsed={collapsed} onCollapsedChange={setCollapsed} />
        </ErrorBoundary>
      }
      belowRibbon={
        <>
          {ctl.infoBar && (
            <MessageBar
              kind={ctl.infoBar.kind}
              onClose={() => {
                ctl.infoBar = null;
                ctl.emit();
              }}
            >
              {ctl.infoBar.text}
            </MessageBar>
          )}
          {autoRecover.bar}
          {ctl.showFormulaBar && (
            <ErrorBoundary region="Formül çubuğu" compact recovery={recovery}>
              <FormulaBar ctl={ctl} />
            </ErrorBoundary>
          )}
        </>
      }
      statusBar={
        <ErrorBoundary region="Durum çubuğu" compact recovery={recovery}>
          <StatusArea ctl={ctl} />
        </ErrorBoundary>
      }
      backstage={
        backstage ? (
          <ErrorBoundary region="Dosya görünümü" recovery={recovery}>
            <FileView ctl={ctl} files={files} page={backstage} onClose={() => (setBackstage(null), requestAnimationFrame(() => ctl.focusGrid()))} />
          </ErrorBoundary>
        ) : undefined
      }
    >
      <div className="tb-root">
        <div className="tb-workarea">
          <ErrorBoundary region="Çalışma sayfası" recovery={recovery}>
            {sheet.viewMode === "pageLayout" ? (
              <PageLayoutView ctl={ctl} />
            ) : (
              <Grid ctl={ctl} onContextMenu={openContext} onFilterButton={(row, col, x, y) => setFilterMenu({ row, col, x, y })} />
            )}
          </ErrorBoundary>
          {activePivot && ctl.showFieldList && (
            <ErrorBoundary region="Özet Tablo Alanları" compact recovery={recovery}>
              <PivotFieldsPane ctl={ctl} sheet={sheet} def={activePivot} />
            </ErrorBoundary>
          )}
          {ctl.formatPane === "selection" && (
            <ErrorBoundary region="Seçim bölmesi" compact recovery={recovery}>
              <SelectionPane ctl={ctl} />
            </ErrorBoundary>
          )}
          {ctl.formatPane === "chart" && activeChart && (
            <ErrorBoundary region="Biçimlendirme bölmesi" compact recovery={recovery}>
              <ChartFormatPane ctl={ctl} chart={activeChart} />
            </ErrorBoundary>
          )}
        </div>
        {ctl.showWatch && (
          <ErrorBoundary region="İzleme Penceresi" compact recovery={recovery}>
            <WatchWindow ctl={ctl} />
          </ErrorBoundary>
        )}
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
