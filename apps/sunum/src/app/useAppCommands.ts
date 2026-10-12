/** Şerit, menü ve kısayolların kullandığı uygulama komutları (App.tsx'ten ayrıldı). */
import { useMemo, useRef } from "react";
import { getNative } from "@officehere/core";
import { makeTextBox } from "../model/factory";
import { MASTER_SLIDE_ID } from "../model/masterView";
import { sectionsOf } from "../model/sections";
import { WORDART_PRESETS } from "../ribbon/presets";
import type { AppCommands, DialogKind } from "../ribbon/types";
import { runOptions, showOrder, type ShowSettings } from "../show/settings";
import { newChannelId, openAudienceWindow, PresenterLink } from "../show/showSync";
import type { Actions } from "../state/actions";
import { currentLink, followLink, linkContext, removeLink } from "../state/linkActions";
import { closeMasterView, documentPres, masterLayoutName, masterRenameLayout, openMasterView } from "../state/masterActions";
import { pickMedia, selectedChart } from "../state/mediaActions";
import { addSection, renameSectionCmd } from "../state/sectionActions";
import type { PaneKind, ViewMode } from "../state/store";
import type { ExtraDialog } from "./AppDialogs";
import { stageFrame, type ShowState } from "./appState";
import { insertImageFiles, pasteFromRibbon } from "./clipboard";

export interface CommandDeps {
  a: Actions;
  setDialog: (d: DialogKind | null) => void;
  setExtra: (d: ExtraDialog | null) => void;
  setShow: (s: ShowState | null) => void;
  setBackstage: (v: boolean) => void;
  setPrinting: (v: null | "print" | "pdf") => void;
  setFindQuery: (q: string) => void;
  showSettings: ShowSettings;
}

export function useAppCommands(d: CommandDeps): AppCommands {
  const { a, setDialog, setExtra, setShow, setBackstage, setPrinting, setFindQuery, showSettings } = d;
  const cmdRef = useRef<AppCommands | null>(null);

  const cmd: AppCommands = useMemo(() => {
    const store = a.store;
    /** Gösteri başlamadan önceki ortak hazırlık. */
    const prepare = () => {
      if (a.s.masterBase) closeMasterView(a);
      a.stopEditing();
      return a.s.pres.slides.length > 0;
    };
    const fullscreen = () => document.documentElement.requestFullscreen?.().catch(() => undefined);
    /** Sunucu Görünümü: izleyici gösterisi ikinci pencerede, sunucu konsolu bu pencerede. */
    const startPresenter = (start: number, opts: ShowState["opts"]) => {
      const id = newChannelId();
      const link = new PresenterLink(id, () => ({ pres: documentPres(a), opts: opts ?? {} }));
      const win = openAudienceWindow(id);
      if (!win) {
        link.end();
        a.message("İzleyici penceresi açılamadı (açılır pencere engellendi); Sunucu Görünümü tek pencerede gösteriliyor.", "warning");
        fullscreen();
        setShow({ mode: "presenter", start, opts });
        return;
      }
      setShow({ mode: "presenter", start, opts, link });
    };
    const startWith = (from: "start" | "current" | number, extra: Partial<ShowState["opts"]> = {}, settings = showSettings) => {
      if (!prepare()) return;
      const p = a.s.pres;
      const opts = { ...runOptions(p, settings), ...extra };
      const order = opts.order ?? showOrder(p, { kind: "all" });
      let start = typeof from === "number" ? from : from === "start" ? (order[0] ?? 0) : a.s.slideIdx;
      if (from === "start" && !order.length) start = 0;
      if (settings.presenter && settings.type === "speaker" && !extra.rehearse) return startPresenter(start, opts);
      if (settings.type === "browsed") {
        setShow({ mode: "window", start, opts });
        return;
      }
      // Tam ekran isteği kullanıcı hareketi içinde yapılmalı.
      fullscreen();
      setShow({ mode: "full", start, opts });
    };
    const c: AppCommands = {
      paste: (mode) => {
        if (mode === "text" && store.clipboard?.kind === "shapes") {
          const text = (store.clipboard.shapes ?? []).map((s) => (s.kind === "sp" && s.text ? s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n") : "")).filter(Boolean).join("\n");
          if (text) a.pasteText(text);
          return;
        }
        void pasteFromRibbon(a);
      },
      cut: () => {
        if (a.s.editing) document.execCommand("cut");
        else if (a.s.focus === "slides") a.copySlides(true);
        else a.copySelection(true);
      },
      copy: () => {
        if (a.s.editing) document.execCommand("copy");
        else if (a.s.focus === "slides" || !a.s.sel.length) a.copySlides(false);
        else a.copySelection(false);
      },
      insertPictures: () => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = "image/*";
        input.multiple = true;
        input.onchange = () => {
          void insertImageFiles(a, Array.from(input.files ?? [])).then(() => a.store.set({ ribbonTab: "pictureFormat" }));
        };
        input.click();
      },
      startShow: (from) => startWith(from),
      presenterView: () => {
        if (!prepare()) return;
        const opts = runOptions(a.s.pres, showSettings);
        startPresenter(a.s.slideIdx, opts);
      },
      rehearseTimings: () => startWith("start", { rehearse: true, manual: true, loop: false, kiosk: false }, { ...showSettings, type: "speaker", presenter: false }),
      startCustomShow: (id) => startWith("start", {}, { ...showSettings, range: { kind: "custom", id } }),
      customShows: () => setExtra({ kind: "customShows" }),
      print: () => {
        a.stopEditing();
        if (a.s.masterBase) closeMasterView(a);
        setBackstage(false);
        setPrinting("print");
      },
      exportPdf: () => {
        if (!getNative()?.exportPdf) {
          cmdRef.current?.print();
          return;
        }
        a.stopEditing();
        if (a.s.masterBase) closeMasterView(a);
        setBackstage(false);
        setPrinting("pdf");
      },
      previewTransition: () => {
        if (a.s.pres.slides.length) setShow({ mode: "window", start: a.s.slideIdx, preview: "transition", frame: stageFrame() });
      },
      previewAnimations: () => {
        if (a.s.pres.slides.length) setShow({ mode: "window", start: a.s.slideIdx, preview: "animations", frame: stageFrame() });
      },
      setView: (v: ViewMode) => {
        a.stopEditing();
        if (v === "master") openMasterView(a);
        else if (a.s.masterBase) closeMasterView(a, v);
        else a.store.set({ view: v, sel: [] });
      },
      togglePane: (p: PaneKind) => a.store.set((s) => ({ pane: s.pane === p ? null : p })),
      pickShape: (item) => {
        if (item.geom === "textbox") a.store.set({ insert: { kind: "textbox" } });
        else a.store.set({ insert: { kind: "shape", geom: item.geom, head: item.head, tail: item.tail } });
      },
      insertWordArt: (i) => {
        const w = WORDART_PRESETS[i];
        const p = a.pres;
        const tb = makeTextBox(p.width * 0.2, p.height * 0.38, p.width * 0.6, 120, "Metninizi buraya yazın");
        tb.name = `Sanatsal Metin ${a.nextShapeNumber()}`;
        tb.text = { ...tb.text!, wrap: true, levels: [{ para: { align: "ctr" }, run: { size: 54, color: w.color, outline: w.outline, shadow: w.shadow || undefined, glow: w.glow ? w.color : undefined, bold: w.bold } }] };
        const i0 = a.s.slideIdx;
        a.store.commit(
          (pp) => {
            const slides = pp.slides.slice();
            slides[i0] = { ...slides[i0], shapes: [...slides[i0].shapes, tb] };
            return { ...pp, slides };
          },
          { sel: [tb.id], editing: { shapeId: tb.id, selectAll: true } },
        );
        a.store.set({ ribbonTab: "shapeFormat" });
      },
      insertDateTime: () => setDialog("dateTime"),
      insertSlideNumber: () => {
        if (a.s.editing && a.textHandle) a.textHandle.insertRun("‹#›", { field: "slidenum" });
        else setDialog("headerFooter");
      },
      spelling: () => a.store.set((s) => ({ pane: s.pane === "spelling" ? null : "spelling" })),
      zoomTo: (z, fit) => a.store.set({ zoom: Math.max(10, Math.min(400, z)), fit: !!fit }),
      insertLink: () => setExtra({ kind: "link", ctx: linkContext(a) }),
      insertChart: () => setExtra({ kind: "chartType" }),
      changeChartType: () => {
        const ch = selectedChart(a);
        if (ch) setExtra({ kind: "chartType", changeId: ch.id, current: ch.chart });
      },
      editChartData: () => {
        const ch = selectedChart(a);
        if (ch) setExtra({ kind: "chartData", id: ch.id, chart: ch.chart });
      },
      insertMedia: (kind) => pickMedia(a, kind),
      playMedia: () => {
        const sh = a.selectedShapes()[0];
        if (sh?.kind === "pic" && sh.media) setExtra({ kind: "media", shape: sh });
      },
      openLink: () => {
        const l = currentLink(a);
        if (l) followLink(a, l);
      },
      removeLink: () => removeLink(a),
      addSection: () => {
        const id = addSection(a);
        if (id) cmdRef.current?.renameSection(id);
      },
      renameLayout: () => {
        if (!a.s.masterBase || a.slide()?.id === MASTER_SLIDE_ID) return;
        setExtra({ kind: "rename", req: { title: "Düzeni Yeniden Adlandır", label: "Düzen adı:", value: masterLayoutName(a), onOk: (n) => masterRenameLayout(a, n) } });
      },
      renameSection: (id) => {
        const sec = sectionsOf(a.s.pres.slides).find((x) => x.id === id);
        if (sec) setExtra({ kind: "rename", req: { title: "Bölümü Yeniden Adlandır", label: "Bölüm adı:", value: sec.name, onOk: (n) => renameSectionCmd(a, id, n) } });
      },
      find: (q) => {
        setFindQuery(q ?? "");
        setDialog("find");
      },
      setSlideSize: (w, h) => setExtra({ kind: "scaleChoice", w, h }),
      reuseSlides: () => a.store.set({ pane: "reuse" }),
    };
    return c;
  }, [a, showSettings, setDialog, setExtra, setShow, setBackstage, setPrinting, setFindQuery]);
  cmdRef.current = cmd;
  return cmd;
}
