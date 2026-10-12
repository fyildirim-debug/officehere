// Çalışma alanı: Yazdırma Düzeni (sayfalar) ve Web Düzeni; cetveller, gezinti ve açıklama bölmeleri.
// Sayfalar düzen bilgisindeki (bölüm geometrili) konumlarla çizilir; düşük yakınlaştırmada
// sayfalar yan yana dizilir (Birden Çok Sayfa).
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { EditorContent, useEditorState } from "@tiptap/react";
import type { Editor, JSONContent } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { pxToTw, type DocSettings } from "../model/document";
import type { PageInfo, PaginationController } from "../editor/pagination";
import type { ViewOptions } from "../state/context";
import { PageSheets, hfKey, hfVariantOf, hfLabel } from "./PageSheets";
import { HeaderFooterEditor } from "./HeaderFooterEditor";
import { CommentsPane, commentAnchors } from "./CommentsPane";
import { Ruler } from "./Ruler";
import { VRuler } from "./VRuler";
import { NavigationPane } from "./NavigationPane";
import { setParaAttrs } from "../editor/commands";
import { collectNotes, setNoteContent } from "../editor/notes";
import { sectionGeoms, type SecGeom } from "../editor/sections";
import { applyPageSetup } from "../state/pageSetup";

export const PAGE_GAP = 16;
const CANVAS_PAD = 20;
const COMMENTS_W = 270;
/** Yan yana sayfalar arası yatay boşluk. */
export const TILE_GAP = 16;

interface Props {
  editor: Editor;
  settings: DocSettings;
  updateSettings: (fn: (s: DocSettings) => DocSettings) => void;
  controller: PaginationController;
  view: ViewOptions;
  setView: (p: Partial<ViewOptions>) => void;
  hfEditing: "header" | "footer" | null;
  hfPage: number;
  setHfEditing: (w: "header" | "footer" | null, page?: number) => void;
  onHfEditor: (ed: Editor | null) => void;
  activeComment: string | null;
  setActiveComment: (id: string | null) => void;
  deleteComment: (id: string) => void;
  version: number;
  navFocus: number;
  onPageSetup: () => void;
  onContextMenu: (e: React.MouseEvent) => void;
  scrollRef: React.RefObject<HTMLDivElement | null>;
  printing: boolean;
  formatPainter: boolean;
}

/** Sayfanın sayfa kapsayıcısındaki dikdörtgeni (yan yana dizimde satır/sütun konumu). */
export interface PageRect {
  x: number;
  y: number;
}

export function DocumentView(props: Props) {
  const { editor, settings, updateSettings, controller, view, setView, hfEditing, hfPage, setHfEditing, onHfEditor, version, scrollRef, printing } = props;
  // Düzen bilgisi: sayfalar ve dipnot yerleşimi (başvuruları yalnız değişince yenilenir).
  const subscribe = useCallback((cb: () => void) => controller.subscribe(cb), [controller]);
  const layoutPages = useSyncExternalStore(subscribe, () => controller.info.pages);
  const layoutNotes = useSyncExternalStore(subscribe, () => controller.info.footnotes);
  const layoutBodyW = useSyncExternalStore(subscribe, () => controller.info.bodyW);
  const tileOk = useSyncExternalStore(subscribe, () => controller.info.tileOk !== false);
  const pagesRef = useRef<HTMLDivElement>(null);
  const rulerHostRef = useRef<HTMLDivElement>(null);
  const vrulerHostRef = useRef<HTMLDivElement>(null);
  const [rulerOffset, setRulerOffset] = useState(0);
  const [vOffset, setVOffset] = useState(0);
  const [scrollW, setScrollW] = useState(1200);

  const print = view.mode === "print";
  const gap = view.focusMode || printing ? 0 : PAGE_GAP;
  const scale = printing ? 1 : view.zoom / 100;
  const doc = editor.state.doc;
  // İlk bölümün geometrisi gövdenin üst konumunu belirler (ölçüm öncesi de doğru olsun diye doğrudan hesaplanır).
  const geoms = useMemo(() => sectionGeoms(doc, settings.page), [doc, settings.page]);
  const pages: PageInfo[] = print ? layoutPages : [];
  const bodyW = print ? Math.max(layoutBodyW || 0, ...geoms.map((g) => g.w)) : 0;
  const pageCount = print ? pages.length : 1;
  const first = geoms[0];
  const lastPage = pages[pages.length - 1];
  const flowH = lastPage ? lastPage.top + lastPage.h : first.h;

  // Birden Çok Sayfa: sayfa boyları eşitse ve pencereye en az iki sayfa sığıyorsa yan yana dizilir.
  const uniform = pages.length > 1 && pages.every((p) => Math.abs(p.h - pages[0].h) < 0.5 && Math.abs(p.mt - pages[0].mt) < 0.5);
  const perRowFit = Math.floor((scrollW / scale - 2 * CANVAS_PAD + TILE_GAP) / (bodyW + TILE_GAP));
  const tiled = print && !printing && !hfEditing && !view.focusMode && uniform && perRowFit >= 2;
  // Satır kaydırılamıyorsa (tablo satırı ya da çok sütunlu bölüm sınırı) sayfalar tek satırda dizilir.
  const perRow = tiled ? (tileOk ? Math.min(perRowFit, pageCount) : pageCount) : 1;
  const stride = (pages[0]?.h ?? first.h) + gap;
  const rows = tiled ? Math.ceil(pageCount / perRow) : pageCount;

  // Sayfa geometrisini sayfalama eklentisine bildir.
  useEffect(() => {
    controller.setGeometry({ enabled: print, gap, defaultPage: settings.page, tiles: tiled ? Math.min(perRowFit, pageCount) : 0 });
  }, [controller, print, gap, settings.page, tiled, perRowFit, pageCount]);

  // Gövdenin üst konumu değişince (ilk bölüm kenar boşluğu) yeniden ölç.
  const bodyTop = first.mt;
  useEffect(() => {
    controller.request();
  }, [controller, bodyTop, bodyW]);

  const rectOf = useCallback(
    (i: number): PageRect => {
      const p = pages[i];
      if (!p) return { x: 0, y: 0 };
      if (!tiled) return { x: p.left, y: p.top };
      const r = Math.floor(i / perRow);
      const c = i % perRow;
      return { x: c * (bodyW + TILE_GAP) + p.left, y: r * stride };
    },
    [pages, tiled, perRow, bodyW, stride],
  );

  const anchors = commentAnchors(editor);
  const notes = useMemo(() => new Map(collectNotes(doc).map((n) => [n.id, { label: n.label, text: n.text, content: n.content }])), [doc]);
  const onNoteChange = useCallback((id: string, json: JSONContent) => setNoteContent(editor.view, id, (json.content ?? []) as JSONContent[]), [editor]);
  const showComments = view.showComments && anchors.length > 0 && print && !tiled;

  // Cetvelleri sayfayla hizala.
  const syncRuler = useCallback(() => {
    const host = rulerHostRef.current;
    const pagesEl = pagesRef.current;
    const sc = scrollRef.current;
    if (sc) setScrollW(sc.clientWidth);
    if (!pagesEl) return;
    const pr = pagesEl.getBoundingClientRect();
    if (host) setRulerOffset(pr.left - host.getBoundingClientRect().left);
    const vh = vrulerHostRef.current;
    if (vh) setVOffset(pr.top - vh.getBoundingClientRect().top);
  }, [scrollRef]);
  useLayoutEffect(syncRuler, [syncRuler, view.zoom, view.navPane, showComments, bodyW, view.ruler, view.mode, pages]);
  useEffect(() => {
    const sc = scrollRef.current;
    if (!sc) return;
    const ro = new ResizeObserver(syncRuler);
    ro.observe(sc);
    sc.addEventListener("scroll", syncRuler, { passive: true });
    return () => {
      ro.disconnect();
      sc.removeEventListener("scroll", syncRuler);
    };
  }, [scrollRef, syncRuler]);

  // Ctrl + fare tekerleği ile yakınlaştırma.
  useEffect(() => {
    const sc = scrollRef.current;
    if (!sc) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const z = view.zoom;
      const next = Math.min(500, Math.max(10, Math.round((z + (e.deltaY < 0 ? 10 : -10)) / 10) * 10));
      if (next !== z) setView({ zoom: next });
    };
    sc.addEventListener("wheel", onWheel, { passive: false });
    return () => sc.removeEventListener("wheel", onWheel);
  }, [scrollRef, view.zoom, setView]);

  // Sayfa zeminine (editör dışına) tıklanınca imleci en yakın konuma koy; üst/alt bilgiye çift tık.
  const pagePoint = (e: React.MouseEvent) => {
    const r = pagesRef.current!.getBoundingClientRect();
    const x = (e.clientX - r.left) / scale;
    const y = (e.clientY - r.top) / scale;
    let page = 0;
    for (let i = 0; i < pages.length; i++) {
      const pr = rectOf(i);
      if (y >= pr.y - gap / 2 && (tiled ? x >= pr.x - TILE_GAP / 2 && x <= pr.x + pages[i].w + TILE_GAP / 2 : true)) page = i;
      if (tiled && y >= pr.y - gap / 2 && y <= pr.y + pages[i].h && x >= pr.x - TILE_GAP / 2 && x <= pr.x + pages[i].w + TILE_GAP / 2) {
        page = i;
        break;
      }
    }
    const pr = rectOf(page);
    return { x: x - pr.x, page, local: y - pr.y };
  };

  const onPagesMouseDown = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (!print || e.button !== 0) return;
    if (t.closest(".ProseMirror") || t.closest(".bl-hf-edit") || t.closest(".bl-comments") || t.closest(".bl-footnotes")) return;
    e.preventDefault();
    if (hfEditing || e.detail > 1) return;
    const { page, local } = pagePoint(e);
    const pg = pages[page];
    const pos = pg && local < pg.mt ? (controller.info.pageStarts[page] ?? 0) : null;
    if (pos !== null) {
      const sel = TextSelection.near(editor.state.doc.resolve(Math.min(pos + 1, editor.state.doc.content.size)));
      editor.view.dispatch(editor.state.tr.setSelection(sel));
    } else {
      const hit = editor.view.posAtCoords({ left: e.clientX, top: e.clientY });
      if (hit) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(hit.pos))));
    }
    // Kaydırmadan odaklan (çift tıklamada ikinci tıklama aynı noktaya düşsün).
    editor.view.focus();
  };

  const onPagesDoubleClick = (e: React.MouseEvent) => {
    if (!print) return;
    const t = e.target as HTMLElement;
    if (t.closest(".bl-hf-edit") || t.closest(".bl-comments") || t.closest(".bl-footnotes")) return;
    const { page, local } = pagePoint(e);
    const pg = pages[page];
    if (!pg) return;
    if (hfEditing) {
      // Gövdeye çift tık üst/alt bilgiden çıkar.
      if (local > pg.mt && local < pg.h - pg.mb) setHfEditing(null);
      return;
    }
    if (local < pg.mt - 2) setHfEditing("header", page);
    else if (local > pg.h - pg.mb + 2) setHfEditing("footer", page);
  };

  const hfVariant = hfEditing ? hfVariantOf(pages[hfPage], hfPage, settings) : "default";
  const hfField = hfEditing ? hfKey(hfEditing, hfVariant) : null;
  const onHfChange = useCallback(
    (json: JSONContent) => {
      if (hfField) updateSettings((s) => ({ ...s, [hfField]: json }));
    },
    [hfField, updateSettings],
  );

  const setMargins = (section: number | undefined) => (v: { left?: number; right?: number; top?: number; bottom?: number }) =>
    applyPageSetup(
      editor,
      settings,
      updateSettings,
      (p) => ({
        ...p,
        margin: {
          ...p.margin,
          ...(v.left !== undefined ? { left: pxToTw(v.left) } : {}),
          ...(v.right !== undefined ? { right: pxToTw(v.right) } : {}),
          ...(v.top !== undefined ? { top: pxToTw(v.top) } : {}),
          ...(v.bottom !== undefined ? { bottom: pxToTw(v.bottom) } : {}),
        },
      }),
      "auto",
      section,
    );

  const rulerVisible = view.ruler && print && !printing && !view.focusMode && !tiled;
  const contentW = tiled ? perRow * bodyW + (perRow - 1) * TILE_GAP : bodyW;
  const pagesW = contentW + (showComments ? COMMENTS_W + 12 : 0);
  const pagesH = tiled ? rows * stride - gap : flowH;
  const canvasW = pagesW + CANVAS_PAD * 2;
  const hfRect = hfEditing ? rectOf(hfPage) : null;
  const hfPg = hfEditing ? (pages[hfPage] ?? pages[0]) : null;

  return (
    <div className={`bl-workspace is-${view.mode}${hfEditing ? " is-hf-editing" : ""}${props.formatPainter ? " is-painter" : ""}${tiled ? " is-tiled" : ""}`}>
      {view.navPane && !printing && (
        <NavigationPane editor={editor} controller={controller} pageCount={pageCount} version={version} focusSearch={props.navFocus} onClose={() => setView({ navPane: false })} />
      )}
      <div className="bl-main">
        {rulerVisible && (
          <div className="bl-ruler-host" ref={rulerHostRef}>
            <div className="bl-ruler-corner" />
            <HRulerBar editor={editor} controller={controller} pages={pages} geoms={geoms} rectOf={rectOf} offset={rulerOffset} scale={scale} pagesRef={pagesRef} setMargins={setMargins} onPageSetup={props.onPageSetup} />
          </div>
        )}
        <div className="bl-scroll-row">
          {rulerVisible && (
            <div className="bl-vruler-host" ref={vrulerHostRef}>
              <VRulerBar editor={editor} controller={controller} pages={pages} rectOf={rectOf} offset={vOffset} scale={scale} setMargins={setMargins} onPageSetup={props.onPageSetup} />
            </div>
          )}
          <div className={`bl-scroll oh-scroll${view.gridlines ? " has-grid" : ""}`} ref={scrollRef}>
            <div className="bl-canvas" style={print ? { zoom: scale, width: canvasW, padding: `${CANVAS_PAD}px ${CANVAS_PAD}px` } : { zoom: scale }}>
              <div
                className="bl-pages"
                ref={pagesRef}
                style={print ? { width: pagesW, height: pagesH } : undefined}
                onMouseDown={onPagesMouseDown}
                onDoubleClick={onPagesDoubleClick}
                onContextMenu={props.onContextMenu}
              >
                {print && (
                  <PageSheets
                    pages={pages}
                    rects={tiled ? pages.map((_, i) => rectOf(i)) : null}
                    settings={settings}
                    gridlines={view.gridlines}
                    editing={hfEditing}
                    editingPage={hfPage}
                    footnotes={layoutNotes}
                    notes={notes}
                    onNoteChange={onNoteChange}
                    sectionCount={geoms.length}
                  />
                )}
                <div
                  className={`bl-body bl-doc${tiled ? " is-tiled" : ""}`}
                  style={
                    print
                      ? {
                          top: bodyTop,
                          width: tiled ? contentW : bodyW,
                          ["--bl-minh" as string]: `${Math.max(0, flowH - bodyTop)}px`,
                          ...(tiled ? { height: rows * stride, columnCount: perRow, columnGap: TILE_GAP, columnFill: "auto" as const, ["--bl-bodyw" as string]: `${bodyW}px` } : {}),
                        }
                      : { background: settings.pageColor ?? undefined }
                  }
                >
                  <EditorContent editor={editor} />
                </div>
                {hfEditing && print && hfPg && hfRect && hfField && (
                  <HeaderFooterEditor
                    key={`${hfField}:${hfPage}`}
                    which={hfEditing}
                    label={hfLabel(hfEditing, hfVariant, settings.oddEven, hfPg.section, geoms.length)}
                    page={hfPg}
                    rect={hfRect}
                    content={settings[hfField]}
                    onChange={onHfChange}
                    onEditor={onHfEditor}
                    onClose={() => setHfEditing(null)}
                  />
                )}
                {showComments && !printing && (
                  <CommentsPane
                    editor={editor}
                    comments={settings.comments}
                    anchors={anchors}
                    container={pagesRef.current}
                    scale={scale}
                    left={bodyW + 12}
                    activeId={props.activeComment}
                    onActivate={props.setActiveComment}
                    onChange={(id, text) => updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { ...s.comments[id], text } } }))}
                    onDelete={props.deleteComment}
                    onResolve={(id) => updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { ...s.comments[id], resolved: !s.comments[id].resolved } } }))}
                    version={version + pageCount}
                  />
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/** İmleçteki sayfa (yalnız sayfa değişince yeniden çizer; düzen bilgisi ve seçim izlenir). */
function useCaretPage(editor: Editor, controller: PaginationController, count: number) {
  const sub = useCallback(
    (cb: () => void) => {
      const off = controller.subscribe(cb);
      editor.on("transaction", cb);
      return () => {
        off();
        editor.off("transaction", cb);
      };
    },
    [editor, controller],
  );
  const page = useSyncExternalStore(sub, () => Math.max(0, Math.min(count - 1, controller.pageOf(editor.state.selection.from))));
  return { sub, page };
}

interface BarProps {
  editor: Editor;
  controller: PaginationController;
  pages: PageInfo[];
  rectOf: (i: number) => PageRect;
  offset: number;
  scale: number;
  setMargins: (section: number | undefined) => (v: { left?: number; right?: number; top?: number; bottom?: number }) => void;
  onPageSetup: () => void;
}

function HRulerBar({ editor, controller, pages, geoms, rectOf, offset, scale, pagesRef, setMargins, onPageSetup }: BarProps & { geoms: SecGeom[]; pagesRef: React.RefObject<HTMLDivElement | null> }) {
  const { sub, page } = useCaretPage(editor, controller, pages.length);
  const cur = pages[page];
  const sec = cur ? (geoms[cur.section] ?? geoms[0]) : null;
  const rect = rectOf(page);
  // Çok sütunlu bölümde imlecin sütunu (girintiler o sütuna göre çizilir).
  const col = useSyncExternalStore(sub, () => {
    if (!sec || sec.cols < 2) return 0;
    try {
      const c = editor.view.coordsAtPos(editor.state.selection.from);
      const pr = pagesRef.current?.getBoundingClientRect();
      if (!pr) return 0;
      return Math.max(0, Math.min(sec.cols - 1, Math.floor(((c.left - pr.left) / scale - rect.x - sec.ml) / (sec.colW + sec.colGap))));
    } catch {
      return 0;
    }
  });
  // Cetvel girintileri: imleçteki paragrafın değerleri.
  const ind = useEditorState({
    editor,
    selector: ({ editor: ed }) => {
      const { $from } = ed.state.selection;
      for (let d = $from.depth; d >= 0; d--) {
        const n = $from.node(d);
        if (n.type.name === "paragraph" || n.type.name === "heading")
          return { l: (n.attrs.indentLeft as number | null) ?? 0, r: (n.attrs.indentRight as number | null) ?? 0, f: (n.attrs.indentFirst as number | null) ?? 0 };
      }
      return { l: 0, r: 0, f: 0 };
    },
    equalityFn: (a, b) => !!a && !!b && a.l === b.l && a.r === b.r && a.f === b.f,
  });
  if (!cur || !sec) return null;
  const columns = sec && sec.cols > 1 ? Array.from({ length: sec.cols }, (_, i) => ({ left: sec.ml + i * (sec.colW + sec.colGap), width: sec.colW })) : undefined;
  return (
    <Ruler
      offset={offset + rect.x * scale}
      scale={scale}
      pageW={cur.w}
      ml={cur.ml}
      mr={cur.mr}
      columns={columns}
      column={col}
      indentLeft={ind?.l ?? 0}
      indentRight={ind?.r ?? 0}
      indentFirst={ind?.f ?? 0}
      onIndent={(v) => {
        const clean = Object.fromEntries(Object.entries(v).map(([k, x]) => [k, x && Math.abs(x) > 0.01 ? x : null]));
        setParaAttrs(editor, clean);
      }}
      onMargins={setMargins(cur.section)}
      onDoubleClick={onPageSetup}
    />
  );
}

function VRulerBar({ editor, controller, pages, rectOf, offset, scale, setMargins, onPageSetup }: BarProps) {
  const { page } = useCaretPage(editor, controller, pages.length);
  const cur = pages[page];
  if (!cur) return null;
  return <VRuler offset={offset + rectOf(page).y * scale} scale={scale} pageH={cur.h} mt={cur.mt} mb={cur.mb} onMargins={setMargins(cur.section)} onDoubleClick={onPageSetup} />;
}
