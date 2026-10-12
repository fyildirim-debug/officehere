// Çalışma alanı: Yazdırma Düzeni (sayfalar) ve Web Düzeni; cetvel, gezinti ve açıklama bölmeleri.
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { EditorContent } from "@tiptap/react";
import type { Editor, JSONContent } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { pageMetrics, pxToTw, type DocSettings } from "../model/document";
import type { PaginationController } from "../editor/pagination";
import type { ViewOptions } from "../state/context";
import { PageSheets, type PageGeom } from "./PageSheets";
import { HeaderFooterEditor } from "./HeaderFooterEditor";
import { CommentsPane, commentAnchors } from "./CommentsPane";
import { Ruler } from "./Ruler";
import { NavigationPane } from "./NavigationPane";
import { setParaAttrs } from "../editor/commands";
import { collectNotes, setNoteText } from "../editor/notes";

export const PAGE_GAP = 16;
const CANVAS_PAD = 20;
const COMMENTS_W = 270;

interface Props {
  editor: Editor;
  settings: DocSettings;
  updateSettings: (fn: (s: DocSettings) => DocSettings) => void;
  controller: PaginationController;
  view: ViewOptions;
  setView: (p: Partial<ViewOptions>) => void;
  hfEditing: "header" | "footer" | null;
  setHfEditing: (w: "header" | "footer" | null) => void;
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

export function DocumentView(props: Props) {
  const { editor, settings, updateSettings, controller, view, setView, hfEditing, setHfEditing, onHfEditor, version, scrollRef, printing } = props;
  const layout = useSyncExternalStore(
    (cb) => controller.subscribe(cb),
    () => controller.info,
  );
  const pagesRef = useRef<HTMLDivElement>(null);
  const rulerHostRef = useRef<HTMLDivElement>(null);
  const [rulerOffset, setRulerOffset] = useState(0);
  const [hfPage, setHfPage] = useState(0);

  const m = pageMetrics(settings.page);
  const print = view.mode === "print";
  const gap = view.focusMode || printing ? 0 : PAGE_GAP;
  const stride = m.h + gap;
  const scale = printing ? 1 : view.zoom / 100;
  const cols = Math.max(1, settings.page.columns);
  const colGap = settings.page.columnSpace / 15;
  const colW = cols > 1 ? (m.contentW - (cols - 1) * colGap) / cols : m.contentW;
  const g: PageGeom = { w: m.w, h: m.h, mt: m.mt, mb: m.mb, ml: m.ml, mr: m.mr, headerY: m.headerY, footerY: m.footerY, stride };
  const pageCount = print ? layout.pageCount : 1;
  const pagesH = pageCount * stride - gap;

  // Sayfa geometrisini sayfalama eklentisine bildir.
  useEffect(() => {
    controller.setGeometry({ enabled: print, pageHeight: m.h, marginTop: m.mt, marginBottom: m.mb, gap, columns: cols, columnWidth: colW, columnGap: colGap });
  }, [controller, print, m.h, m.mt, m.mb, gap, cols, colW, colGap]);

  const anchors = commentAnchors(editor);
  const doc = editor.state.doc;
  const notes = useMemo(() => new Map(collectNotes(doc).map((n) => [n.id, { label: n.label, text: n.text }])), [doc]);
  const onNoteChange = useCallback((id: string, text: string) => setNoteText(editor.view, id, text), [editor]);
  const showComments = view.showComments && anchors.length > 0 && print;

  // Cetveli sayfayla hizala.
  const syncRuler = useCallback(() => {
    const host = rulerHostRef.current;
    const pages = pagesRef.current;
    if (!host || !pages) return;
    setRulerOffset(pages.getBoundingClientRect().left - host.getBoundingClientRect().left);
  }, []);
  useLayoutEffect(syncRuler, [syncRuler, view.zoom, view.navPane, showComments, m.w, view.ruler, view.mode]);
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
    const page = Math.max(0, Math.min(pageCount - 1, Math.floor(y / stride)));
    return { x, y, page, local: y - page * stride };
  };

  const onPagesMouseDown = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (!print || e.button !== 0) return;
    if (t.closest(".ProseMirror") || t.closest(".bl-hf-edit") || t.closest(".bl-comments") || t.closest(".bl-footnotes")) return;
    e.preventDefault();
    if (hfEditing || e.detail > 1) return;
    const { page, local } = pagePoint(e);
    const pos = local < m.mt ? (controller.info.pageStarts[page] ?? 0) : null;
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
    if (t.closest(".bl-hf-edit") || t.closest(".bl-comments")) return;
    const { page, local } = pagePoint(e);
    if (hfEditing) {
      // Gövdeye çift tık üst/alt bilgiden çıkar.
      if (local > m.mt && local < m.h - m.mb) setHfEditing(null);
      return;
    }
    if (local < m.mt - 2) {
      setHfPage(page);
      setHfEditing("header");
    } else if (local > m.h - m.mb + 2) {
      setHfPage(page);
      setHfEditing("footer");
    }
  };

  const onHfChange = useCallback(
    (json: JSONContent) => {
      updateSettings((s) => (hfEditing === "footer" ? { ...s, footer: json } : { ...s, header: json }));
    },
    [hfEditing, updateSettings],
  );

  const curPara = (() => {
    const { $from } = editor.state.selection;
    for (let d = $from.depth; d >= 0; d--) {
      const n = $from.node(d);
      if (n.type.name === "paragraph" || n.type.name === "heading") return n;
    }
    return null;
  })();

  const rulerVisible = view.ruler && print && !printing && !view.focusMode;
  const canvasW = m.w + (showComments ? COMMENTS_W + 12 : 0) + CANVAS_PAD * 2;

  return (
    <div className={`bl-workspace is-${view.mode}${hfEditing ? " is-hf-editing" : ""}${props.formatPainter ? " is-painter" : ""}`}>
      {view.navPane && !printing && (
        <NavigationPane editor={editor} controller={controller} pageCount={pageCount} version={version} focusSearch={props.navFocus} onClose={() => setView({ navPane: false })} />
      )}
      <div className="bl-main">
        {rulerVisible && (
          <div className="bl-ruler-host" ref={rulerHostRef}>
            <Ruler
              offset={rulerOffset}
              scale={scale}
              pageW={m.w}
              ml={m.ml}
              mr={m.mr}
              indentLeft={(curPara?.attrs.indentLeft as number | null) ?? 0}
              indentRight={(curPara?.attrs.indentRight as number | null) ?? 0}
              indentFirst={(curPara?.attrs.indentFirst as number | null) ?? 0}
              onIndent={(v) => {
                const clean = Object.fromEntries(Object.entries(v).map(([k, x]) => [k, x && Math.abs(x) > 0.01 ? x : null]));
                setParaAttrs(editor, clean);
              }}
              onMargins={(v) =>
                updateSettings((s) => ({
                  ...s,
                  page: { ...s.page, margin: { ...s.page.margin, ...(v.left !== undefined ? { left: pxToTw(v.left) } : {}), ...(v.right !== undefined ? { right: pxToTw(v.right) } : {}) } },
                }))
              }
              onDoubleClick={props.onPageSetup}
            />
          </div>
        )}
        <div className={`bl-scroll oh-scroll${view.gridlines ? " has-grid" : ""}`} ref={scrollRef}>
          <div className="bl-canvas" style={print ? { zoom: scale, width: canvasW, padding: `${CANVAS_PAD}px ${CANVAS_PAD}px` } : { zoom: scale }}>
            <div
              className="bl-pages"
              ref={pagesRef}
              style={print ? { width: m.w + (showComments ? COMMENTS_W + 12 : 0), height: pagesH } : undefined}
              onMouseDown={onPagesMouseDown}
              onDoubleClick={onPagesDoubleClick}
              onContextMenu={props.onContextMenu}
            >
              {print && <PageSheets count={pageCount} g={g} settings={settings} gridlines={view.gridlines} editing={hfEditing} editingPage={hfPage} footnotes={layout.footnotes} notes={notes} onNoteChange={onNoteChange} />}
              <div
                className={`bl-body bl-doc${cols > 1 ? " has-columns" : ""}`}
                style={
                  print
                    ? {
                        top: m.mt,
                        width: m.w,
                        ["--bl-ml" as string]: `${m.ml}px`,
                        ["--bl-mr" as string]: `${m.mr + (m.contentW - colW)}px`,
                        ["--bl-minh" as string]: `${Math.max(0, pagesH - m.mt)}px`,
                        height: cols > 1 ? Math.max(0, pagesH - m.mt) : undefined,
                        background: undefined,
                      }
                    : { background: settings.pageColor ?? undefined }
                }
              >
                <EditorContent editor={editor} />
              </div>
              {hfEditing && print && (
                <HeaderFooterEditor
                  key={hfEditing}
                  which={hfEditing}
                  page={hfPage}
                  g={g}
                  content={hfEditing === "header" ? settings.header : settings.footer}
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
                  left={m.w + 12}
                  activeId={props.activeComment}
                  onActivate={props.setActiveComment}
                  onChange={(id, text) => updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { ...s.comments[id], text } } }))}
                  onDelete={props.deleteComment}
                  onResolve={(id) => updateSettings((s) => ({ ...s, comments: { ...s.comments, [id]: { ...s.comments[id], resolved: !s.comments[id].resolved } } }))}
                  version={version + layout.pageCount}
                />
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
