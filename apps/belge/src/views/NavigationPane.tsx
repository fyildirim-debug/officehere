// Gezinti bölmesi: başlıklar ağacı, sayfalar ve arama sonuçları.
import { useEffect, useMemo, useRef, useState } from "react";
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { Dismiss16Regular, Search16Regular, ChevronDown12Regular, ChevronUp12Regular } from "@fluentui/react-icons";
import { collectHeadings } from "../editor/commands";
import { clearSearch, getSearch, gotoMatch, gotoMatchIndex, setSearch } from "../editor/search";
import type { PaginationController } from "../editor/pagination";

interface Props {
  editor: Editor;
  controller: PaginationController;
  pageCount: number;
  version: number;
  focusSearch: number;
  onClose: () => void;
}

type NavTab = "headings" | "pages" | "results";

export function NavigationPane({ editor, controller, pageCount, version, focusSearch, onClose }: Props) {
  const [tab, setTab] = useState<NavTab>("headings");
  const [query, setQuery] = useState(() => getSearch(editor.state)?.query ?? "");
  const inputRef = useRef<HTMLInputElement>(null);
  const headings = useMemo(() => collectHeadings(editor.state.doc), [editor, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const search = getSearch(editor.state);
  const matches = search?.matches ?? [];
  const curPos = editor.state.selection.from;
  const activeHeading = useMemo(() => {
    let idx = -1;
    headings.forEach((h, i) => {
      if (h.pos <= curPos) idx = i;
    });
    return idx;
  }, [headings, curPos]);

  useEffect(() => {
    if (focusSearch) {
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [focusSearch]);

  useEffect(() => () => clearSearch(editor.view), [editor]);

  const go = (pos: number) => {
    const tr = editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(Math.min(pos + 1, editor.state.doc.content.size)))).scrollIntoView();
    editor.view.dispatch(tr);
    editor.commands.focus();
    requestAnimationFrame(() => {
      const dom = editor.view.domAtPos(pos + 1).node as HTMLElement;
      (dom.nodeType === 1 ? dom : dom.parentElement)?.scrollIntoView({ block: "start" });
    });
  };

  const runSearch = (q: string) => {
    setQuery(q);
    setSearch(editor.view, { query: q, current: 0 });
    setTab(q ? "results" : "headings");
  };

  const snippet = (from: number, to: number) => {
    const $f = editor.state.doc.resolve(from);
    const text = $f.parent.textContent;
    const start = $f.parentOffset;
    const len = to - from;
    const a = Math.max(0, start - 30);
    return (
      <>
        {a > 0 ? "…" : ""}
        {text.slice(a, start)}
        <mark>{text.slice(start, start + len)}</mark>
        {text.slice(start + len, start + len + 40)}
      </>
    );
  };

  return (
    <aside className="bl-nav" aria-label="Gezinti">
      <div className="bl-nav-head">
        <span>Gezinti</span>
        <button type="button" className="bl-icon-btn" onClick={onClose} aria-label="Kapat">
          <Dismiss16Regular />
        </button>
      </div>
      <div className="bl-nav-search">
        <input
          ref={inputRef}
          value={query}
          placeholder="Belgede ara"
          onChange={(e) => runSearch(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              gotoMatch(editor.view, e.shiftKey ? -1 : 1);
            }
            if (e.key === "Escape") {
              runSearch("");
              editor.commands.focus();
            }
          }}
        />
        {query ? (
          <button type="button" className="bl-icon-btn" onClick={() => runSearch("")} aria-label="Aramayı temizle">
            <Dismiss16Regular />
          </button>
        ) : (
          <Search16Regular />
        )}
      </div>
      {query && (
        <div className="bl-nav-count">
          <span>{matches.length ? `${(search?.current ?? 0) + 1} / ${matches.length} sonuç` : "Sonuç yok"}</span>
          <button type="button" className="bl-icon-btn" disabled={!matches.length} onClick={() => gotoMatch(editor.view, -1)} aria-label="Önceki">
            <ChevronUp12Regular />
          </button>
          <button type="button" className="bl-icon-btn" disabled={!matches.length} onClick={() => gotoMatch(editor.view, 1)} aria-label="Sonraki">
            <ChevronDown12Regular />
          </button>
        </div>
      )}
      <div className="bl-nav-tabs" role="tablist">
        {(
          [
            ["headings", "Başlıklar"],
            ["pages", "Sayfalar"],
            ["results", "Sonuçlar"],
          ] as [NavTab, string][]
        ).map(([id, label]) => (
          <button key={id} type="button" role="tab" aria-selected={tab === id} className={tab === id ? "is-active" : ""} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>
      <div className="bl-nav-body oh-scroll">
        {tab === "headings" &&
          (headings.length ? (
            <ul className="bl-nav-tree">
              {headings.map((h, i) => (
                <li key={`${h.pos}-${i}`} style={{ paddingLeft: 8 + (h.level - 1) * 14 }} className={i === activeHeading ? "is-active" : ""}>
                  <button type="button" onClick={() => go(h.pos)} title={h.text}>
                    {h.text}
                  </button>
                </li>
              ))}
            </ul>
          ) : (
            <p className="bl-nav-empty">Belgede başlık yok. Başlık eklemek için Giriş sekmesindeki Stiller galerisinden Başlık stillerini kullanın.</p>
          ))}
        {tab === "pages" && (
          <div className="bl-nav-pages">
            {Array.from({ length: pageCount }, (_, i) => (
              <button key={i} type="button" className={controller.pageOf(curPos) === i ? "is-active" : ""} onClick={() => go(controller.info.pageStarts[i] ?? 0)}>
                <span className="bl-nav-thumb" />
                <span>{i + 1}</span>
              </button>
            ))}
          </div>
        )}
        {tab === "results" &&
          (matches.length ? (
            <ul className="bl-nav-results">
              {matches.slice(0, 500).map((m, i) => (
                <li key={`${m.from}`} className={search?.current === i ? "is-active" : ""}>
                  <button type="button" onClick={() => gotoMatchIndex(editor.view, i)}>
                    {snippet(m.from, m.to)}
                  </button>
                </li>
              ))}
            </ul>
          ) : (
            <p className="bl-nav-empty">{query ? "Eşleşme bulunamadı." : "Metin, açıklama, sayfa ve diğer öğeleri bulmak için arama kutusunu kullanın."}</p>
          ))}
      </div>
    </aside>
  );
}
