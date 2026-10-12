/**
 * Başlık çubuğundaki "Ara (Alt+Q)" komut araması (Belge, Tablo ve Sunum'da ortak).
 * Şerit komutlarından (ipucu adı + kısayol + açıklama) ve uygulamanın verdiği ek
 * komutlardan (Dosya menüsü vb.) bir dizin oluşturur, sonuçları açılır listede
 * gösterir ve seçilen komutu çalıştırır. Kutu boşken son kullanılan komutlar
 * listelenir; listenin sonunda uygulamanın "bul" satırı bulunur.
 *
 * Dizin, tüm sekmelerin içeriği görünmeyen bir kapta bir kez çizilerek
 * çıkarılır; komut çalıştırılırken ilgili sekmeye geçilip aynı düğmeye
 * (daraltılmış grupta grup açılarak) tıklanır.
 */
import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";
import {
  DocumentAdd20Regular,
  DocumentArrowRight20Regular,
  DocumentPdf20Regular,
  FolderOpen20Regular,
  Info20Regular,
  Print20Regular,
  Save20Regular,
  SaveEdit20Regular,
  Search20Regular,
  Settings20Regular,
} from "@fluentui/react-icons";
import type { RibbonTabDef } from "./Ribbon";
import { loadRecentCommands, pushRecentCommand, rankCommands, resolveRecentCommands, saveRecentCommands, type CommandEntry } from "./commandIndex";
import { activateCommand, isShowing, locateCommand, mergeCommands, nextFrame, scrapeCommands } from "./CommandSearchDom";

const MAX_RESULTS = 8;

/** Listenin sonundaki "bul" satırı (belgede/sayfada metin arama). */
export interface CommandSearchFind {
  /** Satırın metni (ör. q => `Belgede bul: ${q}`). */
  label: (query: string) => string;
  /** Satırın üstündeki bölüm başlığı; verilmezse ince bir ayraç çizilir. */
  heading?: string;
  /** Soluk ikinci satır (ör. "Giriş › Düzenleme"). */
  path?: string;
  shortcut?: string;
  run: (query: string) => void;
}

export interface CommandSearchProps {
  /** Son kullanılanların saklandığı anahtarı ayırır ("belge", "tablo", "sunum"). */
  app: string;
  /** Şerit sekmeleri (dizin bunların içeriğinden çıkarılır). */
  tabs: RibbonTabDef[];
  /** Şerit daraltılmış mı (komut çalıştırılırken açılır). */
  collapsed: boolean;
  onExpand: () => void;
  /** Komutun sekmesine geçer (Dosya görünümü açıksa kapatılmalı). */
  onTabChange: (id: string) => void;
  /** Şeritte düğmesi olmayan komutlar (Dosya eylemleri, galeri öğeleri). */
  extra?: CommandEntry[];
  find?: CommandSearchFind;
  /** true: hiçbir satır önceden seçili gelmez; seçim yokken Enter "bul" satırını çalıştırır. */
  findOnEnter?: boolean;
  /** Komut o an kullanılamıyorsa gösterilecek bilgi iletisi. */
  onMessage?: (text: string) => void;
  /** Esc ile aramadan çıkınca odağı belgeye/ızgaraya döndürür. */
  onDismiss?: () => void;
  placeholder?: string;
}

/** Dosya komutlarının simgeleri (kimliğe göre; bkz. fileCommands). */
const FILE_ICONS: Record<string, ReactNode> = {
  "file:new": <DocumentAdd20Regular />,
  "file:open": <FolderOpen20Regular />,
  "file:save": <Save20Regular />,
  "file:saveas": <SaveEdit20Regular />,
  "file:pdf": <DocumentPdf20Regular />,
  "file:print": <Print20Regular />,
  "file:export": <DocumentArrowRight20Regular />,
  "file:info": <Info20Regular />,
  "file:options": <Settings20Regular />,
};

type Row = { kind: "cmd"; c: CommandEntry } | { kind: "find" };

export function CommandSearch({ app, tabs, collapsed, onExpand, onTabChange, extra, find, findOnEnter = false, onMessage, onDismiss, placeholder = "Ara (Alt+Q)" }: CommandSearchProps) {
  const [host, setHost] = useState<HTMLElement | null>(null);
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(-1);
  const [index, setIndex] = useState<CommandEntry[] | null>(null);
  const [indexing, setIndexing] = useState(false);
  const [recent, setRecent] = useState(() => loadRecentCommands(app));
  const [pos, setPos] = useState<{ left: number; top: number; width: number } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const boxRef = useRef<HTMLLabelElement>(null);
  const hiddenRef = useRef<HTMLDivElement>(null);
  const listId = `oh-cmd-list-${useId().replace(/[^\w-]/g, "")}`;
  const optId = (i: number) => `${listId}-${i}`;

  // Başlık çubuğundaki arama yuvasına yerleş (başlık çubuğu yeniden çizilirse yeniden bulunur).
  useLayoutEffect(() => {
    if (host?.isConnected) return;
    const el = document.querySelector<HTMLElement>(".oh-titlebar-search");
    if (el !== host) setHost(el);
  });

  // Alt+Q: aramaya odaklan. Ctrl+Alt (AltGr+Q = "@") yazıyı engellemesin diye dışarıda kalır.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.altKey && !e.ctrlKey && !e.metaKey && (e.key === "q" || e.key === "Q" || e.code === "KeyQ")) {
        e.preventDefault();
        inputRef.current?.focus();
        inputRef.current?.select();
      }
    };
    // Yakalama aşamasında: görev bölmeleri tuş olaylarını durdursa da Alt+Q çalışsın.
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, []);

  // Dizin: gizli kapta çizilen sekmelerden çıkarılır.
  useLayoutEffect(() => {
    if (!indexing || !hiddenRef.current) return;
    setIndex(mergeCommands(scrapeCommands(hiddenRef.current, tabs), extra));
    setIndexing(false);
  }, [indexing, tabs, extra]);

  const ensureIndex = () => {
    if (!index && !indexing) setIndexing(true);
  };

  const query = q.trim();
  const results = useMemo(() => (index && query ? rankCommands(index, query).slice(0, MAX_RESULTS) : []), [index, query]);
  const recents = useMemo(() => (index && !query ? resolveRecentCommands(recent, index) : []), [index, query, recent]);
  const rows: Row[] = query ? [...results.map((c): Row => ({ kind: "cmd", c })), ...(find ? [{ kind: "find" } as Row] : [])] : recents.map((c): Row => ({ kind: "cmd", c }));
  const showList = open && (query ? !!index || !!find : recents.length > 0);
  const current = active >= 0 && active < rows.length ? active : -1;

  // Liste arama kutusunun altında, ekrandan taşmadan açılır.
  useLayoutEffect(() => {
    if (!showList) return;
    const place = () => {
      const r = boxRef.current?.getBoundingClientRect();
      if (!r) return;
      const vw = window.innerWidth;
      const width = Math.min(Math.max(r.width, 360), vw - 16);
      setPos({ left: Math.max(8, Math.min(r.left, vw - width - 8)), top: r.bottom + 4, width });
    };
    place();
    window.addEventListener("resize", place);
    return () => window.removeEventListener("resize", place);
  }, [showList]);

  // Klavyeyle seçilen satır görünür kalır.
  useEffect(() => {
    if (showList && current >= 0) document.getElementById(optId(current))?.scrollIntoView({ block: "nearest" });
  }, [showList, current]); // eslint-disable-line react-hooks/exhaustive-deps

  const reset = () => {
    setOpen(false);
    setQ("");
    setActive(-1);
    inputRef.current?.blur();
  };

  const remember = (c: CommandEntry) => {
    const next = pushRecentCommand(loadRecentCommands(app), c);
    setRecent(next);
    saveRecentCommands(app, next);
  };

  const runEntry = async (c: CommandEntry) => {
    reset();
    if (c.run) {
      remember(c);
      c.run();
      return;
    }
    // Sekme zaten açıksa beklemeden çalışır (hemen ardından basılan tuşlar komutun odağıyla yarışmasın).
    if (!isShowing(tabs, c.tabId, collapsed)) {
      if (collapsed) onExpand();
      onTabChange(c.tabId);
      await nextFrame();
    }
    const el = await locateCommand(c);
    if (el && activateCommand(el)) remember(c);
    else onMessage?.(`"${c.title}" komutu şu anda kullanılamıyor.`);
  };

  const runFind = () => {
    if (!find) return;
    const text = query;
    reset();
    find.run(text);
  };

  const runRow = (i: number) => {
    const row = rows[i];
    if (!row) return;
    if (row.kind === "cmd") void runEntry(row.c);
    else runFind();
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    e.stopPropagation();
    const n = rows.length;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setOpen(true);
      setActive(n ? (current + 1) % n : -1);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setOpen(true);
      setActive(n ? (current <= 0 ? n - 1 : current - 1) : -1);
    } else if (e.key === "Enter") {
      if (e.nativeEvent.isComposing) return;
      e.preventDefault();
      if (current >= 0) runRow(current);
      else if (findOnEnter) runFind();
    } else if (e.key === "Escape") {
      e.preventDefault();
      reset();
      onDismiss?.();
    }
  };

  const item = (i: number, key: string, icon: ReactNode, title: string, path: string, shortcut?: string) => (
    <div
      key={key}
      id={optId(i)}
      role="option"
      aria-selected={i === current}
      className={`oh-cmd-item${i === current ? " is-active" : ""}`}
      onMouseMove={() => i !== current && setActive(i)}
      onClick={() => runRow(i)}
    >
      <span className="oh-cmd-icon" aria-hidden="true">
        {icon}
      </span>
      <span className="oh-cmd-text">
        <span className="oh-cmd-title">{title}</span>
        {path && <span className="oh-cmd-path">{path}</span>}
      </span>
      {shortcut && <kbd className="oh-cmd-key">{shortcut}</kbd>}
    </div>
  );

  const cmdItem = (c: CommandEntry, i: number) =>
    item(i, `${i}:${c.id}`, c.icon ? <span className="oh-cmd-svg" dangerouslySetInnerHTML={{ __html: c.icon }} /> : FILE_ICONS[c.id], c.title, `${c.tabLabel}${c.group ? ` › ${c.group}` : ""}`, c.shortcut);

  if (!host) return null;
  return (
    <>
      {createPortal(
        <label className="oh-search oh-cmd-search" ref={boxRef}>
          <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="currentColor" strokeWidth="1.3" />
            <path d="M10 10l4.5 4.5" stroke="currentColor" strokeWidth="1.3" />
          </svg>
          <input
            ref={inputRef}
            value={q}
            placeholder={placeholder}
            aria-label="Komut ara"
            role="combobox"
            aria-autocomplete="list"
            aria-expanded={showList}
            aria-controls={listId}
            aria-activedescendant={showList && current >= 0 ? optId(current) : undefined}
            spellCheck={false}
            autoComplete="off"
            onFocus={() => {
              ensureIndex();
              setRecent(loadRecentCommands(app));
              setOpen(true);
            }}
            onChange={(e) => {
              setQ(e.target.value);
              setActive(e.target.value.trim() && !findOnEnter ? 0 : -1);
              setOpen(true);
              ensureIndex();
            }}
            onKeyDown={onKeyDown}
            onBlur={() => {
              // Sekmeler (bağlamsal sekmeler) değişebilir: dizin bir sonraki aramada tazelenir.
              setIndex(null);
              setOpen(false);
            }}
          />
        </label>,
        host,
      )}
      {showList &&
        pos &&
        createPortal(
          <div className="oh-cmd-pop" style={{ left: pos.left, top: pos.top, width: pos.width }} onMouseDown={(e) => e.preventDefault()} onContextMenu={(e) => e.preventDefault()}>
            <div className="oh-cmd-list" id={listId} role="listbox" aria-label="Arama sonuçları">
              {query ? (
                <>
                  {results.length > 0 && (
                    <div role="group" aria-label="Eylemler">
                      <div className="oh-cmd-head" aria-hidden="true">
                        Eylemler
                      </div>
                      {results.map(cmdItem)}
                    </div>
                  )}
                  {!results.length && index && <div className="oh-cmd-empty">Eşleşen komut yok.</div>}
                  {find && (
                    <div role="group" aria-label={find.heading ?? "Bul"}>
                      {find.heading ? (
                        <div className="oh-cmd-head" aria-hidden="true">
                          {find.heading}
                        </div>
                      ) : (
                        index && <div className="oh-cmd-sep" aria-hidden="true" />
                      )}
                      {item(results.length, "find", <Search20Regular />, find.label(query), find.path ?? "", find.shortcut)}
                    </div>
                  )}
                </>
              ) : (
                <div role="group" aria-label="Son kullanılanlar">
                  <div className="oh-cmd-head" aria-hidden="true">
                    Son kullanılanlar
                  </div>
                  {recents.map(cmdItem)}
                </div>
              )}
            </div>
          </div>,
          document.body,
        )}
      {indexing &&
        createPortal(
          <div ref={hiddenRef} className="oh-cmd-index" aria-hidden="true" inert>
            {tabs
              .filter((t) => !t.hidden)
              .map((t) => (
                <div key={t.id} data-cmd-tab={t.id}>
                  {t.content}
                </div>
              ))}
          </div>,
          document.body,
        )}
    </>
  );
}
