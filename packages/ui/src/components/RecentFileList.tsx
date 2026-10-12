import { useEffect, useMemo, useRef, useState, type KeyboardEvent, type MouseEvent } from "react";
import { Delete20Regular, FolderOpen20Regular, Open20Regular, Pin16Filled, Pin16Regular, Pin20Regular, PinOff20Regular, Search16Regular } from "@fluentui/react-icons";
import {
  canShowInFolder,
  filterRecent,
  folderOf,
  formatDateTime,
  formatRelativeTime,
  getRecentFiles,
  getRecentFileStatus,
  isDesktop,
  onRecentFilesChange,
  orderRecent,
  pinRecentFile,
  recentKey,
  removeRecentFile,
  shortenPath,
  showInFolder,
  type RecentFile,
} from "@officehere/core";
import { AppLogo } from "./AppShell";
import { Dialog, DialogButton } from "./Dialog";
import { Menu, MenuItem, MenuSeparator, useContextMenu } from "./Menu";

export interface RecentFileListProps {
  files: RecentFile[];
  onOpen: (f: RecentFile) => void;
  emptyText?: string;
  /** Üstte "Son dosyalarda ara" kutusu (varsayılan: açık). */
  searchable?: boolean;
}

/**
 * Son kullanılan dosyalar: sabitlenmiş ve son gruplar, arama, sağ tık menüsü, sabitleme düğmesi.
 * Sabitleme/kaldırma ortak depoya yazılır; liste, üst bileşenin elindeki kopya eski kalsa bile
 * depodaki güncel durumla eşlenir (sayfa değişip liste yeniden kurulduğunda da).
 */
export function RecentFileList({ files, onOpen, emptyText = "Son kullanılan dosya yok.", searchable = true }: RecentFileListProps) {
  const [query, setQuery] = useState("");
  const [store, setStore] = useState<{ key: string; map: Map<string, RecentFile> } | null>(null);
  const [missing, setMissing] = useState<ReadonlySet<string>>(new Set());
  const [ask, setAsk] = useState<RecentFile | null>(null);
  const ctx = useContextMenu();
  const listRef = useRef<HTMLDivElement>(null);
  const searchRef = useRef<HTMLInputElement>(null);

  const filesKey = files.map((f) => `${recentKey(f)}#${f.pinned ? 1 : 0}`).join("\n");
  const keyRef = useRef(filesKey);
  keyRef.current = filesKey;

  useEffect(() => {
    let live = true;
    const load = () => {
      const key = keyRef.current;
      getRecentFiles()
        .then((list) => live && setStore({ key, map: new Map(list.map((f) => [recentKey(f), f])) }))
        .catch(() => {});
    };
    load();
    const off = onRecentFilesChange(load);
    return () => {
      live = false;
      off();
    };
  }, [filesKey]);

  const items = useMemo(() => {
    const map = store?.key === filesKey ? store.map : null;
    const base = map
      ? files.flatMap((f) => {
          const cur = map.get(recentKey(f));
          return cur ? [{ ...f, pinned: cur.pinned === true }] : [];
        })
      : files;
    return orderRecent(base);
  }, [files, filesKey, store]);

  // Taşınmış/silinmiş dosyalar (masaüstü); pencereye dönülünce yeniden denetlenir.
  const pathsKey = items.map((f) => f.path).join("\n");
  useEffect(() => {
    if (!isDesktop() || !pathsKey) return;
    let live = true;
    const check = () => {
      getRecentFileStatus(pathsKey.split("\n"))
        .then((st) => live && setMissing(new Set(Object.keys(st).filter((p) => st[p] === false))))
        .catch(() => {});
    };
    check();
    window.addEventListener("focus", check);
    return () => {
      live = false;
      window.removeEventListener("focus", check);
    };
  }, [pathsKey]);

  if (items.length === 0) return <p className="oh-bs-empty">{emptyText}</p>;

  const open = async (f: RecentFile) => {
    if (missing.has(f.path)) {
      setAsk(f);
      return;
    }
    if (isDesktop()) {
      const st = await getRecentFileStatus([f.path]);
      if (st[f.path] === false) {
        setMissing((prev) => new Set(prev).add(f.path));
        setAsk(f);
        return;
      }
    }
    onOpen(f);
  };
  const pin = (f: RecentFile, pinned: boolean) => void pinRecentFile(f, pinned);
  const remove = (f: RecentFile) => void removeRecentFile(f);

  const menuFor = (f: RecentFile) => (
    <Menu>
      <MenuItem icon={<Open20Regular />} label="Aç" onClick={() => void open(f)} />
      <MenuItem icon={f.pinned ? <PinOff20Regular /> : <Pin20Regular />} label={f.pinned ? "Sabitlemeyi kaldır" : "Sabitle"} onClick={() => pin(f, !f.pinned)} />
      {canShowInFolder() && <MenuItem icon={<FolderOpen20Regular />} label="Klasörde göster" disabled={missing.has(f.path)} onClick={() => void showInFolder(f.path)} />}
      <MenuSeparator />
      <MenuItem icon={<Delete20Regular />} label="Listeden kaldır" onClick={() => remove(f)} />
    </Menu>
  );

  const onContextMenu = (e: MouseEvent<HTMLElement>, f: RecentFile) => {
    // Klavyeden (Menü tuşu, Shift+F10) açılınca menü satırın altında belirir.
    const r = e.currentTarget.getBoundingClientRect();
    const inside = e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
    ctx.open({ clientX: inside ? e.clientX : r.left + 48, clientY: inside ? e.clientY : r.bottom - 4, preventDefault: () => e.preventDefault() }, menuFor(f));
  };

  const moveFocus = (e: KeyboardEvent<HTMLElement>) => {
    const all = [...(listRef.current?.querySelectorAll<HTMLButtonElement>(".oh-recent-open") ?? [])];
    const i = all.indexOf(e.currentTarget as HTMLButtonElement);
    let next: HTMLButtonElement | undefined;
    if (e.key === "ArrowDown") next = all[i + 1];
    else if (e.key === "ArrowUp") next = i <= 0 ? undefined : all[i - 1];
    else if (e.key === "Home") next = all[0];
    else if (e.key === "End") next = all[all.length - 1];
    else return;
    e.preventDefault();
    if (next) next.focus();
    else if (e.key === "ArrowUp" && searchable) searchRef.current?.focus();
  };

  const now = Date.now();
  const row = (f: RecentFile) => {
    const isMissing = missing.has(f.path);
    const folder = f.path === f.name ? "Bu cihaz" : shortenPath(folderOf(f.path), 72);
    return (
      <div
        key={recentKey(f)}
        role="listitem"
        className={`oh-recent-row${f.pinned ? " is-pinned" : ""}${isMissing ? " is-missing" : ""}`}
        onContextMenu={(e) => onContextMenu(e, f)}
      >
        <button type="button" className="oh-recent-open" onClick={() => void open(f)} onKeyDown={moveFocus} title={f.path}>
          <span className="oh-recent-icon">
            <AppLogo app={f.app} size={32} />
          </span>
          <span className="oh-recent-text">
            <span className="oh-recent-name">
              <span className="oh-recent-filename">{f.name}</span>
              {isMissing && <span className="oh-recent-badge">Bulunamadı</span>}
            </span>
            <span className="oh-recent-path">{folder}</span>
          </span>
          <span className="oh-recent-time" title={formatDateTime(f.openedAt)}>
            {formatRelativeTime(f.openedAt, now)}
          </span>
        </button>
        <button
          type="button"
          className="oh-recent-pin"
          aria-label={f.pinned ? `${f.name}: sabitlemeyi kaldır` : `${f.name}: sabitle`}
          aria-pressed={f.pinned === true}
          title={f.pinned ? "Sabitlemeyi kaldır" : "Listede sabitle"}
          onClick={() => pin(f, !f.pinned)}
        >
          {f.pinned ? <Pin16Filled /> : <Pin16Regular />}
        </button>
      </div>
    );
  };

  const shown = filterRecent(items, query);
  const pinned = shown.filter((f) => f.pinned);
  const rest = shown.filter((f) => !f.pinned);

  return (
    <div className="oh-recent" ref={listRef}>
      {searchable && (
        <div className="oh-recent-search">
          <Search16Regular />
          <input
            ref={searchRef}
            type="search"
            placeholder="Son dosyalarda ara"
            aria-label="Son dosyalarda ara"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Escape" && query) {
                e.stopPropagation();
                setQuery("");
              } else if (e.key === "ArrowDown") {
                e.preventDefault();
                listRef.current?.querySelector<HTMLButtonElement>(".oh-recent-open")?.focus();
              }
            }}
          />
        </div>
      )}
      {shown.length === 0 && <p className="oh-bs-empty">“{query.trim()}” ile eşleşen dosya yok.</p>}
      {pinned.length > 0 && (
        <>
          <h3 className="oh-recent-group">Sabitlenmiş</h3>
          <div role="list" aria-label="Sabitlenmiş" className="oh-recent-list">
            {pinned.map(row)}
          </div>
        </>
      )}
      {rest.length > 0 && (
        <>
          {pinned.length > 0 && <h3 className="oh-recent-group">Son</h3>}
          <div role="list" aria-label="Son" className="oh-recent-list">
            {rest.map(row)}
          </div>
        </>
      )}
      {ctx.element}
      <Dialog
        open={ask !== null}
        title="OfficeHere"
        onClose={() => setAsk(null)}
        footer={
          <>
            <DialogButton
              primary
              onClick={() => {
                if (ask) remove(ask);
                setAsk(null);
              }}
            >
              Kaldır
            </DialogButton>
            <DialogButton onClick={() => setAsk(null)}>İptal</DialogButton>
          </>
        }
      >
        <p className="oh-dialog-message">Dosya bulunamadı. Listeden kaldırılsın mı?</p>
        {ask && <p className="oh-recent-dialog-path">{ask.path}</p>}
      </Dialog>
    </div>
  );
}
