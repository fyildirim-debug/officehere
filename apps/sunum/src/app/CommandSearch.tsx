/**
 * Başlık çubuğundaki "Ara (Alt+Q)": şerit komutlarından (etiket + ipucu +
 * açıklama) bir dizin oluşturur, sonuçları açılır listede gösterir ve
 * seçilen komutu çalıştırır. Listenin sonunda sunuda metin arayan "Bul"
 * satırı bulunur.
 *
 * Dizin, tüm sekmelerin içeriği görünmeyen bir kapta bir kez çizilerek
 * çıkarılır; komut çalıştırılırken ilgili sekmeye geçilip aynı düğmeye
 * (daraltılmış grupta grup açılarak) tıklanır.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Popover, type RibbonTabDef } from "@officehere/ui";
import type { RibbonCtx } from "../ribbon/types";
import { rankCommands, type CommandEntry } from "./commandIndex";

const MAX_RESULTS = 8;

function scrape(root: HTMLElement, tabs: RibbonTabDef[]): CommandEntry[] {
  const out: CommandEntry[] = [];
  root.querySelectorAll<HTMLElement>("[data-cmd-tab]").forEach((pane) => {
    const tabId = pane.dataset.cmdTab!;
    const tabLabel = tabs.find((t) => t.id === tabId)?.label ?? "";
    const counts = new Map<string, number>();
    pane.querySelectorAll<HTMLElement>("[data-tip]").forEach((el) => {
      const title = el.dataset.tip!;
      const group = el.closest<HTMLElement>(".oh-ribbon-group")?.getAttribute("aria-label") ?? "";
      if (el.classList.contains("oh-ribbon-launcher")) return;
      const key = `${group}\u0000${title}`;
      const nth = counts.get(key) ?? 0;
      counts.set(key, nth + 1);
      out.push({ id: `${tabId}:${key}:${nth}`, title, desc: el.dataset.tipDesc, shortcut: el.dataset.tipKey, tabId, tabLabel, group, nth });
    });
  });
  // Aynı adlı komutlardan yalnız ilki (ör. birden çok sekmedeki "Yazı Tipi").
  const seen = new Set<string>();
  return out.filter((c) => {
    const k = `${c.title}|${c.desc ?? ""}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

const nextFrame = () => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())));

/** Canlı şeritte komutun öğesini bulur (gerekirse daraltılmış grubu açar). */
async function locate(c: CommandEntry): Promise<HTMLElement | null> {
  const body = document.querySelector(".oh-ribbon-body");
  if (!body) return null;
  const sel = `[data-tip="${CSS.escape(c.title)}"]`;
  const group = Array.from(body.querySelectorAll<HTMLElement>(".oh-ribbon-group[role=group]")).find((g) => g.getAttribute("aria-label") === c.group);
  if (group) return group.querySelectorAll<HTMLElement>(sel)[c.nth] ?? null;
  const collapsed = Array.from(body.querySelectorAll<HTMLElement>(".oh-ribbon-group.is-collapsed")).find((g) => g.textContent?.trim().startsWith(c.group));
  const btn = collapsed?.querySelector("button");
  if (btn) {
    btn.click();
    await nextFrame();
    const pop = Array.from(document.querySelectorAll<HTMLElement>(".oh-ribbon-group-popup")).pop();
    return pop?.querySelectorAll<HTMLElement>(sel)[c.nth] ?? null;
  }
  return body.querySelector<HTMLElement>(sel);
}

export function CommandSearch({ ctx, tabs, collapsed, onExpand }: { ctx: RibbonCtx; tabs: RibbonTabDef[]; collapsed: boolean; onExpand: () => void }) {
  const [host, setHost] = useState<HTMLElement | null>(null);
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(0);
  const [index, setIndex] = useState<CommandEntry[] | null>(null);
  const [indexing, setIndexing] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const boxRef = useRef<HTMLLabelElement>(null);
  const hiddenRef = useRef<HTMLDivElement>(null);

  // Başlık çubuğundaki arama yuvasına yerleş.
  useLayoutEffect(() => {
    setHost(document.querySelector<HTMLElement>(".oh-titlebar-search"));
  }, []);

  // Alt+Q: aramaya odaklan.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.altKey && (e.key === "q" || e.key === "Q" || e.code === "KeyQ")) {
        e.preventDefault();
        inputRef.current?.focus();
        inputRef.current?.select();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Dizin: gizli kapta çizilen sekmelerden çıkarılır.
  useLayoutEffect(() => {
    if (!indexing || !hiddenRef.current) return;
    setIndex(scrape(hiddenRef.current, tabs));
    setIndexing(false);
  }, [indexing, tabs]);

  const ensureIndex = () => {
    if (!index && !indexing) setIndexing(true);
  };

  const results = useMemo(() => (index && q.trim() ? rankCommands(index, q).slice(0, MAX_RESULTS) : []), [index, q]);
  const rows = q.trim() ? results.length + 1 : 0;

  const close = useCallback(() => setOpen(false), []);

  const runEntry = async (c: CommandEntry) => {
    setOpen(false);
    setQ("");
    inputRef.current?.blur();
    if (collapsed) onExpand();
    ctx.actions.store.set({ ribbonTab: c.tabId });
    await nextFrame();
    const el = await locate(c);
    if (!el) {
      ctx.actions.message(`"${c.title}" komutu şu anda kullanılamıyor.`, "info");
      return;
    }
    if (el instanceof HTMLInputElement || el instanceof HTMLSelectElement) {
      el.focus();
      return;
    }
    const target = el.matches("button") ? el : (el.querySelector<HTMLElement>("button, input") ?? el);
    if ((target as HTMLButtonElement).disabled) {
      ctx.actions.message(`"${c.title}" komutu şu anda kullanılamıyor.`, "info");
      return;
    }
    target.click();
  };

  const runRow = (i: number) => {
    if (i < results.length) void runEntry(results[i]);
    else {
      const text = q.trim();
      setOpen(false);
      setQ("");
      inputRef.current?.blur();
      ctx.cmd.find(text);
    }
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    e.stopPropagation();
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setOpen(true);
      setActive((x) => (rows ? (x + 1) % rows : 0));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActive((x) => (rows ? (x - 1 + rows) % rows : 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (rows) runRow(Math.min(active, rows - 1));
    } else if (e.key === "Escape") {
      setQ("");
      setOpen(false);
      inputRef.current?.blur();
    }
  };

  if (!host) return null;
  return (
    <>
      {createPortal(
        <label className="oh-search cmd-search" ref={boxRef}>
          <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="6.5" cy="6.5" r="4.5" fill="none" stroke="currentColor" strokeWidth="1.3" />
            <path d="M10 10l4.5 4.5" stroke="currentColor" strokeWidth="1.3" />
          </svg>
          <input
            ref={inputRef}
            value={q}
            placeholder="Ara (Alt+Q)"
            aria-label="Komut ara"
            role="combobox"
            aria-expanded={open && rows > 0}
            aria-controls="cmd-search-list"
            onFocus={() => {
              ensureIndex();
              if (q.trim()) setOpen(true);
            }}
            onChange={(e) => {
              setQ(e.target.value);
              setActive(0);
              setOpen(!!e.target.value.trim());
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
      <Popover open={open && rows > 0} onClose={close} anchor={boxRef} className="cmd-search-pop">
        <div className="oh-menu cmd-search-list" id="cmd-search-list" role="listbox" aria-label="Arama sonuçları">
          {results.length > 0 && <div className="cmd-search-head">Eylemler</div>}
          {results.map((c, i) => (
            <div key={c.id} role="option" aria-selected={i === active} className={`cmd-search-item${i === active ? " is-active" : ""}`} onMouseEnter={() => setActive(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => runRow(i)}>
              <span className="cmd-search-title">{c.title}</span>
              <span className="cmd-search-path">
                {c.tabLabel}
                {c.group ? ` › ${c.group}` : ""}
                {c.shortcut ? ` (${c.shortcut})` : ""}
              </span>
            </div>
          ))}
          {!results.length && index && <div className="cmd-search-empty">Eşleşen komut yok.</div>}
          <div className="cmd-search-head">Sunuda bul</div>
          <div role="option" aria-selected={active === results.length} className={`cmd-search-item${active === results.length ? " is-active" : ""}`} onMouseEnter={() => setActive(results.length)} onMouseDown={(e) => e.preventDefault()} onClick={() => runRow(results.length)}>
            <span className="cmd-search-title">“{q.trim()}” için Bul</span>
            <span className="cmd-search-path">Giriş › Düzenleme (Ctrl+F)</span>
          </div>
        </div>
      </Popover>
      {indexing &&
        createPortal(
          <div ref={hiddenRef} className="cmd-index" aria-hidden="true">
            {tabs
              .filter((t) => !t.hidden)
              .map((t) => (
                <div key={t.id} data-cmd-tab={t.id} className="cmd-index-tab">
                  {t.content}
                </div>
              ))}
          </div>,
          document.body,
        )}
    </>
  );
}
