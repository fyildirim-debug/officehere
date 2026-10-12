/** Sayfa sekmeleri: yeni sayfa, yeniden adlandırma, sürükleyerek sıralama, kaydırma okları. */
import { useEffect, useRef, useState } from "react";
import { tipProps } from "@officehere/ui";
import { contrastText } from "@officehere/core";
import type { Controller } from "../controller";

export function SheetTabs({ ctl, onContextMenu }: { ctl: Controller; onContextMenu: (e: React.MouseEvent, index: number) => void }) {
  const wb = ctl.wb;
  const [renaming, setRenaming] = useState<number | null>(null);
  const [name, setName] = useState("");
  const [dragging, setDragging] = useState<{ from: number; over: number | null } | null>(null);
  const stripRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    // Etkin sekmeyi görünür yap
    const el = stripRef.current?.querySelector<HTMLElement>(".tb-tab.is-active");
    el?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [wb.activeIndex]);

  // Geçersiz ad uyarısı açıkken odak kaybı yeniden doğrulamaz (uyarı iki kez açılmasın)
  const alerting = useRef(false);
  const inputRef = useRef<HTMLInputElement>(null);
  // Uyarı kapanınca odak ad kutusuna döner (Excel gibi düzeltmeye devam edilir)
  useEffect(() => {
    if (!alerting.current || renaming === null) return;
    const h = window.setInterval(() => {
      if (document.querySelector(".oh-dialog")) return;
      window.clearInterval(h);
      inputRef.current?.focus();
    }, 50);
    return () => window.clearInterval(h);
  });
  const commitRename = (fromBlur = false) => {
    if (renaming === null || alerting.current) return;
    const err = wb.isValidSheetName(name, wb.sheets[renaming]);
    if (err) {
      // Odak kaybında geçersiz ad sessizce bırakılır (uyarı yalnız Enter'da; iki kez açılmaz)
      if (fromBlur) {
        setRenaming(null);
        return;
      }
      alerting.current = true;
      ctl.alert(err);
      return;
    }
    if (name.trim() !== wb.sheets[renaming].name) ctl.ops.renameSheet(renaming, name);
    setRenaming(null);
    ctl.focusGrid();
  };

  const startRename = (i: number) => {
    setRenaming(i);
    setName(wb.sheets[i].name);
  };

  // App bağlam menüsünden yeniden adlandırma isteği
  useEffect(() => {
    const h = (e: Event) => startRename((e as CustomEvent<number>).detail);
    window.addEventListener("tb-rename-sheet", h);
    return () => window.removeEventListener("tb-rename-sheet", h);
  });

  return (
    <div className="tb-tabs">
      <div className="tb-tabs-nav">
        <button type="button" className="tb-tabs-arrow" onClick={() => stripRef.current?.scrollBy({ left: -120, behavior: "smooth" })} {...tipProps("Önceki sayfaya kaydır", undefined, "Ctrl ile ilk sayfaya kaydırın")}>
          <svg width="8" height="10" viewBox="0 0 8 10" aria-hidden="true">
            <path d="M6 1L2 5l4 4" fill="none" stroke="currentColor" strokeWidth="1.3" />
          </svg>
        </button>
        <button type="button" className="tb-tabs-arrow" onClick={() => stripRef.current?.scrollBy({ left: 120, behavior: "smooth" })} {...tipProps("Sonraki sayfaya kaydır", undefined, "Ctrl ile son sayfaya kaydırın")}>
          <svg width="8" height="10" viewBox="0 0 8 10" aria-hidden="true">
            <path d="M2 1l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.3" />
          </svg>
        </button>
      </div>
      <div className="tb-tabs-strip" ref={stripRef} role="tablist">
        {wb.sheets.map((s, i) => {
          if (s.hidden) return null;
          const active = i === wb.activeIndex;
          const style: React.CSSProperties = {};
          if (s.tabColor && !active) {
            style.background = s.tabColor;
            style.color = contrastText(s.tabColor);
          }
          return (
            <div
              key={s.id}
              role="tab"
              aria-selected={active}
              className={`tb-tab${active ? " is-active" : ""}${dragging && dragging.over === i ? " is-drop" : ""}`}
              style={style}
              onPointerDown={(e) => {
                if (e.button !== 0 || renaming === i) return;
                if (i !== wb.activeIndex) ctl.activateSheet(i);
                setDragging({ from: i, over: null });
                (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
              }}
              onPointerMove={(e) => {
                if (!dragging) return;
                const el = document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>(".tb-tab");
                const idx = el ? Number(el.dataset.index) : null;
                if (idx !== dragging.over) setDragging({ ...dragging, over: idx });
              }}
              onPointerUp={() => {
                if (dragging && dragging.over !== null && dragging.over !== dragging.from) ctl.ops.moveSheet(dragging.from, dragging.over);
                setDragging(null);
                if (renaming === null) ctl.focusGrid();
              }}
              data-index={i}
              onDoubleClick={() => startRename(i)}
              onContextMenu={(e) => {
                e.preventDefault();
                if (i !== wb.activeIndex) ctl.activateSheet(i);
                onContextMenu(e, i);
              }}
            >
              {renaming === i ? (
                <input
                  ref={inputRef}
                  className="tb-tab-input"
                  autoFocus
                  value={name}
                  maxLength={31}
                  onChange={(e) => setName(e.target.value)}
                  onFocus={(e) => {
                    if (!document.querySelector(".oh-dialog")) alerting.current = false;
                    e.currentTarget.select();
                  }}
                  onBlur={() => commitRename(true)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      // Enter'ın tuş olayı açılan uyarının Tamam düğmesini tetiklemesin
                      e.preventDefault();
                      commitRename();
                    }
                    if (e.key === "Escape") {
                      alerting.current = false;
                      setRenaming(null);
                      ctl.focusGrid();
                    }
                    e.stopPropagation();
                  }}
                  onPointerDown={(e) => e.stopPropagation()}
                />
              ) : (
                <span>{s.name}</span>
              )}
              {active && s.tabColor && <span className="tb-tab-color" style={{ background: s.tabColor }} />}
            </div>
          );
        })}
      </div>
      <button type="button" className="tb-tab-add" onClick={() => ctl.ops.addSheet(wb.activeIndex + 1)} aria-label="Yeni sayfa" {...tipProps("Yeni sayfa", "Shift+F11")}>
        <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
          <circle cx="7" cy="7" r="6.2" fill="none" stroke="currentColor" strokeWidth="1" />
          <path d="M7 4v6M4 7h6" stroke="currentColor" strokeWidth="1.2" />
        </svg>
      </button>
    </div>
  );
}
