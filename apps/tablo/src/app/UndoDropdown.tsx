/** Hızlı Erişim'de Geri Al / Yinele açılır listesi: birden çok adımı tek seferde geri alır ya da yineler. */
import { useRef, useState } from "react";
import { Popover, tipProps } from "@officehere/ui";
import type { Controller } from "../controller";
import * as E from "../actions/edit";

export function UndoDropdown({ ctl, redo = false }: { ctl: Controller; redo?: boolean }) {
  const [open, setOpen] = useState(false);
  const [hover, setHover] = useState(0);
  const ref = useRef<HTMLButtonElement>(null);
  const h = ctl.ops.history;
  const items = open ? (redo ? h.recentRedo(100) : h.recent(100)) : [];
  const enabled = redo ? h.canRedo : h.canUndo;
  const run = (n: number) => {
    setOpen(false);
    for (let i = 0; i < n; i++) redo ? E.redo(ctl) : E.undo(ctl);
    ctl.focusGrid();
  };
  const word = redo ? "Yinele" : "Geri Al";
  return (
    <>
      <button
        ref={ref}
        type="button"
        className="oh-qat-btn tb-qat-drop"
        disabled={!enabled}
        aria-label={`${word} listesi`}
        aria-haspopup="listbox"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => {
          setHover(0);
          setOpen(!open);
        }}
        {...tipProps(`${word} listesi`)}
      >
        <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
          <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.2" />
        </svg>
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={ref}>
        <div className="oh-menu tb-undo-list" role="listbox" aria-label={`${word} listesi`} onMouseLeave={() => setHover(0)}>
          <div className="tb-undo-items">
            {items.map((label, i) => (
              <div key={i} role="option" aria-selected={i <= hover} className={`tb-undo-item${i <= hover ? " is-on" : ""}`} onMouseEnter={() => setHover(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => run(i + 1)}>
                {label}
              </div>
            ))}
          </div>
          <div className="tb-undo-foot">
            {word} {hover + 1} Eylem
          </div>
        </div>
      </Popover>
    </>
  );
}
