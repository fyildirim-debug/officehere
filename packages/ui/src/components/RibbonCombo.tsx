import { useEffect, useId, useRef, useState, type CSSProperties } from "react";
import { Popover } from "./Popover";
import { tipProps } from "./Tooltip";
import { nextEnabled, scrollToReveal } from "./menuNav";
import { setKeyboardMode } from "./menuKeyboard";

/* ------------------------------------------------------------------ */
/* Açılan kutu (yazı tipi, boyut)                                       */
/* ------------------------------------------------------------------ */

export interface ComboOption {
  value: string;
  label?: string;
  /** Seçenek satırına uygulanacak stil (yazı tipi önizlemesi). */
  style?: CSSProperties;
  group?: string;
}

export interface RibbonComboProps {
  value: string;
  options: (string | ComboOption)[];
  onCommit: (value: string) => void;
  width?: number;
  editable?: boolean;
  tip?: string;
  shortcut?: string;
  disabled?: boolean;
  placeholder?: string;
  /** Liste yüksekliği (px). */
  listHeight?: number;
}

/**
 * Yazılabilir açılan kutu. Klavye: Alt+↓ ya da ↓ listeyi açar; liste açıkken ↑/↓ etkin seçeneği
 * taşır (odak kutuda kalır, aria-activedescendant), Enter seçer, Esc kapatıp önceki değeri korur.
 */
export function RibbonCombo({ value, options, onCommit, width = 120, editable = true, tip, shortcut, disabled, placeholder, listHeight = 320 }: RibbonComboProps) {
  const [text, setText] = useState(value);
  const [open, setOpen] = useState(false);
  // Klavyeyle vurgulanan seçenek (-1: yok; fareyle açılınca vurgu olmaz).
  const [active, setActive] = useState(-1);
  const ref = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const pointer = useRef<{ x: number; y: number } | null>(null);
  const opts = options.map((o) => (typeof o === "string" ? { value: o } : o));
  const id = useId();
  const optId = (i: number) => `${id}-opt-${i}`;

  useEffect(() => setText(value), [value]);

  useEffect(() => {
    if (!open) return;
    // Seçili öğeyi ortala; scrollIntoView yerine yalnız listeyi kaydır ki üst öğeler (sayfa) kaymasın.
    const list = listRef.current;
    const sel = list?.querySelector(".is-selected") as HTMLElement | null;
    if (list && sel) list.scrollTop = sel.offsetTop - list.clientHeight / 2 + sel.offsetHeight / 2;
  }, [open]);

  useEffect(() => {
    // Etkin seçeneği görünür tut (yalnız liste kayar).
    const list = listRef.current;
    const el = active >= 0 ? (document.getElementById(optId(active)) as HTMLElement | null) : null;
    if (list && el) list.scrollTop = scrollToReveal(list.scrollTop, list.clientHeight, el.offsetTop, el.offsetHeight);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active]);

  const setOpenList = (o: boolean) => {
    setOpen(o);
    setActive(-1);
  };

  const commit = (v: string) => {
    setOpenList(false);
    if (v.trim() === "") {
      setText(value);
      return;
    }
    onCommit(v);
  };

  /** ↑/↓: etkin seçenek yoksa geçerli değerden başlar. */
  const move = (delta: 1 | -1) => {
    const sel = opts.findIndex((o) => o.value === value);
    const from = active >= 0 ? active : sel;
    const next = from < 0 ? nextEnabled(opts.length, -1, delta, undefined, false) : nextEnabled(opts.length, from, delta, undefined, false);
    if (next >= 0) setActive(next);
    else if (from >= 0 && active < 0) setActive(from);
    if (listRef.current) setKeyboardMode(listRef.current, true);
  };

  let lastGroup: string | undefined;
  return (
    <>
      <div ref={ref} className={`oh-combo${open ? " is-open" : ""}${disabled ? " is-disabled" : ""}`} style={{ width }} {...tipProps(tip, shortcut)}>
        <input
          className="oh-combo-input"
          value={text}
          disabled={disabled}
          readOnly={!editable}
          placeholder={placeholder}
          aria-label={tip}
          role="combobox"
          aria-expanded={open}
          aria-controls={open ? `${id}-list` : undefined}
          aria-autocomplete="list"
          aria-activedescendant={open && active >= 0 ? optId(active) : undefined}
          onChange={(e) => {
            setText(e.target.value);
            setActive(-1);
          }}
          onFocus={(e) => e.currentTarget.select()}
          onMouseDown={!editable ? (e) => (e.preventDefault(), setOpenList(!open)) : undefined}
          onKeyDown={(e) => {
            if (e.nativeEvent.isComposing) return;
            if (e.key === "Enter") {
              e.preventDefault();
              if (open && active >= 0 && opts[active]) {
                setText(opts[active].value);
                commit(opts[active].value);
              } else commit(text);
            } else if (e.key === "Escape") {
              setText(value);
              setOpenList(false);
            } else if (e.key === "ArrowDown" && !open) {
              e.preventDefault();
              setOpenList(true);
            } else if ((e.key === "ArrowDown" || e.key === "ArrowUp") && open) {
              e.preventDefault();
              // Alt+↑ listeyi kapatır; ↑/↓ etkin seçeneği taşır.
              if (e.altKey) {
                if (e.key === "ArrowUp") setOpenList(false);
              } else move(e.key === "ArrowDown" ? 1 : -1);
            } else if (open && (e.key === "PageDown" || e.key === "PageUp")) {
              e.preventDefault();
              const step = e.key === "PageDown" ? 8 : -8;
              const from = active >= 0 ? active : opts.findIndex((o) => o.value === value);
              setActive(Math.max(0, Math.min(opts.length - 1, from + step)));
              if (listRef.current) setKeyboardMode(listRef.current, true);
            } else if (e.key === "Tab" && open) {
              setOpenList(false);
            }
          }}
          onBlur={() => {
            if (!open && text !== value) commit(text);
          }}
        />
        <button
          type="button"
          className="oh-combo-arrow"
          tabIndex={-1}
          disabled={disabled}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setOpenList(!open)}
          aria-label="Listeyi aç"
        >
          <svg className="oh-chevron" width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
            <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
          </svg>
        </button>
      </div>
      <Popover
        open={open}
        onClose={(reason) => {
          // Esc önceki değeri korur (yazılan metin geri alınır).
          if (reason === "escape") setText(value);
          setOpenList(false);
        }}
        anchor={ref}
      >
        <div
          ref={listRef}
          id={`${id}-list`}
          className="oh-combo-list"
          style={{ minWidth: width, maxHeight: listHeight }}
          role="listbox"
          aria-label={tip}
          onMouseMove={(e) => {
            // Yalnız gerçek fare hareketi klavye kipini kaldırır (liste kayınca gelen olay sayılmaz).
            const p = pointer.current;
            pointer.current = { x: e.clientX, y: e.clientY };
            if (p && (p.x !== e.clientX || p.y !== e.clientY)) setKeyboardMode(e.currentTarget, false);
          }}
        >
          {opts.map((o, i) => {
            const header = o.group && o.group !== lastGroup ? o.group : null;
            lastGroup = o.group;
            return (
              <div key={`${o.group ?? ""}:${o.value}`} role="presentation">
                {header && (
                  <div className="oh-menu-header" role="presentation">
                    {header}
                  </div>
                )}
                <button
                  type="button"
                  id={optId(i)}
                  role="option"
                  tabIndex={-1}
                  aria-selected={o.value === value}
                  className={`oh-combo-option${o.value === value ? " is-selected" : ""}${i === active ? " is-active" : ""}`}
                  style={o.style}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    setText(o.value);
                    commit(o.value);
                  }}
                >
                  {o.label ?? o.value}
                </button>
              </div>
            );
          })}
        </div>
      </Popover>
    </>
  );
}
