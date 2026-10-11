import { useState } from "react";
import { useMenuClose } from "./Menu";

/** "Tablo Ekle" ızgarası: fareyle satır × sütun seçilir. */
export function TableGridPicker({ rows = 8, cols = 10, onPick }: { rows?: number; cols?: number; onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState<{ r: number; c: number }>({ r: 0, c: 0 });
  const close = useMenuClose();
  return (
    <div className="oh-tablegrid">
      <div className="oh-menu-header">{hover.r > 0 ? `${hover.c}x${hover.r} Tablo` : "Tablo Ekle"}</div>
      <div className="oh-tablegrid-cells" style={{ gridTemplateColumns: `repeat(${cols}, 16px)` }} onMouseLeave={() => setHover({ r: 0, c: 0 })}>
        {Array.from({ length: rows * cols }, (_, i) => {
          const r = Math.floor(i / cols) + 1;
          const c = (i % cols) + 1;
          const on = r <= hover.r && c <= hover.c;
          return (
            <button
              key={i}
              type="button"
              className={`oh-tablegrid-cell${on ? " is-on" : ""}`}
              onMouseEnter={() => setHover({ r, c })}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onPick(r, c);
                close();
              }}
              aria-label={`${c}x${r} tablo`}
            />
          );
        })}
      </div>
    </div>
  );
}
