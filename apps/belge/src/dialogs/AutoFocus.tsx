// Sekmeli iletişim kutularında ilk alan odaklansın (çerçeve ilk sekme düğmesini odaklar).
import { useEffect, useRef, type ReactNode } from "react";

export function AutoFocus({ children, deps = [] }: { children?: ReactNode; deps?: unknown[] }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const id = requestAnimationFrame(() => {
      const root = ref.current?.closest(".oh-dialog-body") ?? ref.current;
      const el = root?.querySelector<HTMLElement>('input:not([type="checkbox"]):not([type="radio"]):not(:disabled), select:not([size]), textarea');
      el?.focus();
      if (el instanceof HTMLInputElement) el.select();
    });
    return () => cancelAnimationFrame(id);
  }, deps); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div ref={ref} style={{ display: "contents" }}>
      {children}
    </div>
  );
}
