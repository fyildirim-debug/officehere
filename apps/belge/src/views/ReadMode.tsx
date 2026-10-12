// Okuma Modu: metin ekran sütunlarına akar, sayfalar yatay çevrilir.
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { ChevronLeft24Regular, ChevronRight24Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, RibbonMenuButton } from "@officehere/ui";

interface Props {
  html: string;
  title: string;
  onExit: () => void;
  onFind: () => void;
}

type Bg = "none" | "sepia" | "inverse";

export function ReadMode({ html, title, onExit, onFind }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [page, setPage] = useState(0);
  const [pages, setPages] = useState(1);
  const [bg, setBg] = useState<Bg>("none");
  const [width, setWidth] = useState<"narrow" | "default" | "wide">("default");

  const measure = useCallback(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.clientWidth;
    const n = Math.max(1, Math.round(el.scrollWidth / w));
    setPages(n);
    setPage((p) => Math.min(p, n - 1));
  }, []);

  useLayoutEffect(measure, [measure, html, width]);
  useEffect(() => {
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [measure]);

  useEffect(() => {
    const el = ref.current;
    if (el) el.scrollTo({ left: page * el.clientWidth, behavior: "smooth" });
  }, [page]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onExit();
      else if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") setPage((p) => Math.min(pages - 1, p + 1));
      else if (e.key === "ArrowLeft" || e.key === "PageUp") setPage((p) => Math.max(0, p - 1));
      else return;
      e.preventDefault();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [pages, onExit]);

  return (
    <div className={`bl-read is-${bg}`}>
      <div className="bl-read-bar">
        <RibbonMenuButton
          label="Dosya"
          menu={
            <Menu>
              <MenuItem label="Belgeyi Düzenle" onClick={onExit} />
            </Menu>
          }
        />
        <RibbonMenuButton
          label="Araçlar"
          menu={
            <Menu>
              <MenuItem label="Bul" shortcut="Ctrl+F" onClick={onFind} />
            </Menu>
          }
        />
        <RibbonMenuButton
          label="Görünüm"
          menu={
            <Menu>
              <MenuItem label="Belgeyi Düzenle" onClick={onExit} />
              <MenuSeparator />
              <MenuItem
                label="Sütun Genişliği"
                submenu={
                  <Menu>
                    <MenuItem label="Dar" checked={width === "narrow"} onClick={() => setWidth("narrow")} />
                    <MenuItem label="Varsayılan" checked={width === "default"} onClick={() => setWidth("default")} />
                    <MenuItem label="Geniş" checked={width === "wide"} onClick={() => setWidth("wide")} />
                  </Menu>
                }
              />
              <MenuItem
                label="Sayfa Rengi"
                submenu={
                  <Menu>
                    <MenuItem label="Yok" checked={bg === "none"} onClick={() => setBg("none")} />
                    <MenuItem label="Sepya" checked={bg === "sepia"} onClick={() => setBg("sepia")} />
                    <MenuItem label="Ters" checked={bg === "inverse"} onClick={() => setBg("inverse")} />
                  </Menu>
                }
              />
            </Menu>
          }
        />
        <span className="bl-read-title">{title}</span>
      </div>
      <div className="bl-read-stage">
        <button type="button" className="bl-read-nav" disabled={page === 0} onClick={() => setPage((p) => Math.max(0, p - 1))} aria-label="Önceki ekran">
          <ChevronLeft24Regular />
        </button>
        <div ref={ref} className={`bl-read-cols bl-doc is-${width}`} dangerouslySetInnerHTML={{ __html: html }} />
        <button type="button" className="bl-read-nav" disabled={page >= pages - 1} onClick={() => setPage((p) => Math.min(pages - 1, p + 1))} aria-label="Sonraki ekran">
          <ChevronRight24Regular />
        </button>
      </div>
      <div className="bl-read-foot">
        Ekran {page + 1} / {pages}
      </div>
    </div>
  );
}
