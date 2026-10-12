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

const COL_W = { narrow: 400, default: 500, wide: 720 } as const;

export function ReadMode({ html, title, onExit, onFind }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [page, setPage] = useState(0);
  const [pages, setPages] = useState(1);
  const [bg, setBg] = useState<Bg>("none");
  const [width, setWidth] = useState<"narrow" | "default" | "wide">("default");

  const [step, setStep] = useState(1);
  /**
   * Word'deki gibi ekran, ortalanmış sabit genişlikte sütunlara bölünür: sütun sayısı pencere genişliğine
   * göre (en çok 3), kenar boşlukları eşit; her ekran bir sütun grubudur.
   */
  const measure = useCallback(() => {
    const el = ref.current;
    if (!el) return;
    const W = el.clientWidth;
    const c = COL_W[width];
    const n = Math.max(1, Math.min(3, Math.floor((W - 48) / (c + 80))));
    const colW = Math.min(c, W - 64);
    // Her sütun ekranın 1/n'lik diliminde ortalanır; ekran adımı tam pencere genişliğidir.
    const g = (W - n * colW) / n;
    const m = g / 2;
    el.style.paddingLeft = `${m}px`;
    el.style.paddingRight = `${m}px`;
    el.style.columnWidth = `${colW}px`;
    el.style.columnCount = String(n);
    el.style.columnGap = `${g}px`;
    setStep(W);
    const total = Math.max(1, Math.round((el.scrollWidth - 2 * m + g) / (colW + g)));
    const pg = Math.max(1, Math.ceil(total / n));
    setPages(pg);
    setPage((p) => Math.min(p, pg - 1));
  }, [width]);

  useLayoutEffect(measure, [measure, html, width]);
  useEffect(() => {
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [measure]);

  useEffect(() => {
    const el = ref.current;
    if (el) el.scrollTo({ left: page * step, behavior: "smooth" });
  }, [page, step]);

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
