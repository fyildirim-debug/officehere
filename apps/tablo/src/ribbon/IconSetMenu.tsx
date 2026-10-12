/** Koşullu Biçimlendirme > Simge Kümeleri galerisi. */
import { useEffect, useRef } from "react";
import { MenuSeparator, useMenuClose } from "@officehere/ui";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { ICON_SETS, defaultThresholds, iconCount } from "../model/condfmt";
import { drawIcon } from "../grid/icons";

/** Kümenin simgelerini yan yana (yüksekten düşüğe) çizen küçük önizleme. */
export function IconSetPreview({ set, size = 16 }: { set: string; size?: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const n = iconCount(set);
  useEffect(() => {
    const c = ref.current;
    const ctx = c?.getContext("2d");
    if (!c || !ctx) return;
    const dpr = window.devicePixelRatio || 1;
    c.width = n * (size + 3) * dpr;
    c.height = size * dpr;
    ctx.scale(dpr, dpr);
    for (let i = 0; i < n; i++) drawIcon(ctx, set, n - 1 - i, i * (size + 3), 0, size);
  }, [set, n, size]);
  return <canvas ref={ref} style={{ width: n * (size + 3), height: size }} aria-hidden="true" />;
}

export function IconSetMenu({ ctl }: { ctl: Controller }) {
  const close = useMenuClose();
  return (
    <div className="oh-menu tb-iconset-menu" style={{ padding: 6, width: 300 }}>
      {ICON_SETS.map((g, gi) => (
        <div key={g.group}>
          {gi > 0 && <MenuSeparator />}
          <div className="tb-iconset-group">{g.group}</div>
          <div className="tb-iconset-grid">
            {g.sets.map((s) => (
              <button
                key={s.id}
                type="button"
                className="oh-gallery-item tb-iconset-item"
                title={s.label}
                aria-label={s.label}
                onClick={() => {
                  D.addCondRule(ctl, { type: "iconSet", iconSet: s.id, iconThresholds: defaultThresholds(iconCount(s.id)) });
                  close();
                }}
              >
                <IconSetPreview set={s.id} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
