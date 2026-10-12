/** Çizim sekmesi: kalem, kurşun kalem, vurgulayıcı ve silgiyle serbest mürekkep çizimi. */
import { Cursor24Regular, Delete24Regular, Edit24Regular, Eraser24Regular, Highlight24Regular, Pen24Regular } from "@fluentui/react-icons";
import { MenuHeader, RibbonButton, RibbonGroup, RibbonSplitButton, useMenuClose } from "@officehere/ui";
import type { ReactElement } from "react";
import type { InkTool } from "../state/store";
import type { RibbonCtx } from "./types";

type PenKind = Exclude<InkTool["tool"], "eraser">;

const DEFAULTS: Record<PenKind, InkTool> = {
  pen: { tool: "pen", color: "#000000", width: 1 },
  pencil: { tool: "pencil", color: "#595959", width: 1.5 },
  highlighter: { tool: "highlighter", color: "#FFFF00", width: 12 },
};

/** Her kalemin son kullanılan rengi/kalınlığı (oturum boyunca). */
const lastUsed: Partial<Record<PenKind, InkTool>> = {};

const PEN_COLORS = ["#000000", "#FFFFFF", "#7F7F7F", "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0", "#E84C8B"];
const HIGHLIGHT_COLORS = ["#FFFF00", "#00FF00", "#00FFFF", "#FF00FF", "#FFC000", "#FF6600", "#9999FF", "#C0C0C0"];
const PEN_WIDTHS = [0.25, 0.5, 1, 2, 3.5, 5];
const HIGHLIGHT_WIDTHS = [6, 8, 12, 16, 20];

export const PEN_LABELS: Record<PenKind, string> = { pen: "Kalem", pencil: "Karakalem", highlighter: "Vurgulayıcı" };

function current(kind: PenKind): InkTool {
  return lastUsed[kind] ?? DEFAULTS[kind];
}

function PenMenu({ kind, onPick }: { kind: PenKind; onPick: (t: InkTool) => void }) {
  const close = useMenuClose();
  const cur = current(kind);
  const colors = kind === "highlighter" ? HIGHLIGHT_COLORS : PEN_COLORS;
  const widths = kind === "highlighter" ? HIGHLIGHT_WIDTHS : PEN_WIDTHS;
  const pick = (patch: Partial<InkTool>) => {
    const next = { ...cur, ...patch };
    lastUsed[kind] = next;
    onPick(next);
    close();
  };
  return (
    <div className="oh-menu ink-menu">
      <MenuHeader>Renk</MenuHeader>
      <div className="ink-swatches">
        {colors.map((c) => (
          <button
            key={c}
            type="button"
            className={`ink-swatch${c === cur.color ? " is-selected" : ""}`}
            style={{ background: c }}
            title={c}
            aria-label={c}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => pick({ color: c })}
          />
        ))}
      </div>
      <MenuHeader>Kalınlık</MenuHeader>
      <div className="ink-widths">
        {widths.map((w) => (
          <button
            key={w}
            type="button"
            className={`ink-width${w === cur.width ? " is-selected" : ""}`}
            title={`${String(w).replace(".", ",")} nk`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => pick({ width: w })}
          >
            <span style={{ height: Math.max(1, Math.min(14, (w * 96) / 72)), background: kind === "highlighter" ? cur.color : "currentColor", opacity: kind === "highlighter" ? 0.6 : 1 }} />
            <em>{`${String(w).replace(".", ",")} nk`}</em>
          </button>
        ))}
      </div>
    </div>
  );
}

export function DrawTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a } = ctx;
  const ink = state.ink;
  const noSlide = !state.pres.slides.length;
  const use = (t: InkTool | null) => {
    if (a.s.editing) a.stopEditing();
    a.store.set({ ink: t, insert: null, cropId: null, painter: null, sel: t ? [] : a.s.sel });
  };
  const pen = (kind: PenKind, icon: ReactElement) => {
    const cur = current(kind);
    return (
      <RibbonSplitButton
        key={kind}
        size="large"
        icon={icon}
        label={PEN_LABELS[kind]}
        checked={ink?.tool === kind}
        disabled={noSlide}
        colorBar={cur.color}
        description={kind === "highlighter" ? "Slaydın üzerinde yarı saydam vurgu çizin." : "Slaydın üzerine serbest çizim yapın. Rengi ve kalınlığı oktan seçin."}
        onClick={() => use(ink?.tool === kind ? null : cur)}
        menu={<PenMenu kind={kind} onPick={(t) => use(t)} />}
      />
    );
  };
  const inkCount = a.slide()?.shapes.filter((s) => s.kind === "sp" && s.ink).length ?? 0;
  return (
    <>
      <RibbonGroup label="Çizim Araçları" icon={<Pen24Regular />}>
        <RibbonButton size="large" icon={<Cursor24Regular />} label="Seç" checked={!ink} description="Çizimi bırakıp nesne seçimine dönün (Esc)." onClick={() => use(null)} />
        <RibbonButton
          size="large"
          icon={<Eraser24Regular />}
          label="Silgi"
          checked={ink?.tool === "eraser"}
          disabled={noSlide}
          description="Üzerinden geçtiğiniz mürekkep çizimlerini silin."
          onClick={() => use(ink?.tool === "eraser" ? null : { tool: "eraser", color: "#000000", width: 1 })}
        />
        {pen("pen", <Pen24Regular />)}
        {pen("pencil", <Edit24Regular />)}
        {pen("highlighter", <Highlight24Regular />)}
      </RibbonGroup>
      <RibbonGroup label="Düzenle" icon={<Delete24Regular />}>
        <RibbonButton size="large" icon={<Delete24Regular />} label="Mürekkebi Temizle" disabled={!inkCount} description="Bu slayttaki tüm mürekkep çizimlerini siler." onClick={() => a.clearInk()} />
      </RibbonGroup>
    </>
  );
}
