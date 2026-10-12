/** Geçişler sekmesi: Önizleme, Bu Slayda Geçiş, Zamanlama. */
import { Play24Regular, SlideTransition24Regular, ArrowRepeatAll20Regular,
  Timer24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonRow, RibbonStack } from "@officehere/ui";
import { STOP_SOUND, TRANSITION_SOUNDS, playTransitionSound } from "../show/sounds";
import { GroupedGallery, TimeSpinner } from "./GroupedGallery";
import type { TransitionType } from "../model/types";
import { TRANSITIONS, transitionDef } from "../show/catalog";
import type { RibbonCtx } from "./types";

/** Galeri simgeleri (bu projeye özgü basit çizimler). */
export function TransitionIcon({ type }: { type: TransitionType }) {
  const a = "#c4462b";
  const b = "#9aa7b8";
  const frame = <rect x="3" y="5" width="26" height="18" fill="#fff" stroke="#7a7a7a" strokeWidth="1" />;
  let art: React.ReactNode = null;
  switch (type) {
    case "none":
      art = <path d="M8 9l16 10M24 9L8 19" stroke="#b0b0b0" strokeWidth="1.5" />;
      break;
    case "fade":
    case "reveal":
      art = (
        <>
          <defs>
            <linearGradient id={`tg-${type}`} x1="0" x2="1">
              <stop offset="0" stopColor={type === "reveal" ? "#222" : "#fff"} />
              <stop offset="1" stopColor={a} />
            </linearGradient>
          </defs>
          <rect x="4" y="6" width="24" height="16" fill={`url(#tg-${type})`} />
        </>
      );
      break;
    case "push":
    case "cover":
    case "uncover":
      art = (
        <>
          <rect x="4" y="6" width="12" height="16" fill={b} />
          <rect x="16" y="6" width="12" height="16" fill={a} />
          <path d="M11 14h10M18 11l3 3-3 3" stroke="#fff" strokeWidth="1.4" fill="none" />
        </>
      );
      break;
    case "wipe":
      art = (
        <>
          <rect x="4" y="6" width="14" height="16" fill={a} />
          <rect x="18" y="6" width="10" height="16" fill={b} opacity="0.6" />
        </>
      );
      break;
    case "split":
      art = (
        <>
          <rect x="4" y="6" width="24" height="16" fill={b} />
          <rect x="11" y="6" width="10" height="16" fill={a} />
        </>
      );
      break;
    case "cut":
      art = <path d="M9 8l14 12M23 8L9 20" stroke={a} strokeWidth="2" />;
      break;
    case "randomBar":
    case "blinds":
      art = (
        <>
          {[0, 1, 2, 3, 4, 5].map((i) => (
            <rect key={i} x={4 + i * 4} y="6" width={type === "blinds" ? 2.5 : (i % 3) + 1} height="16" fill={a} />
          ))}
        </>
      );
      break;
    case "shape":
      art = <circle cx="16" cy="14" r="7" fill={a} />;
      break;
    case "zoom":
      art = (
        <>
          <rect x="4" y="6" width="24" height="16" fill={b} />
          <rect x="11" y="10" width="10" height="8" fill={a} />
        </>
      );
      break;
    case "dissolve":
    case "checker":
      art = (
        <>
          {Array.from({ length: 24 }, (_, i) => (i % 2 === (Math.floor(i / 6) % 2) ? <rect key={i} x={4 + (i % 6) * 4} y={6 + Math.floor(i / 6) * 4} width="4" height="4" fill={a} /> : null))}
        </>
      );
      break;
    case "wheel":
    case "wedge":
      art = <path d="M16 14 L16 6 A8 8 0 0 1 24 14 Z" fill={a} />;
      break;
    case "newsflash":
      art = <rect x="11" y="9" width="10" height="10" fill={a} transform="rotate(25 16 14)" />;
      break;
  }
  return (
    <svg width="32" height="28" viewBox="0 0 32 28" aria-hidden="true">
      {frame}
      {art}
    </svg>
  );
}

function Tile({ type, label }: { type: TransitionType; label: string }) {
  return (
    <span style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 1, fontSize: 10, lineHeight: 1.1 }}>
      <TransitionIcon type={type} />
      <span style={{ whiteSpace: "nowrap", overflow: "hidden", maxWidth: 62, textOverflow: "ellipsis" }}>{label}</span>
    </span>
  );
}

export function TransitionsTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const slide = state.pres.slides[state.slideIdx];
  const t = slide?.transition;
  const type = t?.type ?? "none";
  const def = transitionDef(type);
  const disabled = !slide;
  return (
    <>
      <RibbonGroup label="Önizleme" icon={<Play24Regular />}>
        <RibbonButton size="large" icon={<Play24Regular />} label="Önizleme" disabled={disabled || type === "none"} description="Bu slaydın geçişini oynatın." onClick={cmd.previewTransition} />
      </RibbonGroup>
      <RibbonGroup label="Bu Slayda Geçiş" icon={<SlideTransition24Regular />}>
        <GroupedGallery
          columns={8}
          itemWidth={64}
          itemHeight={58}
          expandedColumns={7}
          selectedId={type}
          items={TRANSITIONS.map((d) => ({ id: d.type, label: d.label, group: d.group ?? "Hafif", preview: <Tile type={d.type} label={d.label} /> }))}
          onSelect={(id) => {
            const d = transitionDef(id as TransitionType);
            a.setTransition({ type: d.type, option: d.defaultOption, duration: t?.duration && t.type !== "none" ? t.duration : d.type === "cut" ? 0.1 : 1 });
            if (d.type !== "none") window.setTimeout(() => cmd.previewTransition(), 50);
          }}
        />
        <RibbonMenuButton
          size="large"
          icon={<SlideTransition24Regular />}
          label="Efekt Seçenekleri"
          disabled={!def.options}
          menu={
            <Menu>
              {(def.options ?? []).map((o) => (
                <MenuItem
                  key={o.id}
                  label={o.label}
                  checked={(t?.option ?? def.defaultOption) === o.id}
                  onClick={() => {
                    a.setTransition({ option: o.id });
                    window.setTimeout(() => cmd.previewTransition(), 50);
                  }}
                />
              ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Zamanlama" icon={<Timer24Regular />}>
        <RibbonStack>
          <RibbonRow>
            <span className="oh-spinner-label">Ses:</span>
            <select
              className="rb-select"
              aria-label="Ses"
              value={t?.sound ?? ""}
              disabled={disabled}
              onChange={(e) => {
                const v = e.target.value || undefined;
                a.setTransition({ sound: v });
                if (v) playTransitionSound(v);
              }}
            >
              <option value="">[Ses Yok]</option>
              <option value={STOP_SOUND}>[Önceki Sesi Durdur]</option>
              {TRANSITION_SOUNDS.map((x) => (
                <option key={x.id} value={x.id}>
                  {x.label}
                </option>
              ))}
            </select>
          </RibbonRow>
          <TimeSpinner label="Süre:" value={t?.duration ?? 1} step={0.25} max={59} tip="Süre (saniye)" disabled={disabled} onChange={(v) => a.setTransition({ duration: Math.max(0.01, v) })} />
          <RibbonButton icon={<ArrowRepeatAll20Regular />} label="Tümüne Uygula" disabled={disabled} onClick={() => a.applyTransitionToAll()} />
        </RibbonStack>
        <RibbonStack>
          <span className="oh-ribbon-text" style={{ fontWeight: 600 }}>
            Slaydı İlerlet
          </span>
          <RibbonCheckbox label="Fare Tıklatıldığında" checked={t?.advanceOnClick ?? true} onChange={(v) => a.setTransition({ advanceOnClick: v })} />
          <RibbonRow>
            <RibbonCheckbox label="Sonra:" checked={t?.advanceAfter !== undefined} onChange={(v) => a.setTransition({ advanceAfter: v ? (t?.advanceAfter ?? 5) : undefined })} />
            <TimeSpinner minutes value={t?.advanceAfter ?? 0} step={1} max={3600} width={84} tip="Bu süre sonra otomatik ilerle (dakika:saniye)" onChange={(v) => a.setTransition({ advanceAfter: v })} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
