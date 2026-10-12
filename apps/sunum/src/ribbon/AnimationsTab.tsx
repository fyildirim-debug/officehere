/** Animasyonlar sekmesi: Önizleme, Animasyon, Gelişmiş Animasyon, Zamanlama. */
import { useState } from "react";
import { ArrowDown20Regular, ArrowUp20Regular, Play24Regular, StarAdd24Regular, StarEdit24Regular, PaintBrush20Regular, AppsListDetail24Regular,
  Star24Regular,
  Timer24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, RibbonButton, RibbonCombo, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack, useMenuClose } from "@officehere/ui";
import type { AnimClass, Animation } from "../model/types";
import { ANIMATIONS, animDef, CLASS_COLORS, CLASS_LABELS, TRIGGER_LABELS, type AnimDef } from "../show/catalog";
import type { RibbonCtx } from "./types";

export function AnimStar({ cls, size = 22 }: { cls: AnimClass; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 2.5l2.8 6 6.5.7-4.9 4.4 1.4 6.4L12 16.8 6.2 20l1.4-6.4L2.7 9.2l6.5-.7z" fill={CLASS_COLORS[cls]} stroke="rgba(0,0,0,.25)" strokeWidth="0.6" />
    </svg>
  );
}

function Tile({ d }: { d: AnimDef | { id: "none"; label: string; cls?: AnimClass } }) {
  return (
    <span style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2, fontSize: 10, lineHeight: 1.1 }}>
      {"cls" in d && d.cls ? <AnimStar cls={d.cls} /> : <svg width="22" height="22" viewBox="0 0 22 22"><circle cx="11" cy="11" r="8" fill="none" stroke="#999" /><path d="M5 17L17 5" stroke="#999" /></svg>}
      <span style={{ whiteSpace: "nowrap", overflow: "hidden", maxWidth: 60, textOverflow: "ellipsis" }}>{d.label}</span>
    </span>
  );
}

function AnimMenu({ onPick }: { onPick: (d: AnimDef) => void }) {
  const close = useMenuClose();
  return (
    <div className="oh-menu" style={{ maxHeight: "70vh", width: 420 }}>
      {(["entr", "emph", "exit", "path"] as AnimClass[]).map((cls) => (
        <div key={cls}>
          <MenuHeader>{CLASS_LABELS[cls]}</MenuHeader>
          <div className="preset-grid" style={{ gridTemplateColumns: "repeat(6, 64px)" }}>
            {ANIMATIONS.filter((d) => d.cls === cls).map((d) => (
              <button
                key={d.id}
                type="button"
                className="preset-btn"
                title={d.label}
                style={{ height: 50 }}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  onPick(d);
                  close();
                }}
              >
                <Tile d={d} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

export function AnimationsTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const slide = state.pres.slides[state.slideIdx];
  const [painter, setPainter] = useState<Animation[] | null>(null);
  const selAnims = slide ? slide.animations.filter((x) => state.sel.includes(x.shapeId)) : [];
  const cur = selAnims[0];
  const def = cur ? animDef(cur.effect) : undefined;
  const hasSel = state.sel.length > 0;
  const update = (patch: Partial<Animation>) => {
    const ids = new Set(selAnims.map((x) => x.id));
    a.setAnimations((list) => list.map((x) => (ids.has(x.id) ? { ...x, ...patch } : x)), "anim-time");
  };
  const move = (dir: -1 | 1) => {
    if (!cur) return;
    a.setAnimations((list) => {
      const i = list.findIndex((x) => x.id === cur.id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= list.length) return list;
      const out = list.slice();
      [out[i], out[j]] = [out[j], out[i]];
      return out;
    });
  };
  const galleryItems = [{ id: "none", label: "Yok", preview: <Tile d={{ id: "none", label: "Yok" }} /> }, ...ANIMATIONS.map((d) => ({ id: d.id, label: `${d.label} (${CLASS_LABELS[d.cls]})`, preview: <Tile d={d} /> }))];
  return (
    <>
      <RibbonGroup label="Önizleme" icon={<Play24Regular />}>
        <RibbonButton size="large" icon={<Play24Regular />} label="Önizleme" disabled={!slide?.animations.length} description="Bu slayttaki animasyonları oynatın." onClick={cmd.previewAnimations} />
      </RibbonGroup>
      <RibbonGroup label="Animasyon" icon={<Star24Regular />} onLauncher={() => cmd.togglePane("animation")} launcherTip="Efekt Seçenekleri">
        <RibbonGallery
          columns={7}
          itemWidth={62}
          itemHeight={52}
          expandedColumns={6}
          selectedId={hasSel ? (cur?.effect ?? "none") : undefined}
          items={galleryItems}
          onSelect={(id) => {
            if (!hasSel) {
              a.message("Animasyon uygulamak için önce bir nesne seçin.", "info");
              return;
            }
            const d = animDef(id);
            a.applyAnimation(id, d?.cls ?? "entr", false, d?.defaultOption);
            window.setTimeout(() => cmd.previewAnimations(), 50);
          }}
        />
        <RibbonMenuButton
          size="large"
          icon={<StarEdit24Regular />}
          label="Efekt Seçenekleri"
          disabled={!def?.options}
          menu={
            <Menu>
              <MenuHeader>Yön</MenuHeader>
              {(def?.options ?? []).map((o) => (
                <MenuItem key={o.id} label={o.label} checked={(cur?.option ?? def?.defaultOption) === o.id} onClick={() => update({ option: o.id })} />
              ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Gelişmiş Animasyon" icon={<StarAdd24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<StarAdd24Regular />}
          label="Animasyon Ekle"
          disabled={!hasSel}
          description="Seçili nesneye ek bir animasyon efekti ekleyin."
          menu={<AnimMenu onPick={(d) => a.applyAnimation(d.id, d.cls, true, d.defaultOption)} />}
        />
        <RibbonStack>
          <RibbonButton icon={<AppsListDetail24Regular />} label="Animasyon Bölmesi" checked={state.pane === "animation"} onClick={() => cmd.togglePane("animation")} />
          <RibbonButton
            icon={<PaintBrush20Regular />}
            label="Animasyon Boyacısı"
            checked={!!painter}
            disabled={!cur && !painter}
            description="Bir nesnenin animasyonunu kopyalayıp diğerine uygulayın: önce kaynağı seçip tıklayın, sonra hedefi seçip yeniden tıklayın."
            onClick={() => {
              if (painter) {
                for (const p of painter) a.applyAnimation(p.effect, p.cls, false, p.option);
                setPainter(null);
              } else setPainter(selAnims);
            }}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Zamanlama" icon={<Timer24Regular />}>
        <RibbonStack>
          <div className="oh-spinner-wrap">
            <span className="oh-spinner-label">Başlangıç:</span>
            <RibbonCombo
              value={cur ? TRIGGER_LABELS[cur.trigger] : ""}
              editable={false}
              width={130}
              disabled={!cur}
              options={Object.values(TRIGGER_LABELS)}
              onCommit={(v) => {
                const k = Object.entries(TRIGGER_LABELS).find(([, l]) => l === v)?.[0] as Animation["trigger"] | undefined;
                if (k) update({ trigger: k });
              }}
            />
          </div>
          <RibbonSpinner label="Süre:" value={cur?.duration ?? 0.5} min={0.01} max={59} step={0.25} decimals={2} width={70} onChange={(v) => update({ duration: v })} />
          <RibbonSpinner label="Gecikme:" value={cur?.delay ?? 0} min={0} max={59} step={0.25} decimals={2} width={70} onChange={(v) => update({ delay: v })} />
        </RibbonStack>
        <RibbonStack>
          <span className="oh-ribbon-text" style={{ fontWeight: 600 }}>
            Animasyonu Yeniden Sırala
          </span>
          <RibbonButton icon={<ArrowUp20Regular />} label="Önce Taşı" disabled={!cur} onClick={() => move(-1)} />
          <RibbonButton icon={<ArrowDown20Regular />} label="Sonra Taşı" disabled={!cur} onClick={() => move(1)} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
