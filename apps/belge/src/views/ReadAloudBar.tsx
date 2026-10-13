// Sesli Oku denetim çubuğu (şeridin altında): önceki/sonraki paragraf, oynat/duraklat, hız ve ses ayarları.
import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from "react";
import type { Editor } from "@tiptap/core";
import { Dismiss20Regular, Next20Regular, Pause20Regular, Play20Regular, Previous20Regular, Settings20Regular } from "@fluentui/react-icons";
import { ReadAloudPlayer, readStart, sortVoices, speechSupported, type ReadState } from "../editor/readAloud";

const PREF = "officehere.belge.sesliOku";

function loadPref(): { rate: number; voice: string | null } {
  try {
    const v = JSON.parse(localStorage.getItem(PREF) ?? "null") as { rate?: number; voice?: string | null } | null;
    return { rate: Math.min(2, Math.max(0.5, Number(v?.rate) || 1)), voice: v?.voice ?? null };
  } catch {
    return { rate: 1, voice: null };
  }
}

function savePref(p: { rate: number; voice: string | null }) {
  try {
    localStorage.setItem(PREF, JSON.stringify(p));
  } catch {
    // Tercih saklanamazsa yalnız bu oturumda geçerli olur.
  }
}

interface Props {
  editor: Editor;
  onClose: () => void;
  notify: (text: string, kind?: "info" | "warning" | "error" | "success") => void;
}

export function ReadAloudBar({ editor, onClose, notify }: Props) {
  const supported = speechSupported();
  const player = useMemo(() => (supported ? new ReadAloudPlayer(editor.view) : null), [editor, supported]);
  const subscribe = useCallback((cb: () => void) => player?.subscribe(cb) ?? (() => {}), [player]);
  const state = useSyncExternalStore(subscribe, () => (player ? `${player.state.status}:${player.state.index}:${player.state.count}` : "idle:0:0"));
  const [status] = state.split(":") as [ReadState["status"]];
  const [voices, setVoices] = useState<SpeechSynthesisVoice[]>([]);
  const [voicesReady, setVoicesReady] = useState(false);
  const [pref, setPref] = useState(loadPref);
  const [settings, setSettings] = useState(false);

  // Desteklenmeyen ortam: bilgilendir ve kapan.
  useEffect(() => {
    if (supported) return;
    notify("Sesli okuma bu ortamda desteklenmiyor (konuşma sentezleyicisi bulunamadı).", "warning");
    onClose();
  }, [supported]); // eslint-disable-line react-hooks/exhaustive-deps

  // Sesler tarayıcıda gecikmeli yüklenir.
  useEffect(() => {
    if (!supported) return;
    const synth = window.speechSynthesis;
    const load = () => {
      const list = sortVoices(synth.getVoices());
      setVoices(list);
      if (list.length) setVoicesReady(true);
    };
    load();
    synth.addEventListener?.("voiceschanged", load);
    // Ses listesi hiç gelmeyen ortamlarda (bazı Linux kurulumları) bir süre sonra boş kabul edilir.
    const t = window.setTimeout(() => setVoicesReady(true), 1200);
    return () => {
      synth.removeEventListener?.("voiceschanged", load);
      window.clearTimeout(t);
    };
  }, [supported]);

  const turkish = voices.filter((v) => v.lang.toLowerCase().startsWith("tr"));
  const voice = voices.find((v) => v.name === pref.voice) ?? turkish[0] ?? voices.find((v) => v.default) ?? null;

  useEffect(() => {
    if (!player) return;
    player.rate = pref.rate;
    player.voice = voice;
  }, [player, pref.rate, voice]);

  // Açılınca imleçten okumaya başlar; kapanınca durur.
  useEffect(() => {
    if (!player) return;
    player.rate = pref.rate;
    player.voice = voice;
    player.start(readStart(editor.state));
    return () => player.stop();
  }, [player]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!player) return null;
  const play = () => {
    if (status === "playing") player.pause();
    else if (status === "paused") player.resume();
    else player.start(readStart(editor.state));
  };
  const update = (p: Partial<typeof pref>) => {
    const next = { ...pref, ...p };
    setPref(next);
    savePref(next);
    player.rate = next.rate;
    player.voice = voices.find((v) => v.name === next.voice) ?? voice;
    player.restartCurrent();
  };

  return (
    <div className="bl-ra-bar" role="toolbar" aria-label="Sesli Oku">
      {voicesReady && !turkish.length && (
        <span className="bl-ra-warn" role="status">
          Türkçe ses bulunamadı; metin {voice ? `"${voice.name}" sesiyle` : "varsayılan sesle"} okunuyor. İşletim sisteminin konuşma ayarlarından Türkçe ses ekleyebilirsiniz.
        </span>
      )}
      <span className="bl-flex" />
      <button type="button" className="bl-ra-btn" title="Önceki paragraf" aria-label="Önceki paragraf" onClick={() => player.skip(-1)} disabled={status === "idle"}>
        <Previous20Regular />
      </button>
      <button type="button" className="bl-ra-btn is-main" title={status === "playing" ? "Duraklat" : "Oynat"} aria-label={status === "playing" ? "Duraklat" : "Oynat"} onClick={play}>
        {status === "playing" ? <Pause20Regular /> : <Play20Regular />}
      </button>
      <button type="button" className="bl-ra-btn" title="Sonraki paragraf" aria-label="Sonraki paragraf" onClick={() => player.skip(1)} disabled={status === "idle"}>
        <Next20Regular />
      </button>
      <span className="bl-ra-settings">
        <button type="button" className={`bl-ra-btn${settings ? " is-on" : ""}`} title="Ayarlar" aria-label="Ayarlar" aria-expanded={settings} onClick={() => setSettings(!settings)}>
          <Settings20Regular />
        </button>
        {settings && (
          <span className="bl-ra-pop" role="group" aria-label="Okuma ayarları">
            <label>
              Okuma hızı: {pref.rate.toLocaleString("tr-TR", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}×
              <input type="range" min={0.5} max={2} step={0.1} value={pref.rate} aria-label="Okuma hızı" onChange={(e) => update({ rate: Number(e.target.value) })} />
            </label>
            <label>
              Ses:
              <select value={voice?.name ?? ""} aria-label="Ses" onChange={(e) => update({ voice: e.target.value || null })} disabled={!voices.length}>
                {!voices.length && <option value="">Varsayılan</option>}
                {voices.map((v) => (
                  <option key={`${v.name}-${v.lang}`} value={v.name}>
                    {v.name} ({v.lang})
                  </option>
                ))}
              </select>
            </label>
          </span>
        )}
      </span>
      <button type="button" className="bl-ra-btn" title="Sesli Oku'yu kapat (Ctrl+Alt+Boşluk)" aria-label="Sesli Oku'yu kapat" onClick={onClose}>
        <Dismiss20Regular />
      </button>
    </div>
  );
}
