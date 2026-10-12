/**
 * Geçiş sesleri: ses dosyası dağıtmak yerine Web Audio API ile bu projede
 * sentezlenen kısa efektler. Sunuda yalnız ses kimliği saklanır.
 */

export const TRANSITION_SOUNDS: { id: string; label: string }[] = [
  { id: "bell", label: "Çan" },
  { id: "chime", label: "Zil" },
  { id: "beep", label: "Bip" },
  { id: "drum", label: "Davul" },
  { id: "click", label: "Tıklama" },
  { id: "whoosh", label: "Rüzgâr" },
  { id: "burst", label: "Patlama" },
  { id: "rise", label: "Yükselen Ton" },
];

/** Özel değerler: "stop" önceki sesi durdurur. */
export const STOP_SOUND = "stop";

export function soundLabel(id: string | undefined): string {
  if (!id) return "[Ses Yok]";
  if (id === STOP_SOUND) return "[Önceki Sesi Durdur]";
  return TRANSITION_SOUNDS.find((s) => s.id === id)?.label ?? id;
}

let ctx: AudioContext | null = null;
let playing: { stop: () => void }[] = [];

function audio(): AudioContext | null {
  try {
    const AC = window.AudioContext ?? (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
    if (!AC) return null;
    ctx ??= new AC();
    if (ctx.state === "suspended") void ctx.resume().catch(() => undefined);
    return ctx;
  } catch {
    return null;
  }
}

function tone(ac: AudioContext, type: OscillatorType, f0: number, f1: number, dur: number, gain = 0.25, at = 0) {
  const t = ac.currentTime + at;
  const o = ac.createOscillator();
  const g = ac.createGain();
  o.type = type;
  o.frequency.setValueAtTime(f0, t);
  o.frequency.exponentialRampToValueAtTime(Math.max(1, f1), t + dur);
  g.gain.setValueAtTime(gain, t);
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  o.connect(g).connect(ac.destination);
  o.start(t);
  o.stop(t + dur + 0.02);
  playing.push({ stop: () => o.stop() });
}

function noise(ac: AudioContext, dur: number, gain: number, filter?: { type: BiquadFilterType; f0: number; f1: number }) {
  const len = Math.max(1, Math.floor(ac.sampleRate * dur));
  const buf = ac.createBuffer(1, len, ac.sampleRate);
  const d = buf.getChannelData(0);
  for (let i = 0; i < len; i++) d[i] = Math.random() * 2 - 1;
  const src = ac.createBufferSource();
  src.buffer = buf;
  const g = ac.createGain();
  const t = ac.currentTime;
  g.gain.setValueAtTime(gain, t);
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  let node: AudioNode = src;
  if (filter) {
    const f = ac.createBiquadFilter();
    f.type = filter.type;
    f.frequency.setValueAtTime(filter.f0, t);
    f.frequency.exponentialRampToValueAtTime(filter.f1, t + dur);
    node.connect(f);
    node = f;
  }
  node.connect(g).connect(ac.destination);
  src.start(t);
  playing.push({ stop: () => src.stop() });
}

export function stopSounds() {
  for (const p of playing) {
    try {
      p.stop();
    } catch {
      /* zaten bitti */
    }
  }
  playing = [];
}

export function playTransitionSound(id: string | undefined) {
  if (!id) return;
  if (id === STOP_SOUND) return stopSounds();
  const ac = audio();
  if (!ac) return;
  switch (id) {
    case "bell":
      tone(ac, "sine", 880, 870, 1.2, 0.3);
      tone(ac, "sine", 1760, 1750, 0.8, 0.08);
      break;
    case "chime":
      tone(ac, "triangle", 1318, 1310, 0.5, 0.2);
      tone(ac, "triangle", 988, 980, 0.7, 0.2, 0.18);
      break;
    case "beep":
      tone(ac, "square", 1000, 1000, 0.15, 0.08);
      break;
    case "drum":
      tone(ac, "sine", 160, 45, 0.45, 0.6);
      break;
    case "click":
      noise(ac, 0.04, 0.4, { type: "highpass", f0: 2000, f1: 4000 });
      break;
    case "whoosh":
      noise(ac, 0.7, 0.25, { type: "bandpass", f0: 300, f1: 3000 });
      break;
    case "burst":
      noise(ac, 0.5, 0.5, { type: "lowpass", f0: 3000, f1: 150 });
      break;
    case "rise":
      tone(ac, "sawtooth", 220, 880, 0.6, 0.08);
      break;
    default:
      break;
  }
}
