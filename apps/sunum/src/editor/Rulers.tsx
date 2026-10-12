/** Cetvel: sıfır noktası slayt ortasında, santimetre bölmeli (Office gibi). */
export function Rulers({ width, height, stageLeft, stageTop, zoom, slideW, slideH, size }: { width: number; height: number; stageLeft: number; stageTop: number; zoom: number; slideW: number; slideH: number; size: number }) {
  const cm = (96 / 2.54) * zoom;
  const ticks = (len: number, origin: number, extent: number) => {
    const out: { pos: number; major: boolean; label?: string }[] = [];
    const half = extent / 2;
    const start = Math.ceil((-origin) / (cm / 4));
    const end = Math.floor((len - origin) / (cm / 4));
    for (let i = start; i <= end; i++) {
      const pos = origin + (i * cm) / 4;
      const major = i % 4 === 0;
      const v = Math.abs(i / 4);
      out.push({ pos, major, label: major && Math.abs((i * cm) / 4) <= half * zoom + 1 ? String(v) : undefined });
    }
    return out;
  };
  const ox = stageLeft + (slideW * zoom) / 2;
  const oy = stageTop + (slideH * zoom) / 2;
  return (
    <>
      <div className="ruler is-h" style={{ left: size, height: size, width }}>
        <div className="ruler-slide" style={{ left: stageLeft, width: slideW * zoom }} />
        {ticks(width, ox, slideW).map((t, i) => (
          <span key={i} className={`ruler-tick${t.major ? " is-major" : ""}`} style={{ left: t.pos }}>
            {t.label && <em>{t.label}</em>}
          </span>
        ))}
      </div>
      <div className="ruler is-v" style={{ top: size, width: size, height }}>
        <div className="ruler-slide" style={{ top: stageTop, height: slideH * zoom }} />
        {ticks(height, oy, slideH).map((t, i) => (
          <span key={i} className={`ruler-tick${t.major ? " is-major" : ""}`} style={{ top: t.pos }}>
            {t.label && <em>{t.label}</em>}
          </span>
        ))}
      </div>
      <div className="ruler-corner" style={{ width: size, height: size }} />
    </>
  );
}
