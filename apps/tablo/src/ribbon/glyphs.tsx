/** Fluent kitaplığında karşılığı olmayan küçük şerit simgeleri (özgün çizim). */
export function SigmaGlyph({ size = 20 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
      <path d="M15 4H5l5.5 6L5 16h10" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
    </svg>
  );
}

export function FxGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <text x="16" y="22" textAnchor="middle" fontSize="17" fontStyle="italic" fontFamily="Georgia, serif" fill="currentColor">
        fx
      </text>
    </svg>
  );
}

export function FreezeGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect x="4" y="5" width="24" height="22" fill="#fff" stroke="currentColor" strokeWidth="1.2" />
      <rect x="4" y="5" width="24" height="6" fill="#CFE8D9" />
      <rect x="4" y="5" width="7" height="22" fill="#CFE8D9" />
      <path d="M4 11h24M11 5v22" stroke="#107C41" strokeWidth="1.6" />
    </svg>
  );
}

export function CalcGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect x="7" y="3" width="18" height="26" rx="2" fill="none" stroke="currentColor" strokeWidth="1.3" />
      <rect x="10" y="6" width="12" height="5" fill="none" stroke="currentColor" strokeWidth="1" />
      {[0, 1, 2].map((r) => [0, 1, 2].map((c) => <rect key={`${r}${c}`} x={10 + c * 4.3} y={14 + r * 4.5} width="3" height="3" fill="currentColor" />))}
    </svg>
  );
}

/** PivotTable: başlık satırı/sütunu vurgulu tablo ve döndürme oku (özgün çizim). */
export function PivotGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect x="9" y="9" width="19" height="19" fill="#fff" stroke="currentColor" strokeWidth="1.2" />
      <rect x="9" y="9" width="19" height="5" fill="#CFE8D9" stroke="currentColor" strokeWidth="1.2" />
      <rect x="9" y="9" width="6" height="19" fill="#CFE8D9" stroke="currentColor" strokeWidth="1.2" />
      <path d="M15 19h13M15 24h13M21 14v14" stroke="currentColor" strokeWidth="0.9" />
      <path d="M4 14V7a2 2 0 0 1 2-2h7" fill="none" stroke="#107C41" strokeWidth="1.5" />
      <path d="M11 2.5L13.5 5 11 7.5M1.5 11.5L4 14l2.5-2.5" fill="none" stroke="#107C41" strokeWidth="1.5" />
    </svg>
  );
}

/** Mini grafik (Çizgi / Sütun) simgesi. */
export function SparkGlyph({ kind, size = 32 }: { kind: "line" | "column" | "winLoss"; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect x="3" y="8" width="26" height="16" fill="#fff" stroke="currentColor" strokeWidth="1.1" />
      {kind === "line" && <path d="M6 20l5-6 4 3 5-7 6 5" fill="none" stroke="#107C41" strokeWidth="1.6" strokeLinejoin="round" />}
      {kind === "column" && [6, 11, 16, 21, 25].map((x, i) => <rect key={x} x={x} y={[14, 11, 17, 12, 15][i]} width="3" height={[8, 11, 5, 10, 7][i]} fill="#107C41" />)}
      {kind === "winLoss" && [6, 11, 16, 21, 25].map((x, i) => <rect key={x} x={x} y={i % 2 ? 16 : 11} width="3" height="5" fill={i % 2 ? "#C00000" : "#107C41"} />)}
    </svg>
  );
}

/** Etkileyenleri/Etkilenenleri İzle: hücreler arası mavi ok. */
export function TraceGlyph({ kind, size = 20 }: { kind: "prec" | "dep"; size?: number }) {
  const from = kind === "prec" ? { x: 4, y: 5 } : { x: 9, y: 10 };
  const to = kind === "prec" ? { x: 15, y: 15 } : { x: 17, y: 4 };
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
      <rect x="1.5" y="1.5" width="6" height="6" fill="none" stroke="currentColor" strokeWidth="0.9" />
      <rect x="12.5" y="12.5" width="6" height="6" fill="none" stroke="currentColor" strokeWidth="0.9" />
      <path d={`M${from.x} ${from.y}L${to.x} ${to.y}`} stroke="#2F6FD8" strokeWidth="1.4" />
      <circle cx={from.x} cy={from.y} r="1.6" fill="#2F6FD8" />
      <path d={`M${to.x} ${to.y}l-4.2 -0.6 3.6 -3.6z`} fill="#2F6FD8" transform={kind === "dep" ? `rotate(-90 ${to.x} ${to.y})` : undefined} />
    </svg>
  );
}

/** İzleme Penceresi: büyüteçli pencere. */
export function WatchGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect x="3" y="5" width="22" height="18" rx="1" fill="#fff" stroke="currentColor" strokeWidth="1.2" />
      <path d="M3 10h22" stroke="currentColor" strokeWidth="1.2" />
      <circle cx="21" cy="21" r="5" fill="#fff" stroke="#107C41" strokeWidth="1.6" />
      <path d="M24.6 24.6L29 29" stroke="#107C41" strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
}
