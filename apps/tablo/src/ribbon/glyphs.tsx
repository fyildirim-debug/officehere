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
