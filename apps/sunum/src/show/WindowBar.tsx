/** Okuma Görünümü / birey tarafından göz atılan gösteri: alt durum çubuğu. */
export function WindowBar({ index, count, onBack, onNext, onMenu, onViewChange, onExit }: { index: number; count: number; onBack: () => void; onNext: () => void; onMenu: (e: React.MouseEvent) => void; onViewChange?: (v: "normal" | "sorter") => void; onExit: () => void }) {
  return (
    <div className="show-bar" onClick={(e) => e.stopPropagation()}>
      <span className="show-bar-title">{`Slayt ${index + 1} / ${count}`}</span>
      <span className="grow" />
      <button type="button" className="show-bar-btn" onClick={onBack} title="Önceki">
        ◀
      </button>
      <button type="button" className="show-bar-btn" onClick={onMenu} title="Menü">
        ☰
      </button>
      <button type="button" className="show-bar-btn" onClick={onNext} title="Sonraki">
        ▶
      </button>
      {onViewChange && (
        <>
          <button type="button" className="show-bar-btn" onClick={() => onViewChange("normal")} title="Normal">
            Normal
          </button>
          <button type="button" className="show-bar-btn" onClick={() => onViewChange("sorter")} title="Slayt Sıralayıcısı">
            Sıralayıcı
          </button>
        </>
      )}
      <button type="button" className="show-bar-btn" onClick={onExit} title={onViewChange ? "Okuma Görünümünü Kapat" : "Gösteriyi Bitir"}>
        ✕
      </button>
    </div>
  );
}
