/** Geliştirici satırı: Hakkında sayfaları ve başlangıç ekranının alt bilgisi. */
export const DEVELOPER = { name: "Furkan YILDIRIM", site: "furkanyildirim.com", url: "https://furkanyildirim.com" } as const;

export function DeveloperCredit({ variant = "line" }: { variant?: "line" | "footer" }) {
  return (
    <p className={`oh-credit is-${variant}`}>
      {variant === "footer" ? "OfficeHere · " : "Geliştiren: "}
      <strong>{DEVELOPER.name}</strong>
      {variant === "footer" ? " tarafından geliştirildi · " : " · "}
      <a href={DEVELOPER.url} target="_blank" rel="noopener noreferrer">
        {DEVELOPER.site}
      </a>
    </p>
  );
}
