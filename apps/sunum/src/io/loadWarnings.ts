/**
 * Açılış uyarıları: dosyanın okunamayan ya da eksik bölümleri kullanıcıya
 * anlaşılır Türkçe cümlelerle bildirilir. Aynı türden slayt sorunları tek
 * cümlede toplanır ("1., 3. ve 5. slaytlar …"); teknik ayrıntılar ayrı tutulur.
 */

export interface SlideProblemText {
  one: (nums: string) => string;
  many: (nums: string) => string;
}

/** Slayt sorunu türleri ve iletileri. */
export const SLIDE_PROBLEM = {
  missing: {
    one: (n) => `${n} slayt dosyada bulunamadı; yerine boş slayt kondu.`,
    many: (n) => `${n} slaytlar dosyada bulunamadı; yerlerine boş slayt kondu.`,
  },
  corrupt: {
    one: (n) => `${n} slayt bozuk olduğu için okunamadı; yerine boş slayt kondu.`,
    many: (n) => `${n} slaytlar bozuk olduğu için okunamadı; yerlerine boş slayt kondu.`,
  },
  tooLarge: {
    one: (n) => `${n} slayt olağandışı büyük olduğu için okunmadı; yerine boş slayt kondu.`,
    many: (n) => `${n} slaytlar olağandışı büyük olduğu için okunmadı; yerlerine boş slayt kondu.`,
  },
  partial: {
    one: (n) => `${n} slayt bozuk; yalnız okunabilen kısmı açıldı.`,
    many: (n) => `${n} slaytlar bozuk; yalnız okunabilen kısımları açıldı.`,
  },
  rels: {
    one: (n) => `${n} slaydın ilişki bölümü eksik ya da bozuk; resimler ve düzen kaybolmuş olabilir.`,
    many: (n) => `${n} slaytların ilişki bölümleri eksik ya da bozuk; resimler ve düzen kaybolmuş olabilir.`,
  },
  notes: {
    one: (n) => `${n} slaydın konuşmacı notları okunamadı.`,
    many: (n) => `${n} slaytların konuşmacı notları okunamadı.`,
  },
  comments: {
    one: (n) => `${n} slaydın açıklamaları okunamadı.`,
    many: (n) => `${n} slaytların açıklamaları okunamadı.`,
  },
} satisfies Record<string, SlideProblemText>;

export type SlideProblem = keyof typeof SLIDE_PROBLEM;

/** [1, 3, 5] → "1., 3. ve 5." */
export function ordinalList(nums: number[]): string {
  const s = [...new Set(nums)].sort((a, b) => a - b).map((n) => `${n}.`);
  if (s.length <= 1) return s.join("");
  if (s.length > 8) return `${s.slice(0, 6).join(", ")} ve ${s.length - 6} diğer`;
  return `${s.slice(0, -1).join(", ")} ve ${s[s.length - 1]}`;
}

export class LoadWarnings {
  private items: string[] = [];
  private slideGroups = new Map<SlideProblem, number[]>();
  /** Teknik ayrıntılar (parça yolları, özgün hata iletileri); günlük ve "Ayrıntılar" için. */
  readonly details: string[] = [];

  /** Genel uyarı (yinelenenler bir kez). */
  add(message: string, detail?: string): void {
    if (!this.items.includes(message)) this.items.push(message);
    if (detail) this.detail(detail);
  }

  /** Belirli bir slayda ait sorun (1 tabanlı slayt numarası). */
  slide(kind: SlideProblem, n: number, detail?: string): void {
    const g = this.slideGroups.get(kind) ?? [];
    g.push(n);
    this.slideGroups.set(kind, g);
    if (detail) this.detail(`${n}. slayt: ${detail}`);
  }

  detail(text: string): void {
    if (this.details.length < 200 && !this.details.includes(text)) this.details.push(text);
  }

  get size(): number {
    return this.items.length + this.slideGroups.size;
  }

  /** Kullanıcıya gösterilecek cümleler: önce slayt sorunları, sonra genel uyarılar. */
  list(): string[] {
    const out: string[] = [];
    for (const [kind, nums] of this.slideGroups) {
      const t = SLIDE_PROBLEM[kind];
      const uniq = [...new Set(nums)];
      out.push(uniq.length === 1 ? t.one(ordinalList(uniq)) : t.many(ordinalList(uniq)));
    }
    return [...out, ...this.items];
  }
}
