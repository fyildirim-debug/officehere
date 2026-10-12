/** Geri al / yinele yığını; her girdi kendi geri alma ve yineleme işlevini taşır. */
export interface UndoEntry {
  label: string;
  undo: () => void;
  redo: () => void;
  /** Yaklaşık bellek maliyeti (hücre sayısı). */
  cost?: number;
}

export class UndoManager {
  private past: UndoEntry[] = [];
  private future: UndoEntry[] = [];
  private totalCost = 0;
  constructor(
    private readonly limit = 100,
    private readonly costLimit = 4_000_000,
  ) {}

  push(e: UndoEntry): void {
    this.past.push(e);
    this.totalCost += e.cost ?? 1;
    this.future = [];
    while (this.past.length > this.limit || (this.totalCost > this.costLimit && this.past.length > 1)) {
      const old = this.past.shift()!;
      this.totalCost -= old.cost ?? 1;
    }
  }

  /** Geçerli yığın derinliği (squash için işaret). */
  mark(): number {
    return this.past.length;
  }

  /** İşaretten sonra eklenen girdileri tek adımda birleştirir. */
  squash(since: number, label: string): void {
    if (this.past.length - since < 2) return;
    const parts = this.past.splice(since);
    const cost = parts.reduce((n, e) => n + (e.cost ?? 1), 0);
    this.past.push({
      label,
      cost,
      undo: () => [...parts].reverse().forEach((e) => e.undo()),
      redo: () => parts.forEach((e) => e.redo()),
    });
  }

  undo(): UndoEntry | undefined {
    const e = this.past.pop();
    if (!e) return undefined;
    this.totalCost -= e.cost ?? 1;
    e.undo();
    this.future.push(e);
    return e;
  }

  redo(): UndoEntry | undefined {
    const e = this.future.pop();
    if (!e) return undefined;
    e.redo();
    this.past.push(e);
    this.totalCost += e.cost ?? 1;
    return e;
  }

  get undoCount(): number {
    return this.past.length;
  }

  get canUndo(): boolean {
    return this.past.length > 0;
  }

  get canRedo(): boolean {
    return this.future.length > 0;
  }

  get undoLabel(): string | undefined {
    return this.past[this.past.length - 1]?.label;
  }

  get redoLabel(): string | undefined {
    return this.future[this.future.length - 1]?.label;
  }

  /** Son girdileri listeler (Geri Al açılır menüsü). */
  recent(n = 10): string[] {
    return this.past.slice(-n).reverse().map((e) => e.label);
  }

  clear(): void {
    this.past = [];
    this.future = [];
    this.totalCost = 0;
  }
}
