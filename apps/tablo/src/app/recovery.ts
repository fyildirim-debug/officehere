/** Çökme uyarısındaki kurtarma eylemleri: kaydet, kurtarma kopyası indir, geri al. */
import { baseName, downloadBlob } from "@officehere/core";
import type { Controller } from "../controller";
import type { RecoveryHandlers } from "./ErrorBoundary";

/** Kitabı XLSX olarak indirir; yazıcı da hata verirse CSV benzeri düz metin yedeği. */
export async function downloadRecoveryCopy(ctl: Controller): Promise<void> {
  const name = `${baseName(ctl.fileName)} (Kurtarılan)`;
  try {
    const { writeXlsx } = await import("../io/xlsx");
    const bytes = await writeXlsx(ctl.wb);
    downloadBlob(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${name}.xlsx`);
  } catch {
    // Son çare: her sayfanın hücre değerleri sekmeyle ayrılmış metin olarak
    const parts: string[] = [];
    for (const s of ctl.wb.sheets) {
      parts.push(`### ${s.name}`);
      const rows = new Map<number, string[]>();
      ctl.wb.engine.forEachCell(s.id, (r, c, rec) => {
        const line = rows.get(r) ?? [];
        line[c] = rec.formula ? `=${rec.formula.text}` : String(rec.value ?? "");
        rows.set(r, line);
      });
      for (const r of [...rows.keys()].sort((a, b) => a - b)) parts.push(`${r + 1}\t${rows.get(r)!.map((x) => x ?? "").join("\t")}`);
    }
    downloadBlob(new Blob([parts.join("\n")], { type: "text/plain" }), `${name}.txt`);
  }
}

export function recoveryHandlers(ctl: Controller): RecoveryHandlers {
  return {
    save: () => {
      if (ctl.saveHook) void ctl.saveHook();
      else void downloadRecoveryCopy(ctl);
    },
    download: () => void downloadRecoveryCopy(ctl),
    undo: () => {
      ctl.edit = null;
      ctl.ops.undo();
    },
  };
}
