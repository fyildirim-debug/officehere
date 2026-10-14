#!/usr/bin/env bash
# OfficeHere sağlamlık ve performans ölçümlerinin tamamı (ölçüm turu). Depoyu değiştirmez;
# çıktılar .derlem/ altındadır (raporlar .derlem/rapor/).
#
# Ön koşullar:
#   pnpm install && pnpm build && pnpm web && pnpm --filter @officehere/desktop build:electron
#   LibreOffice (soffice + python3-uno), poppler-utils (pdfinfo/pdftotext/pdfimages)
#   python3: python-docx openpyxl python-pptx xlsxwriter pillow
#   Playwright Chromium (pnpm exec playwright install chromium); Electron için xvfb-run (Linux, ekransız)
# Kullanım: bash scripts/derlem/olcum/hepsi.sh [küçük_tmpfs_dizini]   (disk dolu senaryosu için isteğe bağlı)
set -euo pipefail
KOK="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$KOK"
export LC_ALL=C.UTF-8 LANG=C.UTF-8

pnpm derlem:uret                                  # derlem: LibreOffice UNO + python kütüphaneleri, formül sonuçları
pnpm derlem --temelsiz || true                    # 1. kaydet-aç turu (+ paket envanteri)
pnpm derlem:lo --temelsiz || true                 # 1. LibreOffice PDF denetimi
pnpm bulanik --temelsiz --paketleme-yok || true   # 2. bozuk dosya (5 sn sınır, zaman aşımları 60 sn ile yeniden)
python3 scripts/derlem/olcum/perf_veri.py         # 3. büyük dosyalar
node --expose-gc --max-old-space-size=8192 scripts/derlem/olcum/perf_node.mjs --paketleme-yok   # 3. G-Ç performansı
node scripts/derlem/olcum/perf_tarayici.mjs       # 3/4. tarayıcı performansı (apps/*/dist gerekli)
node scripts/derlem/olcum/paket_boyut.mjs         # 4. paket boyutu
if command -v xvfb-run >/dev/null; then xvfb-run -a node scripts/derlem/olcum/electron_olcum.mjs; else node scripts/derlem/olcum/electron_olcum.mjs; fi   # 4/5. Electron
node scripts/derlem/olcum/kaydet_sim.mjs ${1:-}   # 5. disk dolu / eşzamanlı kayıt / yazarken çökme
pnpm test:perf                                    # model düzeyi büyük veri (50.000 satır)
echo "Raporlar: $KOK/.derlem/rapor/"
