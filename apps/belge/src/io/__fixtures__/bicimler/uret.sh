#!/usr/bin/env bash
# Biçim test örneklerini LibreOffice ile yeniden üretir:
#   kaynak.fodt → lo-ornek.odt, lo-ornek.doc (97-2003), lo-ornek.rtf, lo-ornek.docx
#   izleme.fodt → izleme.odt, izleme.doc, izleme.rtf, izleme.docx (değişiklik izleme, açıklamalar, varsayılanlar)
#   nesne.fodt  → lo-nesne.odt (gömülü grafik alt belgesi "Object 1/" ve yedek görüntüsü; ODF geçiş katmanı)
# Kullanım: apps/belge/src/io/__fixtures__/bicimler/uret.sh   (soffice PATH'te olmalı)
set -euo pipefail
cd "$(dirname "$0")"
PROFILE="$(mktemp -d)"
trap 'rm -rf "$PROFILE"' EXIT
for fmt in odt doc rtf docx; do
  timeout 120 soffice --headless "-env:UserInstallation=file://$PROFILE" --convert-to "$fmt" --outdir . kaynak.fodt >/dev/null
  mv -f "kaynak.$fmt" "lo-ornek.$fmt"
  timeout 120 soffice --headless "-env:UserInstallation=file://$PROFILE" --convert-to "$fmt" --outdir . izleme.fodt >/dev/null
done
timeout 120 soffice --headless "-env:UserInstallation=file://$PROFILE" --convert-to odt --outdir . nesne.fodt >/dev/null
mv -f nesne.odt lo-nesne.odt
ls -l lo-ornek.* izleme.* lo-nesne.odt
