#!/usr/bin/env bash
# Biçim test örneklerini LibreOffice ile yeniden üretir:
#   kaynak.fodt → lo-ornek.odt, lo-ornek.doc (97-2003), lo-ornek.rtf, lo-ornek.docx
# Kullanım: apps/belge/src/io/__fixtures__/bicimler/uret.sh   (soffice PATH'te olmalı)
set -euo pipefail
cd "$(dirname "$0")"
PROFILE="$(mktemp -d)"
trap 'rm -rf "$PROFILE"' EXIT
for fmt in odt doc rtf docx; do
  timeout 120 soffice --headless "-env:UserInstallation=file://$PROFILE" --convert-to "$fmt" --outdir . kaynak.fodt >/dev/null
  mv -f "kaynak.$fmt" "lo-ornek.$fmt"
done
ls -l lo-ornek.*
