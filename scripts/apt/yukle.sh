#!/usr/bin/env bash
# depo-uret.sh'nin ürettiği apt deposunu Cloudflare R2'ye (S3 uyumlu API, aws CLI) yükler.
# İSKELET: kova, alan adı ve erişim anahtarları Furkan'ındır; hiçbiri depoya girmez (docs/dagitim.md).
#
# Kullanım: scripts/apt/yukle.sh <depo-dizini>
#
# Ortam (CI'da onaylı "apt" GitHub ortamının gizlileri/değişkenleri):
#   R2_HESAP                Cloudflare hesap kimliği (uç nokta https://<hesap>.r2.cloudflarestorage.com)
#   R2_KOVA                 kova adı (ör. officehere-apt); apt.<alan-adı> alt alan adına bağlı, herkese açık okunur
#   AWS_ACCESS_KEY_ID       R2 API belirteci: yalnız bu kovada "Object Read & Write"
#   AWS_SECRET_ACCESS_KEY
#   R2_UC                   uç noktayı geçersiz kılar (yerel S3 taklidiyle sınama ya da başka S3 sağlayıcısı)
#   OH_KURU=1               yalnız ne yapılacağını yazar (aws --dryrun)
#
# Sıra önemlidir: önce pool/ (yeni .deb), sonra dists/ (imzalı dizinler). Böylece istemci hiçbir an
# henüz yüklenmemiş bir paketi gösteren dizin görmez. Eski .deb dosyaları silinmez.
set -euo pipefail

dizin=${1:?Kullanım: $0 <depo-dizini>}
: "${R2_HESAP:?R2_HESAP gerekli}" "${R2_KOVA:?R2_KOVA gerekli}"
: "${AWS_ACCESS_KEY_ID:?AWS_ACCESS_KEY_ID gerekli}" "${AWS_SECRET_ACCESS_KEY:?AWS_SECRET_ACCESS_KEY gerekli}"
command -v aws >/dev/null || { echo "aws CLI yok (GitHub ubuntu-latest'te kurulu gelir)" >&2; exit 1; }
for f in dists/stable/InRelease dists/stable/Release.gpg officehere.pgp officehere.sources; do
  [[ -f "$dizin/$f" ]] || { echo "Eksik: $dizin/$f (önce depo-uret.sh)" >&2; exit 1; }
done

uc=${R2_UC:-https://${R2_HESAP}.r2.cloudflarestorage.com}
# R2 bölge olarak "auto" ister; aws CLI'nin yeni sağlama toplamı başlıklarını yalnız gerekince gönder.
export AWS_DEFAULT_REGION=auto AWS_REQUEST_CHECKSUM_CALCULATION=when_required AWS_RESPONSE_CHECKSUM_VALIDATION=when_required
kuru=()
[[ "${OH_KURU:-}" == 1 ]] && kuru=(--dryrun)
s3() { aws s3 "$@" --endpoint-url "$uc" --only-show-errors "${kuru[@]}"; }

# Paketler değişmez: uzun önbellek.
s3 sync "$dizin/pool" "s3://$R2_KOVA/pool" \
  --cache-control "public, max-age=31536000, immutable" --content-type application/vnd.debian.binary-package
# İmzalı dizinler kısa önbellekle; artık listelenmeyenler silinir.
s3 sync "$dizin/dists" "s3://$R2_KOVA/dists" --delete --cache-control "public, max-age=300"
for f in officehere.pgp officehere.asc officehere.sources; do
  s3 cp "$dizin/$f" "s3://$R2_KOVA/$f" --cache-control "public, max-age=3600"
done
echo "Yüklendi: s3://$R2_KOVA (uç nokta $uc)"
