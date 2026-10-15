#!/usr/bin/env bash
# Kurulu OfficeHere'ı ekransız (Xvfb) başlatır: başlangıç ekranı sayfası yüklenir ve yerel günlüğe
# "başlatıldı" yazılırsa başarılıdır. Pardus 23/25 duman testi (.github/workflows/pardus.yml).
#
# Kullanım: scripts/apt/duman.sh [program]      (varsayılan: /opt/OfficeHere/officehere)
# Gerekenler: xvfb-run (xvfb, xauth), curl.
# Ortam:
#   OH_DUMAN_SURE=90         sayfa için beklenecek en uzun süre (sn)
#   OH_DUMAN_KORUMASIZ=1     Chromium korumalı alanı kapatılır (--no-sandbox). Yalnız kullanıcı ad alanı
#                            ve CAP_SYS_ADMIN vermeyen kapsayıcılar içindir; masaüstünde gerekmez.
# root olarak çağrılırsa uygulama "duman" adlı sıradan bir kullanıcıyla açılır.
set -euo pipefail

prog=${1:-/opt/OfficeHere/officehere}
sure=${OH_DUMAN_SURE:-90}
port=${OH_DUMAN_PORT:-9333}

[[ -x "$prog" ]] || { echo "Program yok: $prog" >&2; exit 1; }
for arac in xvfb-run curl; do
  command -v "$arac" >/dev/null || { echo "Eksik araç: $arac" >&2; exit 1; }
done

# Paketin bildirmediği bir paylaşılan kitaplık eksikse burada görünür.
eksik=$(ldd "$prog" | grep "not found" || true)
if [[ -n "$eksik" ]]; then
  echo "Eksik kitaplıklar:" >&2
  echo "$eksik" >&2
  exit 1
fi

veri=$(mktemp -d)
kayit=$(mktemp)
calistir=()
if [[ $(id -u) -eq 0 ]]; then
  id duman >/dev/null 2>&1 || useradd -m duman
  chown duman "$veri" "$kayit"
  calistir=(runuser -u duman --)
fi
bayraklar=(--remote-debugging-port="$port")
[[ "${OH_DUMAN_KORUMASIZ:-}" == 1 ]] && bayraklar+=(--no-sandbox)

# Hesap ve ağ yok: güncelleştirme denetimi kapalı, kullanıcı verisi geçici dizinde.
# Ayrı oturumda başlatılır: bitince Xvfb ve Electron süreçleri birlikte kapatılır.
setsid "${calistir[@]}" env OH_USER_DATA="$veri" OH_UPDATE_MODE=off \
  xvfb-run -a -s "-screen 0 1280x800x24" "$prog" "${bayraklar[@]}" >"$kayit" 2>&1 &
pid=$!

bitir() {
  kill -TERM -- "-$pid" 2>/dev/null || true
  pkill -TERM -f "^$prog" 2>/dev/null || true
  wait "$pid" 2>/dev/null || true
}
trap bitir EXIT

baslangic=$SECONDS
sayfa=""
while ((SECONDS - baslangic < sure)); do
  if ! kill -0 "$pid" 2>/dev/null; then
    echo "OfficeHere sayfa yüklenmeden kapandı. Çıktı:" >&2
    cat "$kayit" >&2
    exit 1
  fi
  sayfa=$(curl -fsS "http://127.0.0.1:$port/json/list" 2>/dev/null | grep -o '"url": *"file://[^"]*/apps/baslangic/index.html"' || true)
  [[ -n "$sayfa" ]] && break
  sleep 1
done
if [[ -z "$sayfa" ]]; then
  echo "Başlangıç ekranı ${sure} sn içinde yüklenmedi. Çıktı:" >&2
  cat "$kayit" >&2
  exit 1
fi
gunluk="$veri/gunluk/officehere.log"
for _ in $(seq 1 10); do
  grep -q "başlatıldı" "$gunluk" 2>/dev/null && break
  sleep 1
done
grep -q "başlatıldı" "$gunluk" 2>/dev/null || { echo "Yerel günlükte açılış kaydı yok: $gunluk" >&2; exit 1; }

sistem=$(sed -n 's/^PRETTY_NAME="\{0,1\}\([^"]*\)"\{0,1\}$/\1/p' /etc/os-release)
surum=$(dpkg-query -W -f '${Version}' officehere 2>/dev/null || echo "?")
echo "Duman testi geçti: $sistem, officehere $surum, başlangıç ekranı $((SECONDS - baslangic)) sn'de yüklendi."
grep "başlatıldı" "$gunluk" | tail -1
