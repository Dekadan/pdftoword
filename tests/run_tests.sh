#!/usr/bin/env bash
# PDF -> Word aracının uçtan uca testi.
#
# Ne yapar:
#   1. Gerçekçi bir test PDF'i üretir (PyMuPDF): satır satır yerleştirilmiş metin,
#      tireli bölünmeler, tekrar eden üstbilgi + sayfa numaraları, başlık,
#      sayfa sınırını aşan paragraf, Türkçe karakterler.
#   2. webapp'i yerel bir sunucuda açar, Playwright/Chromium ile dönüştürür.
#   3. Çıktıyı (paragraflar + .docx) denetler.
#
# Gerekli: python3, node + global "playwright" (chromium). pip paketleri
#          (pymupdf, python-docx) otomatik kurulur.
set -euo pipefail
cd "$(dirname "$0")"

PORT="${PORT:-8137}"

echo "== Python ortamı =="
python3 -m venv .venv
# shellcheck disable=SC1091
. .venv/bin/activate
pip install -q --upgrade pip >/dev/null 2>&1 || true
pip install -q pymupdf python-docx

echo "== Test PDF üretiliyor =="
python make_test_pdf.py

echo "== Statik sunucu ($PORT) =="
python3 -m http.server "$PORT" --bind 127.0.0.1 --directory ../webapp >/dev/null 2>&1 &
SRV=$!
trap 'kill "$SRV" 2>/dev/null || true' EXIT
sleep 1

echo "== Tarayıcı sürücüsü (Playwright/Chromium) =="
NODE_PATH="$(npm root -g)" node test_run.js \
  "http://127.0.0.1:$PORT/index.html" \
  "$PWD/test_kitap.pdf" "$PWD/out.docx" "$PWD/paras.json"

echo "== Denetimler (tarayıcı uygulaması) =="
python assert.py

echo
echo "== Zor düzen testi (pdf2word motoru: paragraf birleştirme + tire onarımı) =="
python make_hard_pdf.py
python ../pdf2word.py test_zor.pdf zor.docx --no-toc --no-pagenum >/dev/null
python assert_hard.py zor.docx
