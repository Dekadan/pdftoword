# Testler

Aracın gerçek bir tarayıcıda uçtan uca doğrulandığı otomatik testler.

## Çalıştırma

```bash
cd tests
./run_tests.sh
```

Başarılıysa çıktının sonunda şunu görürsünüz:

```
SONUÇ: TÜM TESTLER GEÇTİ ✓✓✓
```

## Gereksinimler

- `python3` (test PDF’i üretmek ve `.docx`’i doğrulamak için `pymupdf`, `python-docx`
  otomatik kurulur)
- `node` + global **Playwright** (Chromium): `npm i -g playwright`
- Türkçe destekli bir serif font (Liberation Serif ya da DejaVu Serif — çoğu
  Linux’ta hazır gelir)

## Dosyalar

| Dosya | Görevi |
|-------|--------|
| `make_test_pdf.py` | Gerçekçi test PDF’i üretir: satır satır metin, tireli bölünmeler, tekrar eden üstbilgi + sayfa no, başlık, sayfa sınırını aşan paragraf. |
| `test_run.js` | Playwright ile Chromium’da `webapp`’i açar, PDF’i yükler, dönüştürür, `.docx`’i indirir ve paragrafları JSON’a döker. |
| `assert.py` | Çıktıyı denetler: tire birleşmesi, paragraf bütünlüğü, sayfa-aşan birleşme, üstbilgi/sayfa no temizliği, başlık algısı, Türkçe koruma ve geçerli `.docx`. |
| `run_tests.sh` | Hepsini sırayla çalıştırır. |

## Neyi doğrular?

- Ayrı satırlar **tek paragrafta** birleşiyor (paragraf bütünlüğü).
- Sayfa sınırını **aşan** paragraf tek paragraf oluyor.
- Satır sonu tireleri birleşiyor (`evre-`/`lerinde` → `evrelerinde`), büyük harfle
  devam eden bileşiklerde tire **korunuyor** (`Sovyet-`/`Rusya` → `Sovyet-Rusya`).
- Tekrar eden **üstbilgi** ve salt **sayfa numarası** satırları temizleniyor.
- **Başlık** (büyük punto) Word’de Heading stiliyle işaretleniyor.
- **Türkçe karakterler** ve cümleler birebir korunuyor.
- Üretilen **`.docx` geçerli** (katı OOXML okuyucu `python-docx` ile açılıyor).
