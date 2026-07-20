# 📖 PDF → Word · Kitap Dönüştürücü

Kitap PDF’lerini, **satırları doğru paragraflara birleştirerek** düzenlenebilir
Word (`.docx`) dosyasına çeviren araçlar. Yayınevinde 500–600 sayfalık kitapları
Word’e aktarırken yaşanan dertler için yapıldı:

1. **“Her satır ayrı oturuyor, paragraf oluşmuyor.”** → Satırlar okunup paragraf
   bütünlüğü kurulur.
2. **“Kelimeler bölünüyor: `dışlaya rak`, `de ğil`, `mü cadelenin`.”** → Satır sonu
   tireleri doğru birleştirilir → `dışlayarak`, `değil`, `mücadelenin`.
3. **“Aşırı kredi/ücret harcamak istemiyorum.”** → Dönüştürmeyi yapay zekâ değil,
   sizin bilgisayarınızdaki program yapar: **kullanmak bedava, çevrimdışı, gizli.**

---

## İki araç var — hangisini kullanmalı?

| | **`pdf2word.py`** (Python) | **`webapp/`** (tarayıcı) |
|---|---|---|
| Kalite | ⭐ **En iyi** (önerilen) | İyi (basit kitaplarda) |
| Satır sonu tireleri | ✅ Tam doğru birleştirir | ⚠️ Sınırlı (bkz. not) |
| Kenar notu / dipnot | ✅ Gövdeyi bölmeden ayırır | Kısmen |
| Toplu (klasör) işleme | ✅ Var | Tek tek |
| Kurulum | Python gerekir | ❌ Gerekmez, çift tıkla aç |
| Maliyet / gizlilik | Bedava, çevrimdışı | Bedava, çevrimdışı |

> **Neden Python sürümü daha iyi?** Türkçe kitaplarda satır sonları çoğunlukla
> *yumuşak tire* (görünmez tire) ile bölünür. Tarayıcıdaki pdf.js bu tireyi
> **siler**, o yüzden `dışlaya` ile `rak` arasına yanlış boşluk girer. Python
> sürümündeki PyMuPDF tireyi **korur** ve kelimeyi doğru birleştirir. **Gerçek
> kitaplar için `pdf2word.py` önerilir.**

---

## 🚀 A) `pdf2word.py` — önerilen yol

### Kurulum (bir kez)
```bash
pip install -r requirements.txt        # ya da: pip install pymupdf python-docx
```

### Kullanım
```bash
# Tek kitap:
python pdf2word.py "Kitap.pdf"                     # Kitap.docx üretir
python pdf2word.py "Kitap.pdf" "Cikti.docx"

# Bir klasördeki bütün PDF'leri topluca:
python pdf2word.py "kitaplar_klasoru/"

# Seçenekler:
python pdf2word.py "Kitap.pdf" --font "Georgia" --size 12 --align left
python pdf2word.py "Kitap.pdf" --no-dehyphen      # metni birebir koru (tire birleştirme yok)
python pdf2word.py "Kitap.pdf" --no-headers       # üstbilgi/sayfa no temizliğini kapat
```

Çıktının sonunda, metin katmanı bozuk (gazete kupürü/şema) sayfalar varsa uyarır:
```
⚠ Metin katmanı bozuk görünen sayfalar (OCR gerekebilir): [32, 49, 85, 102, ...]
```

### Ne yapar
- Satırları doğru paragraflara toplar (**sayfa sınırlarını da aşarak**).
- Satır sonu tirelerini birleştirir; büyük harfle başlayan bileşiklerde tireyi
  korur (`Sovyet-` + `Rusya` → `Sovyet-Rusya`).
- Apostrof sonrası bölünen ekleri onarır (`Türkiye' nin` → `Türkiye'nin`).
- Tekrar eden üstbilgi/altbilgi ve salt sayfa numaralarını atar.
- Büyük puntolu başlıkları Word’de **Heading** olarak biçimler.
- Küçük puntolu **kenar notu/dipnotları** gövdeye karıştırmaz (ayrı paragraf).
- **Kelimeleri/cümleleri değiştirmez.**

---

## 🖥️ B) `webapp/` — kurulumsuz tarayıcı yolu

1. `webapp` klasörünü indirin (GitHub’da **Code → Download ZIP**).
2. İçindeki **`index.html`**’e çift tıklayın (Chrome/Edge). *(`vendor` klasörü yanında kalsın.)*
3. PDF’i sürükleyin → **“Word’e Dönüştür”** → `.docx` iner.

Kurulum ve internet gerektirmez; kitap bilgisayardan çıkmaz. Basit, dizgisi temiz
kitaplarda iyi çalışır. Yoğun tireleme içeren kitaplarda en iyi sonuç için
`pdf2word.py` kullanın.

---

## 🎯 Sınırlar (her iki araç için)

- **Dijital metinli** PDF’ler içindir (yazıyı fareyle seçebildiğiniz kitaplar).
- **Taranmış** kitaplar veya **gazete kupürü/şema** sayfaları için **OCR** gerekir;
  bu araçlar OCR yapmaz. `pdf2word.py` bu sayfaları raporlar. *(İstenirse OCR’lı
  sürüm eklenebilir.)*
- Çok sütunlu karmaşık düzenler ve tablolar tam oturmayabilir.

---

## 🧪 Testler

`tests/` klasöründe, tarayıcı uygulamasını gerçek bir tarayıcıda uçtan uca
doğrulayan otomatik testler var. Ayrıntı: [`tests/README.md`](tests/README.md).

```bash
cd tests && ./run_tests.sh
```

---

## 🔧 Teknik notlar

- `pdf2word.py`: **PyMuPDF** (metin + geometri, yumuşak tire korunur) +
  **python-docx** (çıktı). Paragraf birleştirme; tire onarımı; üstbilgi/altbilgi,
  başlık ve kenar-notu ayrımı.
- `webapp/`: **pdf.js** (yerel `vendor/`) + elle yazılmış OOXML+ZIP; tamamen istemci
  tarafı. pdf.js Apache-2.0’dır (bkz. [`webapp/vendor/README.md`](webapp/vendor/README.md)).
- Üretilen `.docx`, Microsoft Word / LibreOffice / Google Dokümanlar ile açılır.
