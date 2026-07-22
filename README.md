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

## Üç kullanım yolu — hangisi?

| | 🟢 **Uygulama** (indir-çalıştır) | **`pdf2word.py`** (Python) | **`webapp/`** (tarayıcı) |
|---|---|---|---|
| Kimin için | **Herkes** (önerilen) | Geliştirici / otomasyon | Hızlı deneme |
| Kurulum | ❌ Yok — çift tıkla | Python gerekir | ❌ Yok — HTML aç |
| Kalite (tire birleştirme) | ⭐ En iyi | ⭐ En iyi | ⚠️ Sınırlı (bkz. not) |
| Toplu (çok kitap) | ✅ | ✅ | Tek tek |
| Maliyet / gizlilik | Bedava, çevrimdışı | Bedava, çevrimdışı | Bedava, çevrimdışı |

> **Not — neden tarayıcı sürümü daha zayıf?** Türkçe kitaplarda satır sonları
> çoğunlukla *yumuşak tire* (görünmez tire) ile bölünür. Tarayıcıdaki pdf.js bu
> tireyi **siler**, o yüzden `dışlaya` ile `rak` arasına yanlış boşluk girer
> (`dışlaya rak`). **Uygulama** ve **Python** sürümleri PyMuPDF kullanır, tireyi
> **korur** ve kelimeyi doğru birleştirir (`dışlayarak`). Gerçek kitaplar için
> **uygulamayı** kullanın.

---

## ⬇️ En kolay yol: uygulamayı indir, aç, kullan

Kurulum yok, Python yok, komut satırı yok. Çift tıkla açılan gerçek bir program.

1. Deponun **[Releases](../../releases)** sayfasına gidin.
2. Bilgisayarınıza uygun dosyayı indirin:
   - **Windows:** `KitapDonusturucu-Windows.zip` → açın → `KitapDonusturucu.exe`’ye çift tıklayın.
     *(İlk açılışta “Windows PC’nizi korudu” çıkarsa: **Ek bilgi → Yine de çalıştır**.)*
   - **Mac:** `KitapDonusturucu-macOS.zip` → açın → uygulamaya **sağ tıklayın → Aç**.
3. Açılan pencerede **PDF Seç** → **Word’e Çevir**. Word dosyası PDF’in yanına kaydedilir.

Birden çok kitabı aynı anda seçebilirsiniz. Her şey bilgisayarınızda, çevrimdışı çalışır.

> Uygulama, GitHub tarafından otomatik derlenir (`.github/workflows/build-app.yml`).
> Yeni sürüm çıkarmak için bir etiket itmek yeterli: `git tag v0.1.0 && git push --tags`.
> İmzasız olduğu için Windows SmartScreen / Mac Gatekeeper ilk açılışta uyarabilir; yukarıdaki adımlarla açılır.

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
- Satırları doğru paragraflara toplar (**sayfa sınırlarını da aşarak**);
  paragraf araları boşluklu (`--para-space`, varsayılan 6 punto).
- Satır sonu tirelerini birleştirir; büyük harfle başlayan bileşiklerde tireyi
  korur (`Sovyet-` + `Rusya` → `Sovyet-Rusya`).
- Apostrof sonrası bölünen ekleri onarır (`Türkiye' nin` → `Türkiye'nin`).
- **DİPNOTLAR:** gövdedeki işaret rakamlarıyla sayfa altındaki notları eşleştirip
  **gerçek Word dipnotu** yapar (sayfa altında, otomatik numaralı). Eşleşmeyen
  not kaybolmaz, küçük italik paragraf olarak kalır.
- **İÇİNDEKİLER:** PDF’teki içindekiler sayfalarını söker, yerine Word’ün
  **canlı İçindekiler alanını** koyar — açılışta başlıklardan kendiliğinden
  dolar, sayfa numaralarını Word doğru basar.
- **OCR:** metin katmanı bozuk sayfaları (gazete kupürü/şema) ve tümüyle
  **taranmış kitapları** görüntüden okur (Türkçe; `--ocr auto|full|off`).
- Altbilgiye **sayfa numarası** alanı; A4 kâğıt + 2,5 cm kitap marjları.
- Tekrar eden üstbilgi/altbilgi ve salt sayfa numaralarını atar.
- Başlıkları tanır (`BÖLÜM`, BÜYÜK HARF, `V.2.` desenleri) → Heading 1-3.
- Küçük puntolu **kenar notlarını** gövdeyi bölmeden ayırır.
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

## 🎯 Sınırlar

- **Uygulama ve Python sürümü** taranmış kitapları ve bozuk sayfaları **OCR ile**
  okur (Türkçe dil dosyası gömülü). OCR sonucu, taramanın kalitesine bağlıdır —
  düşük çözünürlüklü gazete kupürlerinde hata kalabilir; araç bu sayfaları raporlar.
- Tarayıcı sürümü (webapp) OCR yapmaz; dijital metinli PDF’ler içindir.
- Çok sütunlu karmaşık düzenler ve tablolar tam oturmayabilir.
- Word, metni yeniden akıttığı için sayfalar PDF ile birebir aynı yerde bölünmez;
  sayfa numaraları ve içindekiler **Word’ün kendi düzenine göre** doğrudur
  (baskı öncesi taslakta istenen de budur).

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
