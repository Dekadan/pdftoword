# 📖 PDF → Word · Kitap Dönüştürücü

Kitap PDF’lerini, **satırları doğru paragraflara birleştirerek** düzenlenebilir
Word (`.docx`) dosyasına çeviren, tamamen tarayıcıda çalışan basit bir uygulama.

Yayınevinde 500–600 sayfalık kitapları Word’e aktarırken yaşanan iki dert için yapıldı:

1. **“Her satır ayrı oturuyor, paragraf oluşmuyor.”** Sıradan PDF→Word araçları
   metni satır satır döker; bu araç satırları okuyup **paragraf bütünlüğünü** kurar.
2. **“Aşırı kredi/ücret harcamadan yapmak istiyorum.”** Bu araç **hiç kredi/ücret
   harcamaz**: dönüştürmeyi yapay zekâ değil, sizin bilgisayarınızdaki program yapar.

---

## ✅ Neden bu araç?

- **Bedava ve sınırsız.** Bir kez indirdiniz mi, ister 1 kitap ister 100 kitap
  dönüştürün — masraf yok, kredi yok, abonelik yok.
- **Çevrimdışı ve gizli.** Dosyalar **hiçbir yere yüklenmez**, internet gerekmez.
  Kitap bilgisayarınızdan çıkmaz — telif açısından güvenli.
- **Kelimeleri değiştirmez.** Program metni **yeniden yazmaz**; yalnızca satırları
  doğru paragraflara yerleştirir, satır sonu tirelerini birleştirir ve tekrar eden
  üstbilgi/sayfa numaralarını temizler. Cümleler ve kelimeler olduğu gibi kalır.

---

## 🚀 Nasıl kullanılır (3 adım)

1. Bu depodaki **`webapp`** klasörünü bilgisayarınıza indirin.
   (GitHub’da yeşil **Code → Download ZIP** ile tümünü indirip açabilirsiniz.)
2. `webapp` klasöründeki **`index.html`** dosyasına **çift tıklayın**
   (Chrome veya Edge önerilir).
   > ⚠️ `vendor` klasörü `index.html` ile **aynı yerde** kalmalı — birlikte tutun.
3. PDF’i pencereye **sürükleyip bırakın** (ya da tıklayıp seçin) →
   **“Word’e Dönüştür”** → `.docx` dosyanız **kendiliğinden inecek**.

Birden çok kitabı aynı anda seçebilirsiniz; her biri ayrı Word dosyası olarak iner.

---

## ⚙️ Ayarlar ne işe yarar?

| Ayar | Açıklama |
|------|----------|
| **Yazı tipi / Punto** | Word dosyasının görünümü (Times New Roman 12 gibi). Metni etkilemez. |
| **Hizalama** | İki yana yasla (kitap görünümü) veya sola yasla. |
| **Satır aralığı** | Tek / 1,15 / 1,5 / çift. |
| **İlk satır girintisi** | Her paragrafın ilk satırı içeriden başlar (kitap görünümü). |
| **Satır sonu tirelerini birleştir** | `kelime-` + `nin` → `kelimenin`. Büyük harfle başlayan bileşiklerde (`Sovyet-` + `Rusya`) tire korunur. |
| **Üstbilgi/altbilgi ve sayfa no temizle** | Her sayfada tekrar eden başlık satırlarını ve salt sayfa numarasından oluşan satırları atar. |
| **Başlıkları algıla** | Gövdeden belirgin büyük puntolu kısa satırları Word’de “Başlık” (Heading) olarak biçimler. |

> **Metni birebir korumak isterseniz:** “Metin işleme” bölümündeki kutuları
> kapatın. O zaman tire birleştirme ve üstbilgi temizliği yapılmaz; yalnızca
> satırlar paragraf hâline getirilir.

---

## 🎯 Ne için uygundur, ne için değildir?

**Uygundur:**
- **Dijital metinli** PDF’ler — yani PDF’i açınca yazıyı fareyle **seçip
  kopyalayabildiğiniz** kitaplar. (Çoğu yeni/dizgisi yapılmış kitap böyledir.)
- Tek sütunlu, düz metin ağırlıklı kitaplar (roman, deneme, inceleme…).

**Şu an kapsam dışı / sınırlı:**
- **Taranmış (fotoğraf) PDF’ler.** Yazıyı seçemiyorsanız PDF taranmıştır; metni
  çıkarmak için **OCR** gerekir. Bu araç OCR yapmaz. (İhtiyaç olursa OCR’li bir
  sürüm ayrıca kurulabilir.)
- **Çok sütunlu sayfalar, karmaşık tablolar, kutu/kenar metinleri** düzgün akmayabilir.
- **Dipnotlar** gövde metnine karışabilir.
- **Tire istisnası:** Satır sonunda gerçekten tireli bir bileşik küçük harfle
  devam ediyorsa (nadiren) yanlışlıkla birleştirilebilir. Böyle kitaplarda
  “satır sonu tirelerini birleştir” seçeneğini kapatabilirsiniz.

Her dönüştürmeden sonra sayfanın altındaki **önizleme**den paragraf bütünlüğünü
hızlıca kontrol edebilirsiniz.

---

## 💡 İpuçları

- **Büyük kitaplar (500–600 sayfa):** Çift tıklayarak açtığınızda işlem birkaç
  dakika sürebilir; sekmeyi kapatmayın. Daha hızlı olması için isterseniz klasörü
  küçük bir yerel sunucuyla açabilirsiniz (geliştiriciler için aşağıya bakın).
- Sonuç beklediğiniz gibi değilse, ayarlarla oynayın (özellikle girinti ve
  tire seçenekleri farklı dizgilerde farklı sonuç verir).

---

## 🧪 Geliştiriciler için: testler

`tests/` klasöründe, aracın gerçek bir tarayıcıda uçtan uca doğrulandığı otomatik
testler var (PyMuPDF ile gerçekçi bir test PDF’i üretilir, Playwright ile
Chromium’da dönüştürülür, çıktı denetlenir). Ayrıntı: [`tests/README.md`](tests/README.md).

```bash
cd tests
./run_tests.sh
```

Hızlı yerel sunucu (büyük kitaplarda daha hızlı; gerçek pdf.js worker’ı devreye girer):

```bash
cd webapp
python3 -m http.server 8000
# tarayıcıda: http://localhost:8000/index.html
```

---

## 🔧 Teknik notlar

- **Tamamen istemci tarafı:** metin çıkarımı [pdf.js](https://mozilla.github.io/pdf.js/)
  (yerel `vendor/` altında), `.docx` üretimi ise elle yazılmış küçük bir OOXML + ZIP
  yazıcısıyla yapılır. Sunucu yok, ağ isteği yok.
- **Paragraf birleştirme mantığı:** satır konumları, yazı boyutları, satır aralıkları,
  paragraf girintileri ve cümle sonu noktalaması okunarak satırlar paragraflara toplanır.
- Üretilen `.docx`, Microsoft Word / LibreOffice / Google Dokümanlar tarafından açılır.
- pdf.js Apache-2.0 lisanslıdır; bkz. [`webapp/vendor/README.md`](webapp/vendor/README.md).
