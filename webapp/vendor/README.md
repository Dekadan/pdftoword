# vendor/

Bu klasör, uygulamanın **çevrimdışı** çalışabilmesi için gereken üçüncü taraf
kütüphaneyi içerir. `index.html` ile **aynı yerde** kalmalıdır.

## pdf.js

- Dosyalar: `pdf.min.js`, `pdf.worker.min.js`
- Sürüm: **3.11.174** (`pdfjs-dist` paketinin `legacy` yapısı — eski tarayıcılarla da uyumlu)
- Kaynak: <https://github.com/mozilla/pdf.js>
- Lisans: **Apache License 2.0** — <https://github.com/mozilla/pdf.js/blob/master/LICENSE>

PDF’lerden metin çıkarımı için kullanılır. Herhangi bir ağ isteği yapmaz;
tümüyle tarayıcıda, yerelde çalışır.
