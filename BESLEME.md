# RaporAI veri beslemesi

Bu klasör, updater'ın `https://raw.githubusercontent.com/batu69krk-collab/raporai-data/main/data/manifest.json`
adresinden beklediği düzenin birebir kopyasıdır (üretim: `node scripts/besleme-hazirla.js`).

Yayımlama (GitHub hesabında, depo adı birebir önemli):
1. github.com'da **batu69krk-collab/raporai-data** deposunu **Public** aç (raw erişimi anahtarsız olmalı).
2. Bu klasörün içeriğini deponun köküne it:

   cd veri-besleme
   git init -b main
   git add data
   git commit -m "veri paketi 2026-09-12.5"
   git remote add origin https://github.com/batu69krk-collab/raporai-data.git
   git push -u origin main

3. Doğrula: tarayıcıda manifest URL'si 200 dönmeli; uygulama içinde
   "Güncelleme kontrol" bir sonraki koşumda çalışmalı (sürüm düşürme koruması
   aynı paket sürümünde no-op'tur).

SÜRÜM KİLİDİ NOTU (2026-09-15): updater yalnız DAHA BÜYÜK paketSurum'u güncelleme
olarak önerir; eşit sürüm "güncel" sayılır. Bu yüzden içerik baytları değiştiğinde
`manifestUret` aynı gün etiketinde .N revizyonunu otomatik artırır
(`paketSurumuIlerlet`; test: `npm run test:paket-surum`). Etiketi elle zorlamak
gerekirse `RAPORAI_PAKET_SURUM` env'i otomatik kilidi ezer — ama aynı etiketle
farklı içerik YAYIMLAMAYIN: kurulu cihazlar yeni paketi hiç görmez.

Her yeni veri paketinde: `npm run paket:uret && npm run paket:v3` →
`npm run build && node scripts/paket-sifrele.js` → `node scripts/besleme-hazirla.js` → it.
