# GreenIT + Lighthouse Otomasyon Scripti

SENG434 - Green Software Engineering dersi için otomatik web analiz aracı. Kurumlara ait web sitelerine GreenIT (EcoIndex, CO2, su tüketimi) ve Lighthouse (FCP, LCP, TBT, CLS, Speed Index) testleri uygular, sonuçları kurum bazında Excel dosyalarına ve DOCX raporlarına kaydeder.

## Kurulum

```bash
cd green-test-automation
npm install
```

## Kullanım

### 1. Web Arayüzü (Önerilen)

```bash
npm run ui
```

Tarayıcıda `http://localhost:3000` adresini açın.

Arayüz özellikleri:
- Tekli veya toplu URL ekleme
- GreenIT ve Lighthouse ölçüm sayılarını ayrı ayarlama
- Gerçek zamanlı ilerleme takibi (Cold Cache / Warm Cache / Lighthouse aşamaları)
- Sonuç tablosu ve log görüntüleme
- Testleri istediğiniz an durdurma

### 2. CLI (Komut Satırı)

#### Tek URL
```bash
node index.js --url "https://www.tedas.gov.tr" --name "TEDAŞ"
```

#### Çoklu URL (Dosyadan)
```bash
node index.js --file urls.txt --output ./results
```

### CLI Parametreleri

| Parametre | Açıklama | Varsayılan |
|-----------|----------|------------|
| `--url` | Test edilecek tek URL | - |
| `--name` | Kurum adı (`--url` ile zorunlu) | - |
| `--file` | URL listesi dosyası | - |
| `--greenit-count` | GreenIT ölçüm sayısı (Cold + Warm) | 20 |
| `--lighthouse-count` | Lighthouse ölçüm sayısı | 10 |
| `--count` | Hepsini aynı anda ayarla | - |
| `--output` | Sonuç dosyalarının kaydedileceği klasör | ./results |

### URL Dosya Formatı (`urls.txt`)
```
# Yorum satırları # ile başlar
KURUM_ADI|URL
TEDAŞ|https://www.tedas.gov.tr
AFAD|https://www.afad.gov.tr
Adalet Bakanlığı|https://www.adalet.gov.tr
```

## Çıktı

Her kurum için iki dosya oluşturulur:

### Excel Dosyası (`KURUM_ADI_results.xlsx`)

| Sheet | İçerik | Ölçüm |
|-------|--------|-------|
| Cold Cache GreenIT | Request Count, Page Size, DOM Size, CO2, Water, EcoIndex, Grade | 20x |
| Warm Cache GreenIT | Aynı metrikler (cache'li) | 20x |
| Lighthouse Metrics | FCP, LCP, TBT, CLS, Speed Index | 10x |
| Summary | Ortalama Grade (renk kodlu), not aralıkları tablosu, tüm metriklerin ortalamaları | - |

### DOCX Rapor (`KURUM_ADI_rapor.docx`)

Writing Guideline şablonuna uygun Word raporu otomatik oluşturulur:

| Bölüm | İçerik |
|-------|--------|
| 1. ÖZET | Test özeti ve temel sonuçlar |
| 2. ÇALIŞMANIN AMACI VE KAPSAMI | Amaç, kapsam ve ölçülen metrikler |
| 3. KULLANILAN ARAÇLAR | GreenIT Analysis ve Lighthouse açıklamaları |
| 4. Metodoloji | Test yöntemi ve konfigürasyon detayları |
| 5. ÖLÇÜM SONUÇLARI | 4 tablo: Cold/Warm Cache GreenIT, Lighthouse, Özet |
| 6. SONUÇLARIN ANALİZİ | Sonuçlara göre otomatik üretilen 6 alt bölüm analiz |
| 7. İYİLEŞTİRME ÖNERİLERİ | Sonuçlara göre otomatik üretilen 7 öneri |
| 8. SONUÇ | Genel değerlendirme ve öneriler |

## EcoIndex Not Aralıkları

Summary sheet'te ortalama EcoIndex değerine göre harf notu otomatik hesaplanır ve renk kodlanır:

| Not | Puan Aralığı | Renk |
|-----|-------------|------|
| **A** | 81 – 100 | 🟢 Yeşil |
| **B** | 71 – 80 | 🟢 Açık Yeşil |
| **C** | 56 – 70 | 🟡 Sarı-Yeşil |
| **D** | 41 – 55 | 🟡 Sarı |
| **E** | 26 – 40 | 🟠 Turuncu |
| **F** | 11 – 25 | 🔴 Kırmızı-Turuncu |
| **G** | 0 – 10 | 🔴 Kırmızı |

Kaynak: [EcoIndex.fr](https://www.ecoindex.fr/en/how-it-works/) / [cnumr/GreenIT-Analysis](https://github.com/cnumr/GreenIT-Analysis)

## Proje Yapısı

```
green-test-automation/
├── package.json
├── index.js              # CLI giriş noktası
├── server.js             # Web arayüzü sunucusu
├── public/
│   └── index.html        # Web arayüzü
├── lib/
│   ├── utils.js          # EcoIndex formülü, CLI parser, yardımcılar
│   ├── greenitTest.js    # Puppeteer ile GreenIT metrikleri (cold/warm)
│   ├── lighthouseTest.js # Lighthouse programmatic API
│   ├── excelWriter.js    # ExcelJS ile kurum bazında rapor
│   └── reportWriter.js   # DOCX rapor oluşturma (docx-js)
├── template/
│   └── template.xlsx     # Orijinal Excel şablonu
└── urls-example.txt      # Örnek URL dosyası
```

## Gereksinimler

- Node.js 18+
- Chrome/Chromium (puppeteer otomatik indirir)
