/**
 * Ortak yardımcı fonksiyonlar
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

/**
 * EcoIndex hesaplama (resmi formül)
 * Kaynak: https://www.ecoindex.fr/comment-ca-marche/
 *
 * EcoIndex = 100 - (5 * (3*q(dom) + 2*q(req) + q(size)) / 6)
 * Quantile tablosu kullanılır
 */

// EcoIndex quantile tabloları
const QUANTILES_DOM = [0, 47, 75, 159, 233, 298, 358, 417, 476, 537, 603, 674, 753, 843, 949, 1076, 1237, 1459, 1801, 2479, 594601];
const QUANTILES_REQ = [0, 2, 15, 25, 34, 42, 49, 56, 63, 70, 78, 86, 95, 105, 117, 130, 147, 170, 205, 281, 3920];
const QUANTILES_SIZE = [0, 1.37, 144.7, 319.53, 479.46, 631.97, 783.38, 937.91, 1098.62, 1265.47, 1448.32, 1648.27, 1876.08, 2142.06, 2465.37, 2866.31, 3401.59, 4155.73, 5400.08, 8037.54, 223212.26];

function getQuantileIndex(quantiles, value) {
  for (let i = 1; i < quantiles.length; i++) {
    if (value < quantiles[i]) {
      return i - 1 + (value - quantiles[i - 1]) / (quantiles[i] - quantiles[i - 1]);
    }
  }
  return quantiles.length - 1;
}

export function computeEcoIndex(dom, req, sizeKB) {
  const qDom = getQuantileIndex(QUANTILES_DOM, dom);
  const qReq = getQuantileIndex(QUANTILES_REQ, req);
  const qSize = getQuantileIndex(QUANTILES_SIZE, sizeKB);

  const ecoIndex = 100 - (5 * (3 * qDom + 2 * qReq + qSize) / 6);
  return Math.max(0, Math.min(100, Math.round(ecoIndex * 100) / 100));
}

export function getEcoIndexGrade(ecoIndex) {
  if (ecoIndex > 80) return 'A';
  if (ecoIndex > 70) return 'B';
  if (ecoIndex > 55) return 'C';
  if (ecoIndex > 40) return 'D';
  if (ecoIndex > 25) return 'E';
  if (ecoIndex > 10) return 'F';
  return 'G';
}

/**
 * EcoIndex Not Aralıkları (Resmi)
 * Kaynak: https://www.ecoindex.fr / cnumr/GreenIT-Analysis
 *
 *   A: 81 – 100
 *   B: 71 – 80
 *   C: 56 – 70
 *   D: 41 – 55
 *   E: 26 – 40
 *   F: 11 – 25
 *   G:  0 – 10
 */
export const GRADE_RANGES = [
  { grade: 'A', min: 81, max: 100, color: '4CAF50' },
  { grade: 'B', min: 71, max: 80,  color: '8BC34A' },
  { grade: 'C', min: 56, max: 70,  color: 'CDDC39' },
  { grade: 'D', min: 41, max: 55,  color: 'FFEB3B' },
  { grade: 'E', min: 26, max: 40,  color: 'FF9800' },
  { grade: 'F', min: 11, max: 25,  color: 'FF5722' },
  { grade: 'G', min: 0,  max: 10,  color: 'F44336' }
];

/**
 * CO2 emisyon hesaplama (gCO2e)
 * Resmi formül: 2 + (ecoIndex / 100) * 2
 * Yüksek ecoIndex → düşük CO2 (ters orantılı)
 * Kaynak: https://github.com/cnumr/GreenIT-Analysis
 */
export function computeCO2(ecoIndex) {
  return Math.round((2 + ecoIndex * 0.01) * 100) / 100;
}

/**
 * Su tüketimi hesaplama (cl)
 * Resmi formül: 3 + (ecoIndex / 100) * 3
 * Kaynak: https://github.com/cnumr/GreenIT-Analysis
 */
export function computeWater(ecoIndex) {
  return Math.round((3 + ecoIndex * 0.03) * 100) / 100;
}

/**
 * CLI argümanlarını parse et
 */
export function parseArgs(args) {
  // Önce config.json'dan varsayılan değerleri oku
  let config = {};
  const configPath = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'config.json');
  try {
    if (fs.existsSync(configPath)) {
      config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
    }
  } catch (e) {
    // config.json yoksa veya hatalıysa varsayılanları kullan
  }

  const parsed = {
    url: null,
    name: null,
    file: null,
    author: config.author || 'Arda Yıldız',
    greenitCount: config.greenitCount || 20,
    lighthouseCount: config.lighthouseCount || 10,
    output: config.output || './results'
  };

  for (let i = 2; i < args.length; i++) {
    switch (args[i]) {
      case '--url':
        parsed.url = args[++i];
        break;
      case '--name':
        parsed.name = args[++i];
        break;
      case '--file':
        parsed.file = args[++i];
        break;
      case '--author':
        parsed.author = args[++i];
        break;
      case '--greenit-count':
        parsed.greenitCount = parseInt(args[++i], 10);
        break;
      case '--lighthouse-count':
        parsed.lighthouseCount = parseInt(args[++i], 10);
        break;
      case '--count':
        // Hepsini aynı anda ayarla
        const val = parseInt(args[++i], 10);
        parsed.greenitCount = val;
        parsed.lighthouseCount = val;
        break;
      case '--output':
        parsed.output = args[++i];
        break;
      case '--help':
        printHelp();
        process.exit(0);
    }
  }

  return parsed;
}

export function printHelp() {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║        GreenIT + Lighthouse Otomasyon Scripti               ║
║        SENG434 - Green Software Engineering                 ║
╚══════════════════════════════════════════════════════════════╝

Kullanım:
  node index.js --url <URL> --name <KURUM_ADI> [seçenekler]
  node index.js --file <DOSYA> [seçenekler]

Seçenekler:
  --url <URL>         Test edilecek tek URL
  --name <AD>         Kurum adı (--url ile birlikte)
  --file <DOSYA>      URL listesi dosyası (her satır: KURUM_ADI|URL)
  --greenit-count <N>     GreenIT ölçüm sayısı (varsayılan: 20)
  --lighthouse-count <N>  Lighthouse ölçüm sayısı (varsayılan: 10)
  --count <N>             Hepsini aynı anda ayarla
  --author <AD>       Rapor hazırlayan adı (varsayılan: config.json'dan)
  --output <KLASÖR>   Çıktı klasörü (varsayılan: ./results)
  --help              Bu yardım mesajını göster

Örnekler:
  node index.js --url "https://www.tedas.gov.tr" --name "TEDAŞ" --count 20
  node index.js --file urls.txt --count 20 --output ./sonuclar

URL Dosya Formatı (urls.txt):
  TEDAŞ|https://www.tedas.gov.tr
  AFAD|https://www.afad.gov.tr
  Adalet Bakanlığı|https://www.adalet.gov.tr
`);
}

/**
 * Progress bar göster
 */
export function showProgress(current, total, label) {
  const percent = Math.round((current / total) * 100);
  const filled = Math.round(percent / 2);
  const empty = 50 - filled;
  const bar = '█'.repeat(filled) + '░'.repeat(empty);
  process.stdout.write(`\r  [${bar}] ${percent}% - ${label} (${current}/${total})`);
  if (current === total) process.stdout.write('\n');
}

/**
 * Tarih formatla
 */
export function formatDate() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
}

/**
 * URL'yi normalize et
 */
export function normalizeUrl(url) {
  if (!url) return url;
  url = url.trim();
  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    url = 'https://' + url;
  }
  return url;
}

/**
 * Dosya adı için güvenli string
 */
export function sanitizeFilename(name) {
  return name
    .replace(/[\/\\?%*:|"<>]/g, '-')
    .replace(/\s+/g, '_')
    .trim();
}

/**
 * Bekleme fonksiyonu
 */
export function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
