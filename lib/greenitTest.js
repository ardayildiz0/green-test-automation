/**
 * GreenIT Metrikleri Toplama Modülü
 * Puppeteer ile sayfa yükleyip performance API üzerinden metrikleri toplar.
 * EcoIndex, CO2 ve su tüketimi hesaplar.
 *
 * Anti rate-limiting önlemleri:
 *   - Her ölçümde yeni incognito context (farklı oturum)
 *   - User-Agent rotasyonu (bot tespitini engelleme)
 *   - Random jitter (tahmin edilebilir istek kalıbını kırma)
 */

import { computeEcoIndex, getEcoIndexGrade, computeCO2, computeWater, formatDate, sleep } from './utils.js';

const PAGE_TIMEOUT = 30000; // 30 saniye
const MAX_RETRIES = 3;
const RETRY_DELAYS = [5000, 5000, 5000]; // her retry arası 5s
const BETWEEN_MEASUREMENT_DELAY = 2000; // ölçümler arası bekleme (ms)

// ═══════════════════════════════════════
// USER-AGENT HAVUZU
// ═══════════════════════════════════════
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:133.0) Gecko/20100101 Firefox/133.0',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.2 Safari/605.1.15',
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0',
  'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:133.0) Gecko/20100101 Firefox/133.0',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
];

function getRandomUA() {
  return USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
}

/** Rastgele jitter ekle: delay ± %30 */
function jitter(delayMs) {
  const variance = delayMs * 0.3;
  return Math.round(delayMs + (Math.random() * 2 - 1) * variance);
}

// Varsayılan logger — CLI modunda console'a yazar
const defaultLog = (msg, type = 'info') => {
  if (type === 'error') console.error(msg);
  else console.log(msg);
};

/**
 * Tek bir GreenIT ölçümü yap
 * @param {Object} sharedContext - Varsa bu context kullanılır (warm cache için), yoksa yeni incognito context açılır
 */
async function singleMeasurement(browser, url, coldCache, sharedContext = null) {
  // Cold cache: her ölçüm yeni incognito context → izole
  // Warm cache: paylaşımlı context → cache korunur
  const ownContext = !sharedContext;
  const context = sharedContext || await browser.createBrowserContext();
  const page = await context.newPage();

  try {
    // Rastgele User-Agent ayarla
    await page.setUserAgent(getRandomUA());
    await page.setViewport({ width: 1920, height: 1080 });

    if (coldCache) {
      const client = await page.createCDPSession();
      await client.send('Network.clearBrowserCache');
      await client.send('Network.clearBrowserCookies');
      await page.setCacheEnabled(false);
    } else {
      await page.setCacheEnabled(true);
    }

    let requestCount = 0;
    let totalTransferSize = 0;

    const client = await page.createCDPSession();
    await client.send('Network.enable');

    client.on('Network.responseReceived', () => { requestCount++; });
    client.on('Network.loadingFinished', (params) => {
      totalTransferSize += params.encodedDataLength || 0;
    });

    await page.goto(url, { waitUntil: 'networkidle2', timeout: PAGE_TIMEOUT });
    await sleep(500);

    const domSize = await page.evaluate(() => document.getElementsByTagName('*').length);

    const perfData = await page.evaluate(() => {
      const entries = performance.getEntriesByType('resource');
      let totalSize = 0;
      entries.forEach(entry => { totalSize += entry.transferSize || 0; });
      return { resourceCount: entries.length, resourceSize: totalSize };
    });

    const pageSizeBytes = Math.max(totalTransferSize, perfData.resourceSize);
    const pageSizeKB = Math.round(pageSizeBytes / 1024 * 100) / 100;
    const finalRequestCount = Math.max(requestCount, perfData.resourceCount);

    const ecoIndex = computeEcoIndex(domSize, finalRequestCount, pageSizeKB);
    const grade = getEcoIndexGrade(ecoIndex);
    const co2 = computeCO2(ecoIndex);
    const water = computeWater(ecoIndex);

    return { date: formatDate(), requestCount: finalRequestCount, pageSizeKB, domSize, co2, water, ecoIndex, grade };
  } finally {
    await page.close();
    // Sadece kendi oluşturduğumuz context'i kapat, paylaşımlıyı kapatma
    if (ownContext) await context.close();
  }
}

function makeErrorResult() {
  return { date: formatDate(), requestCount: 'ERROR', pageSizeKB: 'ERROR', domSize: 'ERROR', co2: 'ERROR', water: 'ERROR', ecoIndex: 'ERROR', grade: 'ERROR' };
}

/**
 * Cold cache testleri
 * @param {Object} opts - { onProgress, onLog, shouldStop }
 */
export async function runColdCacheTests(browser, url, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop } = typeof opts === 'function'
    ? { onProgress: opts } // Geriye uyumluluk: 4. parametre callback ise
    : opts;

  const results = [];

  for (let i = 0; i < count; i++) {
    // ── Durdurma kontrolü ──
    if (shouldStop && shouldStop()) {
      onLog(`Cold cache test ${i + 1}/${count} — DURDURULDU`, 'warning');
      break;
    }

    let lastError, result;

    for (let retry = 0; retry < MAX_RETRIES; retry++) {
      if (shouldStop && shouldStop()) break;
      try {
        result = await singleMeasurement(browser, url, true);
        break;
      } catch (err) {
        lastError = err;
        const delay = jitter(RETRY_DELAYS[retry] || 5000);
        onLog(`⚠ Cold #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        onLog(`⏳ ${(delay / 1000).toFixed(1)}s bekleniyor...`, 'info');
        await sleep(delay);
      }
    }

    if (result) {
      results.push({ measurementNo: i + 1, ...result });
      onLog(`✓ Cold #${i + 1}/${count} — EcoIndex: ${result.ecoIndex} (${result.grade}) | Req: ${result.requestCount} | Size: ${result.pageSizeKB}KB | DOM: ${result.domSize}`, 'success');
    } else {
      results.push({ measurementNo: i + 1, ...makeErrorResult() });
      onLog(`✗ Cold #${i + 1}/${count} HATA: ${lastError?.message}`, 'error');
    }

    if (onProgress) onProgress(i + 1, count);
    if (i < count - 1) await sleep(jitter(BETWEEN_MEASUREMENT_DELAY));
  }

  return results;
}

/**
 * Warm cache testleri — tek paylaşımlı context ile (cache korunur)
 * @param {Object} opts - { onProgress, onLog, shouldStop }
 */
export async function runWarmCacheTests(browser, url, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop } = typeof opts === 'function'
    ? { onProgress: opts }
    : opts;

  const results = [];

  // Tüm warm ölçümler için TEK paylaşımlı context — cache korunsun
  const warmContext = await browser.createBrowserContext();
  const ua = getRandomUA();

  try {
    // Warmup (retry ile)
    onLog('Warm cache warmup yükleniyor...', 'info');
    let warmupDone = false;
    for (let wr = 0; wr < 3 && !warmupDone; wr++) {
      const warmupPage = await warmContext.newPage();
      try {
        await warmupPage.setUserAgent(ua);
        await warmupPage.setCacheEnabled(true);
        await warmupPage.goto(url, { waitUntil: 'networkidle2', timeout: PAGE_TIMEOUT });
        await sleep(500);
        onLog('Warmup tamamlandı, cache hazır.', 'info');
        warmupDone = true;
      } catch (err) {
        const delay = jitter(RETRY_DELAYS[wr] || 5000);
        onLog(`⚠ Warmup hatası (deneme ${wr + 1}/3): ${err.message} — ${(delay / 1000).toFixed(1)}s bekleniyor`, 'warning');
        await sleep(delay);
      } finally {
        await warmupPage.close();
      }
    }

    for (let i = 0; i < count; i++) {
      if (shouldStop && shouldStop()) {
        onLog(`Warm cache test ${i + 1}/${count} — DURDURULDU`, 'warning');
        break;
      }

      let lastError, result;

      for (let retry = 0; retry < MAX_RETRIES; retry++) {
        if (shouldStop && shouldStop()) break;
        try {
          // Paylaşımlı context'i geç → cache korunur
          result = await singleMeasurement(browser, url, false, warmContext);
          break;
        } catch (err) {
          lastError = err;
          const delay = jitter(RETRY_DELAYS[retry] || 5000);
          onLog(`⚠ Warm #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
          onLog(`⏳ ${(delay / 1000).toFixed(1)}s bekleniyor...`, 'info');
          await sleep(delay);
        }
      }

      if (result) {
        results.push({ measurementNo: i + 1, ...result });
        onLog(`✓ Warm #${i + 1}/${count} — EcoIndex: ${result.ecoIndex} (${result.grade}) | Req: ${result.requestCount} | Size: ${result.pageSizeKB}KB | DOM: ${result.domSize}`, 'success');
      } else {
        results.push({ measurementNo: i + 1, ...makeErrorResult() });
        onLog(`✗ Warm #${i + 1}/${count} HATA: ${lastError?.message}`, 'error');
      }

      if (onProgress) onProgress(i + 1, count);
      if (i < count - 1) await sleep(jitter(BETWEEN_MEASUREMENT_DELAY));
    }
  } finally {
    // Tüm warm ölçümler bittikten sonra context'i kapat
    await warmContext.close();
  }

  return results;
}
