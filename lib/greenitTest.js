/**
 * GreenIT Metrikleri Toplama Modülü
 * Puppeteer ile sayfa yükleyip performance API üzerinden metrikleri toplar.
 * EcoIndex, CO2 ve su tüketimi hesaplar.
 */

import { computeEcoIndex, getEcoIndexGrade, computeCO2, computeWater, formatDate, sleep } from './utils.js';

const PAGE_TIMEOUT = 60000; // 60 saniye
const MAX_RETRIES = 3;

// Varsayılan logger — CLI modunda console'a yazar
const defaultLog = (msg, type = 'info') => {
  if (type === 'error') console.error(msg);
  else console.log(msg);
};

/**
 * Tek bir GreenIT ölçümü yap
 */
async function singleMeasurement(browser, url, coldCache) {
  const page = await browser.newPage();

  try {
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
        onLog(`⚠ Cold #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        await sleep(1000);
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
    if (i < count - 1) await sleep(300);
  }

  return results;
}

/**
 * Warm cache testleri
 * @param {Object} opts - { onProgress, onLog, shouldStop }
 */
export async function runWarmCacheTests(browser, url, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop } = typeof opts === 'function'
    ? { onProgress: opts }
    : opts;

  const results = [];

  // Warmup
  onLog('Warm cache warmup yükleniyor...', 'info');
  const warmupPage = await browser.newPage();
  try {
    await warmupPage.setCacheEnabled(true);
    await warmupPage.goto(url, { waitUntil: 'networkidle2', timeout: PAGE_TIMEOUT });
    await sleep(500);
    onLog('Warmup tamamlandı, cache hazır.', 'info');
  } catch (err) {
    onLog(`⚠ Warmup hatası: ${err.message}`, 'warning');
  } finally {
    await warmupPage.close();
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
        result = await singleMeasurement(browser, url, false);
        break;
      } catch (err) {
        lastError = err;
        onLog(`⚠ Warm #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        await sleep(1000);
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
    if (i < count - 1) await sleep(300);
  }

  return results;
}
