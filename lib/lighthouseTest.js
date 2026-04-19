/**
 * Lighthouse Test Modülü
 * Lighthouse programmatic API ile performans metriklerini toplar.
 * FCP, LCP, TBT, CLS, Speed Index
 *
 * Cloudflare / bot korumalı siteler için alternatif:
 *   runLighthouseViaPSI — Google PageSpeed Insights API üzerinden ölçüm yapar.
 *   Google'ın sunucuları Cloudflare tarafından engellenmez.
 */

import https from 'https';
import http from 'http';
import { formatDate, sleep } from './utils.js';

const MAX_RETRIES = 3;

// ═══════════════════════════════════════
// CLOUDFLARE TESPİTİ
// ═══════════════════════════════════════

/**
 * Sitenin bot koruması veya erişim engeli olup olmadığını tespit eder.
 * Tarayıcısız (curl) istek atılır — koruma varsa curl'e karşı tepki verir.
 *
 * PSI'ya yönlendirme kriterleri (biri yeterliyse):
 *   - Cloudflare challenge (cf-ray + cf-mitigated / challenge sayfası)
 *   - HTTP 403 Forbidden (Cloudflare olmayan bot engellemesi)
 *   - HTTP 503 (geçici erişim engeli)
 */
export function detectCloudflare(url) {
  return new Promise((resolve) => {
    const parsed = new URL(url);
    const lib = parsed.protocol === 'https:' ? https : http;

    const req = lib.get(url, {
      headers: {
        'User-Agent': 'curl/7.88.1',
        'Accept': '*/*'
      },
      timeout: 10000
    }, (res) => {
      const status       = res.statusCode;
      const hasCfRay     = !!res.headers['cf-ray'];
      const isChallenged = res.headers['cf-mitigated'] === 'challenge';

      // Generic 403/503 → Lighthouse da aynı sonucu alır, PSI kullan
      if (status === 403 || status === 503) {
        res.resume(); // body'yi tüket ve kapat
        resolve(true);
        return;
      }

      let body = '';
      res.on('data', chunk => { body += chunk.toString(); });
      res.on('end', () => {
        const hasChallengePage = body.includes('_cf_chl_opt')
                              || body.includes('Just a moment')
                              || body.includes('cf-browser-verification')
                              || body.includes('challenges.cloudflare.com');
        resolve(hasCfRay && (isChallenged || hasChallengePage));
      });
    });

    req.on('error', () => resolve(false));
    req.on('timeout', () => { req.destroy(); resolve(false); });
  });
}

/**
 * Otomatik Lighthouse: önce Cloudflare tespiti yapar, koruma varsa PSI kullanır.
 * @param {string} url
 * @param {number} port       - Yerel Chromium debug portu
 * @param {number} count
 * @param {Object} opts       - { onProgress, onLog, shouldStop, apiKey }
 * @returns {{ results, usedPSI }}
 */
export async function runLighthouseAuto(url, port, count, opts = {}) {
  const { onLog = defaultLog, apiKey = '', forcePSI = false } = opts;

  if (forcePSI) {
    onLog('🌐 PSI modu aktif → PageSpeed Insights API kullanılıyor.', 'info');
    const results = await runLighthouseViaPSI(url, count, { ...opts, apiKey });
    return { results, usedPSI: true };
  }

  onLog('🔍 Bot koruması kontrol ediliyor...', 'info');
  const protected_ = await detectCloudflare(url);

  if (protected_) {
    onLog('🛡️  Bot koruması / erişim engeli tespit edildi → PageSpeed Insights API kullanılıyor.', 'warning');
    const results = await runLighthouseViaPSI(url, count, { ...opts, apiKey });
    return { results, usedPSI: true };
  } else {
    onLog('✅ Engel tespit edilmedi → Lighthouse yerel olarak çalışıyor.', 'info');
    const results = await runLighthouseTests(url, port, count, opts);
    return { results, usedPSI: false };
  }
}

const defaultLog = (msg, type = 'info') => {
  if (type === 'error') console.error(msg);
  else console.log(msg);
};

/**
 * Tek bir Lighthouse ölçümü yap (verilen port üzerinden)
 */
async function singleLighthouseRun(url, port) {
  const lighthouse = (await import('lighthouse')).default;

  const throttling = {
    rttMs: 40, throughputKbps: 10240, cpuSlowdownMultiplier: 1,
    requestLatencyMs: 0, downloadThroughputKbps: 0, uploadThroughputKbps: 0
  };
  const screenEmulation = {
    mobile: false, width: 1920, height: 1080,
    deviceScaleFactor: 1, disabled: false
  };

  const result = await lighthouse(url, {
    port,
    output: 'json',
    onlyCategories: ['performance'],
    formFactor: 'desktop',
    screenEmulation,
    throttling,
    disableStorageReset: false
  }, {
    extends: 'lighthouse:default',
    settings: { formFactor: 'desktop', throttling, screenEmulation }
  });

  const audits = result.lhr.audits;
  return {
    date: formatDate(),
    fcp: Math.round(audits['first-contentful-paint']?.numericValue || 0),
    lcp: Math.round(audits['largest-contentful-paint']?.numericValue || 0),
    tbt: Math.round(audits['total-blocking-time']?.numericValue || 0),
    cls: Math.round((audits['cumulative-layout-shift']?.numericValue || 0) * 1000) / 1000,
    speedIndex: Math.round(audits['speed-index']?.numericValue || 0)
  };
}

/**
 * Lighthouse testleri — her URL için ayrı, izole Chrome instance'ı kullanır.
 * Puppeteer browser'ı ile port paylaşımı yapılmaz → sıfır sonuç sorunu çözülür.
 * @param {Object} opts - { onProgress, onLog, shouldStop }
 */
export async function runLighthouseTests(url, _unusedPort, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop } = typeof opts === 'function'
    ? { onProgress: opts }
    : opts;

  const { launch } = await import('chrome-launcher');
  const results = [];

  for (let i = 0; i < count; i++) {
    if (shouldStop && shouldStop()) {
      onLog(`Lighthouse test ${i + 1}/${count} — DURDURULDU`, 'warning');
      break;
    }

    let lastError, result;

    for (let retry = 0; retry < MAX_RETRIES; retry++) {
      if (shouldStop && shouldStop()) break;

      // Her ölçüm için temiz, izole bir Chrome başlat → birikmiş state sorunu çözülür
      const chrome = await launch({
        chromeFlags: ['--headless=new', '--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage', '--disable-extensions']
      });
      try {
        result = await singleLighthouseRun(url, chrome.port);
        break;
      } catch (err) {
        lastError = err;
        onLog(`⚠ Lighthouse #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        await sleep(1000);
      } finally {
        try { await chrome.kill(); } catch {}
      }
    }

    if (result) {
      results.push({ measurementNo: i + 1, ...result });
      onLog(`✓ Lighthouse #${i + 1}/${count} — FCP: ${result.fcp}ms | LCP: ${result.lcp}ms | TBT: ${result.tbt}ms | CLS: ${result.cls} | SI: ${result.speedIndex}`, 'success');
    } else {
      results.push({ measurementNo: i + 1, date: formatDate(), fcp: 'ERROR', lcp: 'ERROR', tbt: 'ERROR', cls: 'ERROR', speedIndex: 'ERROR' });
      onLog(`✗ Lighthouse #${i + 1}/${count} HATA: ${lastError?.message}`, 'error');
    }

    if (onProgress) onProgress(i + 1, count);
    if (i < count - 1) await sleep(500);
  }

  return results;
}

// ═══════════════════════════════════════
// PAGESPEED INSIGHTS API (Cloudflare bypass)
// ═══════════════════════════════════════

/**
 * Google PSI API'den tek ölçüm al
 * Google'ın sunucuları siteye erişir → Cloudflare engeli yok
 */
function fetchPSI(url, apiKey) {
  return new Promise((resolve, reject) => {
    // PSI sonuçlarını cache'lemeyi önlemek için URL'ye benzersiz fragment ekle.
    // PSI bunu URL'nin parçası olarak değerlendirip sayfayı yeniden ölçer.
    const cacheBust = `#psi_cb=${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const encoded = encodeURIComponent(url + cacheBust);
    const keyParam = apiKey ? `&key=${apiKey}` : '';
    const apiUrl = `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encoded}&strategy=desktop${keyParam}`;

    https.get(apiUrl, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.error) {
            reject(new Error(`PSI API hatası: ${json.error.message}`));
            return;
          }
          const audits = json.lighthouseResult?.audits;
          if (!audits) {
            reject(new Error('PSI yanıtında lighthouseResult bulunamadı'));
            return;
          }
          resolve({
            date: formatDate(),
            fcp:        Math.round(audits['first-contentful-paint']?.numericValue || 0),
            lcp:        Math.round(audits['largest-contentful-paint']?.numericValue || 0),
            tbt:        Math.round(audits['total-blocking-time']?.numericValue || 0),
            cls:        Math.round((audits['cumulative-layout-shift']?.numericValue || 0) * 1000) / 1000,
            speedIndex: Math.round(audits['speed-index']?.numericValue || 0)
          });
        } catch (e) {
          reject(new Error(`PSI yanıtı parse edilemedi: ${e.message}`));
        }
      });
    }).on('error', reject);
  });
}

/**
 * PageSpeed Insights API ile Lighthouse ölçümleri
 * API key olmadan çalışır ama dakikada ~1 istek limitine tabidir.
 * API key (ücretsiz) ile limit 25.000 istek/gün'e çıkar.
 *
 * @param {string} url
 * @param {number} count
 * @param {Object} opts - { onProgress, onLog, shouldStop, apiKey }
 */
export async function runLighthouseViaPSI(url, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop, apiKey = '' } = typeof opts === 'function'
    ? { onProgress: opts }
    : opts;

  const results = [];
  // API key yoksa 65s bekle (rate-limit: ~1 req/dk), varsa 2s yeterli
  const DELAY_BETWEEN = apiKey ? 2000 : 65000;

  if (!apiKey) {
    onLog('⚠️  PSI API key yok — ölçümler arası 65s bekleniyor (rate-limit). Ücretsiz API key ile bu süre 2s\'e düşer.', 'warning');
  }

  for (let i = 0; i < count; i++) {
    if (shouldStop && shouldStop()) {
      onLog(`PSI test ${i + 1}/${count} — DURDURULDU`, 'warning');
      break;
    }

    let lastError, result;

    for (let retry = 0; retry < MAX_RETRIES; retry++) {
      if (shouldStop && shouldStop()) break;
      try {
        result = await fetchPSI(url, apiKey);
        break;
      } catch (err) {
        lastError = err;
        onLog(`⚠ PSI #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        await sleep(5000);
      }
    }

    if (result) {
      results.push({ measurementNo: i + 1, ...result });
      onLog(`✓ PSI #${i + 1}/${count} — FCP: ${result.fcp}ms | LCP: ${result.lcp}ms | TBT: ${result.tbt}ms | CLS: ${result.cls} | SI: ${result.speedIndex}`, 'success');
    } else {
      results.push({ measurementNo: i + 1, date: formatDate(), fcp: 'ERROR', lcp: 'ERROR', tbt: 'ERROR', cls: 'ERROR', speedIndex: 'ERROR' });
      onLog(`✗ PSI #${i + 1}/${count} HATA: ${lastError?.message}`, 'error');
    }

    if (onProgress) onProgress(i + 1, count);
    if (i < count - 1) await sleep(DELAY_BETWEEN);
  }

  return results;
}
