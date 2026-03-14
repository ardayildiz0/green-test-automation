/**
 * Lighthouse Test Modülü
 * Lighthouse programmatic API ile performans metriklerini toplar.
 * FCP, LCP, TBT, CLS, Speed Index
 */

import { formatDate, sleep } from './utils.js';

const MAX_RETRIES = 3;

const defaultLog = (msg, type = 'info') => {
  if (type === 'error') console.error(msg);
  else console.log(msg);
};

/**
 * Tek bir Lighthouse ölçümü yap
 */
async function singleLighthouseRun(url, port) {
  const lighthouse = (await import('lighthouse')).default;

  const flags = {
    port,
    output: 'json',
    onlyCategories: ['performance'],
    formFactor: 'desktop',
    screenEmulation: {
      mobile: false, width: 1920, height: 1080,
      deviceScaleFactor: 1, disabled: false
    },
    throttling: {
      rttMs: 40, throughputKbps: 10240, cpuSlowdownMultiplier: 1,
      requestLatencyMs: 0, downloadThroughputKbps: 0, uploadThroughputKbps: 0
    },
    disableStorageReset: false
  };

  const config = {
    extends: 'lighthouse:default',
    settings: {
      formFactor: 'desktop',
      throttling: flags.throttling,
      screenEmulation: flags.screenEmulation
    }
  };

  const result = await lighthouse(url, flags, config);
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
 * Lighthouse testleri
 * @param {Object} opts - { onProgress, onLog, shouldStop }
 */
export async function runLighthouseTests(url, port, count, opts = {}) {
  const { onProgress, onLog = defaultLog, shouldStop } = typeof opts === 'function'
    ? { onProgress: opts }
    : opts;

  const results = [];

  for (let i = 0; i < count; i++) {
    // ── Durdurma kontrolü ──
    if (shouldStop && shouldStop()) {
      onLog(`Lighthouse test ${i + 1}/${count} — DURDURULDU`, 'warning');
      break;
    }

    let lastError, result;

    for (let retry = 0; retry < MAX_RETRIES; retry++) {
      if (shouldStop && shouldStop()) break;
      try {
        result = await singleLighthouseRun(url, port);
        break;
      } catch (err) {
        lastError = err;
        onLog(`⚠ Lighthouse #${i + 1} retry ${retry + 1}: ${err.message}`, 'warning');
        await sleep(1000);
      }
    }

    if (result) {
      results.push({ measurementNo: i + 1, ...result });
      onLog(`✓ Lighthouse #${i + 1}/${count} — FCP: ${result.fcp}ms | LCP: ${result.lcp}ms | TBT: ${result.tbt}ms | CLS: ${result.cls} | SI: ${result.speedIndex}`, 'success');
    } else {
      results.push({
        measurementNo: i + 1, date: formatDate(),
        fcp: 'ERROR', lcp: 'ERROR', tbt: 'ERROR', cls: 'ERROR', speedIndex: 'ERROR'
      });
      onLog(`✗ Lighthouse #${i + 1}/${count} HATA: ${lastError?.message}`, 'error');
    }

    if (onProgress) onProgress(i + 1, count);
    if (i < count - 1) await sleep(500);
  }

  return results;
}
