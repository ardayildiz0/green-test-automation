#!/usr/bin/env node

/**
 * GreenIT + Lighthouse Web Arayüzü - Backend
 * SSE (Server-Sent Events) ile gerçek zamanlı ilerleme
 */

import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import puppeteer from 'puppeteer';

import { runColdCacheTests, runWarmCacheTests } from './lib/greenitTest.js';
import { runLighthouseTests } from './lib/lighthouseTest.js';
import { createExcelReport } from './lib/excelWriter.js';
import { createDocxReport } from './lib/reportWriter.js';
import { normalizeUrl, sanitizeFilename, formatAuthorName } from './lib/utils.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;

// Config dosyasından ayarları oku
let appConfig = {};
try {
  const configPath = path.join(__dirname, 'config.json');
  if (fs.existsSync(configPath)) {
    appConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  }
} catch (e) { /* config yoksa varsayılanları kullan */ }

// ═══════════════════════════════════════
// GLOBAL STATE
// ═══════════════════════════════════════
let isRunning = false;
let shouldStop = false;
let browser = null;
let sseClients = [];

// Test sonuçları
let testState = {
  queue: [],         // { name, url, status: 'pending'|'running'|'done'|'error'|'skipped', results: null }
  currentIndex: -1,
  phase: '',         // 'cold', 'warm', 'lighthouse', 'excel', ''
  phaseProgress: 0,
  phaseTotal: 0,
  settings: {
    greenitCount: 20,
    lighthouseCount: 10,
    output: './results',
    authorName: '',
    advisorName: '',
    reportDate: ''
  },
  log: []
};

// ═══════════════════════════════════════
// SSE - GERÇEK ZAMANLI İLETİŞİM
// ═══════════════════════════════════════
function broadcast(event, data) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  sseClients = sseClients.filter(res => {
    try { res.write(msg); return true; }
    catch { return false; }
  });
}

function addLog(message, type = 'info') {
  const entry = { time: new Date().toLocaleTimeString('tr-TR'), message, type };
  testState.log.push(entry);
  if (testState.log.length > 200) testState.log.shift();
  broadcast('log', entry);
}

function broadcastState() {
  broadcast('state', {
    isRunning,
    queue: testState.queue,
    currentIndex: testState.currentIndex,
    phase: testState.phase,
    phaseProgress: testState.phaseProgress,
    phaseTotal: testState.phaseTotal,
    settings: testState.settings
  });
}

// ═══════════════════════════════════════
// TEST RUNNER
// ═══════════════════════════════════════
async function runAllTests() {
  if (isRunning) return;
  isRunning = true;
  shouldStop = false;

  const { greenitCount, lighthouseCount, output } = testState.settings;
  const outputDir = path.resolve(output);
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  addLog('Chromium başlatılıyor...', 'info');
  broadcastState();

  try {
    browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--remote-debugging-port=0']
    });

    const debugPort = new URL(browser.wsEndpoint()).port;
    addLog('Chromium hazır!', 'success');

    // Paylaşılan opsiyonlar — her test fonksiyonuna geçirilir
    const testOpts = (phaseName, phaseTotal) => ({
      shouldStop: () => shouldStop,
      onLog: (msg, type) => addLog(msg, type),
      onProgress: (current) => {
        testState.phaseProgress = current;
        broadcastState();
      }
    });

    for (let i = 0; i < testState.queue.length; i++) {
      if (shouldStop) {
        for (let j = i; j < testState.queue.length; j++) {
          if (testState.queue[j].status === 'pending') testState.queue[j].status = 'skipped';
        }
        addLog('İşlem kullanıcı tarafından durduruldu.', 'warning');
        break;
      }

      const item = testState.queue[i];
      if (item.status !== 'pending') continue;

      testState.currentIndex = i;
      item.status = 'running';
      addLog(`━━━ [${i + 1}/${testState.queue.length}] ${item.name} (${item.url}) ━━━`, 'info');
      broadcastState();

      try {
        // ─── Cold Cache → Warm Cache SIRALI ───
        // Sunucu rate-limiting'ini önlemek için sıralı çalıştırılıyor
        testState.phase = 'cold';
        testState.phaseProgress = 0;
        testState.phaseTotal = greenitCount;

        const coldOpts = {
          shouldStop: () => shouldStop,
          onLog: (msg, type) => addLog(msg, type),
          onProgress: (current) => {
            testState.phase = 'cold';
            testState.phaseProgress = current;
            broadcastState();
          }
        };

        const warmOpts = {
          shouldStop: () => shouldStop,
          onLog: (msg, type) => addLog(msg, type),
          onProgress: (current) => {
            testState.phase = 'warm';
            testState.phaseProgress = current;
            broadcastState();
          }
        };

        addLog(`❄️ Cold Cache başlıyor (${greenitCount} ölçüm)...`, 'info');
        broadcastState();
        const coldResults = await runColdCacheTests(browser, item.url, greenitCount, coldOpts);

        if (shouldStop) { item.status = 'skipped'; continue; }

        addLog(`🔥 Warm Cache başlıyor (${greenitCount} ölçüm)...`, 'info');
        testState.phase = 'warm';
        testState.phaseProgress = 0;
        broadcastState();
        const warmResults = await runWarmCacheTests(browser, item.url, greenitCount, warmOpts);

        if (shouldStop) { item.status = 'skipped'; continue; }

        // ─── Lighthouse ───
        testState.phase = 'lighthouse';
        testState.phaseProgress = 0;
        testState.phaseTotal = lighthouseCount;
        addLog(`💡 Lighthouse başlıyor (${lighthouseCount} ölçüm)...`, 'info');
        broadcastState();

        const lhResults = await runLighthouseTests(item.url, parseInt(debugPort), lighthouseCount, testOpts('lighthouse', lighthouseCount));

        if (shouldStop) { item.status = 'skipped'; continue; }

        // ─── Excel ───
        testState.phase = 'excel';
        addLog('📊 Excel raporu yazılıyor...', 'info');
        broadcastState();

        const safeFilename = sanitizeFilename(item.name);
        const excelPath = path.join(outputDir, `${safeFilename}_results.xlsx`);
        await createExcelReport(excelPath, item.name, item.url, coldResults, warmResults, lhResults);

        // ─── Screenshot ───
        testState.phase = 'screenshot';
        addLog('📸 Ekran görüntüsü alınıyor...', 'info');
        broadcastState();

        const screenshotPath = path.join(outputDir, `${safeFilename}_screenshot.png`);
        try {
          const ssPage = await browser.newPage();
          // Bot tespitini atlatmak için gerçekçi tarayıcı kimliği
          await ssPage.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36');
          await ssPage.setExtraHTTPHeaders({
            'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
          });
          // Headless algılamasını engelle
          await ssPage.evaluateOnNewDocument(() => {
            Object.defineProperty(navigator, 'webdriver', { get: () => false });
          });
          // Masaüstü görünüm, template oranına uygun (1.97:1)
          await ssPage.setViewport({ width: 1280, height: 649 });
          await ssPage.goto(item.url, { waitUntil: 'networkidle2', timeout: 30000 });
          await ssPage.screenshot({ path: screenshotPath, fullPage: false });
          await ssPage.close();
        } catch (ssErr) {
          addLog(`⚠️ Screenshot alınamadı: ${ssErr.message}`, 'warn');
        }

        // ─── DOCX Rapor ───
        testState.phase = 'docx';
        addLog('📝 DOCX rapor yazılıyor...', 'info');
        broadcastState();

        const docxPath = path.join(outputDir, `${safeFilename}_rapor.docx`);
        const authorForReport = testState.settings.authorName || appConfig.author || 'Arda Yıldız';
        const advisorForReport = testState.settings.advisorName || appConfig.advisor || '';
        const dateForReport = testState.settings.reportDate || appConfig.reportDate || '';
        await createDocxReport(docxPath, item.name, item.url, coldResults, warmResults, lhResults, screenshotPath, authorForReport, advisorForReport, dateForReport);

        item.status = 'done';
        item.results = {
          excelPath,
          docxPath,
          coldAvgEcoIndex: avg(coldResults, 'ecoIndex'),
          warmAvgEcoIndex: avg(warmResults, 'ecoIndex'),
          avgFCP: avg(lhResults, 'fcp'),
          avgLCP: avg(lhResults, 'lcp'),
          avgSpeedIndex: avg(lhResults, 'speedIndex')
        };

        addLog(`✅ ${item.name} tamamlandı → Excel: ${excelPath} | Rapor: ${docxPath}`, 'success');
      } catch (err) {
        item.status = 'error';
        item.error = err.message;
        addLog(`❌ ${item.name} HATA: ${err.message}`, 'error');
      }

      broadcastState();
    }
  } catch (err) {
    addLog(`Kritik hata: ${err.message}`, 'error');
  } finally {
    if (browser) {
      try { await browser.close(); } catch {}
      browser = null;
    }
    isRunning = false;
    testState.phase = '';
    testState.currentIndex = -1;
    addLog('Tüm işlemler tamamlandı.', 'success');
    broadcastState();
  }
}

function avg(arr, field) {
  const vals = arr.map(r => r[field]).filter(v => typeof v === 'number');
  if (!vals.length) return 'N/A';
  return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length * 100) / 100;
}

// ═══════════════════════════════════════
// HTTP SERVER
// ═══════════════════════════════════════
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);

  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  // ─── UI ───
  if (url.pathname === '/' || url.pathname === '/index.html') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(fs.readFileSync(path.join(__dirname, 'public', 'index.html'), 'utf-8'));
    return;
  }

  // ─── SSE Stream ───
  if (url.pathname === '/api/events') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive'
    });
    sseClients.push(res);
    // İlk bağlantıda mevcut state gönder
    res.write(`event: state\ndata: ${JSON.stringify({
      isRunning,
      queue: testState.queue,
      currentIndex: testState.currentIndex,
      phase: testState.phase,
      phaseProgress: testState.phaseProgress,
      phaseTotal: testState.phaseTotal,
      settings: testState.settings
    })}\n\n`);
    // Log geçmişi
    res.write(`event: logHistory\ndata: ${JSON.stringify(testState.log)}\n\n`);
    req.on('close', () => { sseClients = sseClients.filter(c => c !== res); });
    return;
  }

  // ─── API: JSON body parse ───
  let body = '';
  if (req.method === 'POST') {
    for await (const chunk of req) body += chunk;
  }

  // ─── Kuyruğa URL ekle ───
  if (url.pathname === '/api/queue' && req.method === 'POST') {
    const data = JSON.parse(body);
    // data.urls = [{ name, url }] veya data.name + data.url
    const urls = data.urls || [{ name: data.name, url: data.url }];
    for (const { name, url: rawUrl } of urls) {
      const normalUrl = normalizeUrl(rawUrl);
      testState.queue.push({ name, url: normalUrl, status: 'pending', results: null, error: null });
    }
    addLog(`${urls.length} URL kuyruğa eklendi.`, 'info');
    broadcastState();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, queueLength: testState.queue.length }));
    return;
  }

  // ─── Kuyruktan URL sil ───
  if (url.pathname.startsWith('/api/queue/') && req.method === 'DELETE') {
    const idx = parseInt(url.pathname.split('/').pop());
    if (testState.queue[idx] && testState.queue[idx].status === 'pending') {
      testState.queue.splice(idx, 1);
      broadcastState();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } else {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Silinemez (çalışıyor veya tamamlandı)' }));
    }
    return;
  }

  // ─── Kuyruğu temizle ───
  if (url.pathname === '/api/queue/clear' && req.method === 'POST') {
    if (!isRunning) {
      testState.queue = [];
      testState.log = [];
      broadcastState();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true }));
    } else {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Test çalışırken temizlenemez' }));
    }
    return;
  }

  // ─── Ayarları güncelle ───
  if (url.pathname === '/api/settings' && req.method === 'POST') {
    const data = JSON.parse(body);
    if (data.greenitCount) testState.settings.greenitCount = parseInt(data.greenitCount);
    if (data.lighthouseCount) testState.settings.lighthouseCount = parseInt(data.lighthouseCount);
    if (data.output) testState.settings.output = data.output;
    if (data.authorName !== undefined) testState.settings.authorName = data.authorName;
    if (data.advisorName !== undefined) testState.settings.advisorName = data.advisorName;
    if (data.reportDate !== undefined) testState.settings.reportDate = data.reportDate;
    broadcastState();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, settings: testState.settings }));
    return;
  }

  // ─── Testleri başlat ───
  if (url.pathname === '/api/start' && req.method === 'POST') {
    if (isRunning) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Zaten çalışıyor' }));
      return;
    }
    const pendingCount = testState.queue.filter(q => q.status === 'pending').length;
    if (pendingCount === 0) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Kuyrukta bekleyen URL yok' }));
      return;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true }));
    // Async olarak başlat
    runAllTests();
    return;
  }

  // ─── Testleri durdur ───
  if (url.pathname === '/api/stop' && req.method === 'POST') {
    shouldStop = true;
    addLog('Durdurma isteği gönderildi...', 'warning');
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true }));
    return;
  }

  // 404
  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Not found' }));
});

server.listen(PORT, () => {
  console.log(`\n╔══════════════════════════════════════════════════════════════╗`);
  console.log(`║   GreenIT + Lighthouse Web Arayüzü                         ║`);
  console.log(`║   http://localhost:${PORT}                                     ║`);
  console.log(`╚══════════════════════════════════════════════════════════════╝\n`);
});
