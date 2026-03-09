#!/usr/bin/env node

/**
 * GreenIT + Lighthouse Otomasyon Scripti
 * SENG434 - Green Software Engineering
 *
 * Kullanım:
 *   node index.js --url <URL> --name <KURUM_ADI> [--count 20] [--output ./results]
 *   node index.js --file urls.txt [--count 20] [--output ./results]
 */

import puppeteer from 'puppeteer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

import { runColdCacheTests, runWarmCacheTests } from './lib/greenitTest.js';
import { runLighthouseTests } from './lib/lighthouseTest.js';
import { createExcelReport } from './lib/excelWriter.js';
import { parseArgs, printHelp, showProgress, normalizeUrl, sanitizeFilename, sleep } from './lib/utils.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ═══════════════════════════════════════════════
// ANA FONKSİYON
// ═══════════════════════════════════════════════

async function main() {
  const args = parseArgs(process.argv);

  // Argüman kontrolü
  if (!args.url && !args.file) {
    printHelp();
    console.error('\n❌ Lütfen --url veya --file parametresi belirtin.\n');
    process.exit(1);
  }

  if (args.url && !args.name) {
    console.error('\n❌ --url kullanırken --name parametresi zorunludur.\n');
    process.exit(1);
  }

  // URL listesini oluştur
  let urlList = [];

  if (args.file) {
    // Dosyadan oku
    if (!fs.existsSync(args.file)) {
      console.error(`\n❌ Dosya bulunamadı: ${args.file}\n`);
      process.exit(1);
    }

    const content = fs.readFileSync(args.file, 'utf-8');
    const lines = content.split('\n').filter(l => l.trim() && !l.trim().startsWith('#'));

    for (const line of lines) {
      const parts = line.split('|').map(p => p.trim());
      if (parts.length >= 2) {
        urlList.push({ name: parts[0], url: normalizeUrl(parts[1]) });
      } else if (parts.length === 1) {
        // URL'den isim türet
        const url = normalizeUrl(parts[0]);
        const hostname = new URL(url).hostname.replace('www.', '');
        urlList.push({ name: hostname, url });
      }
    }
  } else {
    urlList.push({ name: args.name, url: normalizeUrl(args.url) });
  }

  if (urlList.length === 0) {
    console.error('\n❌ İşlenecek URL bulunamadı.\n');
    process.exit(1);
  }

  // Çıktı klasörünü oluştur
  const outputDir = path.resolve(args.output);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Başlangıç bilgisi
  console.log('\n╔══════════════════════════════════════════════════════════════╗');
  console.log('║        GreenIT + Lighthouse Otomasyon Scripti               ║');
  console.log('║        SENG434 - Green Software Engineering                 ║');
  console.log('╚══════════════════════════════════════════════════════════════╝\n');
  console.log(`📋 Toplam URL sayısı: ${urlList.length}`);
  console.log(`📊 GreenIT ölçüm sayısı: ${args.greenitCount}`);
  console.log(`📊 Lighthouse ölçüm sayısı: ${args.lighthouseCount}`);
  console.log(`📁 Çıktı klasörü: ${outputDir}\n`);

  // ═══════════════════════════════════════════════
  // BROWSER BAŞLAT
  // ═══════════════════════════════════════════════

  console.log('🚀 Chromium başlatılıyor...\n');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--remote-debugging-port=0' // Rastgele port
    ]
  });

  // Lighthouse için Chrome debug port'unu al
  const browserWSEndpoint = browser.wsEndpoint();
  const debugPort = new URL(browserWSEndpoint).port;

  const totalUrls = urlList.length;
  const results = [];
  const startTime = Date.now();

  try {
    for (let urlIdx = 0; urlIdx < urlList.length; urlIdx++) {
      const { name, url } = urlList[urlIdx];
      const urlStartTime = Date.now();

      console.log('═'.repeat(60));
      console.log(`\n🏢 [${urlIdx + 1}/${totalUrls}] ${name}`);
      console.log(`🔗 ${url}\n`);

      // ─── AŞAMA 1: Cold Cache GreenIT ───
      console.log('❄️  Cold Cache GreenIT Testleri:');
      const coldResults = await runColdCacheTests(browser, url, args.greenitCount, (current, total) => {
        showProgress(current, total, `Cold Cache #${current}`);
      });
      console.log('');

      // ─── AŞAMA 2: Warm Cache GreenIT ───
      console.log('🔥 Warm Cache GreenIT Testleri:');
      const warmResults = await runWarmCacheTests(browser, url, args.greenitCount, (current, total) => {
        showProgress(current, total, `Warm Cache #${current}`);
      });
      console.log('');

      // ─── AŞAMA 3: Lighthouse ───
      console.log('💡 Lighthouse Testleri:');
      const lhResults = await runLighthouseTests(url, parseInt(debugPort), args.lighthouseCount, (current, total) => {
        showProgress(current, total, `Lighthouse #${current}`);
      });
      console.log('');

      // ─── AŞAMA 4: Excel Oluştur ───
      const safeFilename = sanitizeFilename(name);
      const excelPath = path.join(outputDir, `${safeFilename}_results.xlsx`);

      console.log('📊 Excel raporu oluşturuluyor...');
      await createExcelReport(excelPath, name, url, coldResults, warmResults, lhResults);
      console.log(`✅ Kaydedildi: ${excelPath}`);

      const urlDuration = Math.round((Date.now() - urlStartTime) / 1000);
      console.log(`⏱  Süre: ${Math.floor(urlDuration / 60)}dk ${urlDuration % 60}sn\n`);

      results.push({
        name,
        url,
        excelPath,
        success: true,
        duration: urlDuration
      });

      // URL'ler arası bekleme
      if (urlIdx < urlList.length - 1) {
        console.log('⏳ Sonraki URL için 5 saniye bekleniyor...\n');
        await sleep(5000);
      }
    }
  } catch (err) {
    console.error(`\n💥 Kritik hata: ${err.message}`);
    console.error(err.stack);
  } finally {
    await browser.close();
  }

  // ═══════════════════════════════════════════════
  // SONUÇ ÖZETİ
  // ═══════════════════════════════════════════════

  const totalDuration = Math.round((Date.now() - startTime) / 1000);
  const successCount = results.filter(r => r.success).length;

  console.log('\n' + '═'.repeat(60));
  console.log('\n📊 SONUÇ ÖZETİ');
  console.log('─'.repeat(40));
  console.log(`  Toplam URL: ${totalUrls}`);
  console.log(`  Başarılı:   ${successCount}`);
  console.log(`  Başarısız:  ${totalUrls - successCount}`);
  console.log(`  Toplam süre: ${Math.floor(totalDuration / 60)}dk ${totalDuration % 60}sn`);
  console.log('─'.repeat(40));

  results.forEach(r => {
    const status = r.success ? '✅' : '❌';
    console.log(`  ${status} ${r.name} → ${r.excelPath}`);
  });

  console.log('\n✨ İşlem tamamlandı!\n');
}

// ═══════════════════════════════════════════════
// ÇALIŞTIR
// ═══════════════════════════════════════════════

main().catch(err => {
  console.error('\n💥 Beklenmeyen hata:', err);
  process.exit(1);
});
