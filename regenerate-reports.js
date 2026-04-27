#!/usr/bin/env node

/**
 * Mevcut _results.xlsx dosyalarından _rapor.docx'leri yeniden oluşturur.
 *
 * Yeni ölçüm almadan sadece mevcut verilerden Word raporu çıkarır.
 * reportWriter.js güncellemelerini tüm raporlara uygulamak için kullanılır.
 */

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

import { createDocxReport } from './lib/reportWriter.js';
import { formatAuthorName } from './lib/utils.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const RESULTS_DIR = path.join(__dirname, 'results');

// Config'ten yazar adını ve e-postasını al
let authorName = 'Arda Yıldız';
let authorEmail = 'arda730a@gmail.com';
try {
  const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf-8'));
  if (config.author) authorName = config.author;
  if (config.authorEmail) authorEmail = config.authorEmail;
} catch (e) {
  // use default
}

// Institution name'i xlsx dosya adından çıkar: "Bafra_Belediyesi_(Samsun)_results.xlsx" → "Bafra Belediyesi (Samsun)"
function filenameToInstitution(filename) {
  return filename
    .replace(/_results\.xlsx$/, '')
    .replace(/_/g, ' ');
}

// Bir sheet'i okuyup satırları object array'e dönüştür
function readSheet(sheet, columnMap) {
  const rows = [];
  const rowCount = sheet.rowCount;
  for (let i = 2; i <= rowCount; i++) {
    const row = sheet.getRow(i);
    // Boş satırları atla
    if (!row.hasValues) continue;
    const obj = {};
    let hasData = false;
    columnMap.forEach((key, idx) => {
      const val = row.getCell(idx + 1).value;
      obj[key] = val;
      if (val !== null && val !== undefined && val !== '') hasData = true;
    });
    if (hasData) rows.push(obj);
  }
  return rows;
}

async function regenerateOne(xlsxPath) {
  const baseName = path.basename(xlsxPath, '.xlsx').replace(/_results$/, '');
  const institutionName = filenameToInstitution(path.basename(xlsxPath));
  const docxPath = path.join(RESULTS_DIR, `${baseName}_rapor.docx`);
  const screenshotPath = path.join(RESULTS_DIR, `${baseName}_screenshot.png`);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);

  // GreenIT kolonları: measurementNo, date, requestCount, pageSizeKB, domSize, co2, water, ecoIndex, grade
  const greenitCols = ['measurementNo', 'date', 'requestCount', 'pageSizeKB', 'domSize', 'co2', 'water', 'ecoIndex', 'grade'];
  const lhCols = ['measurementNo', 'date', 'fcp', 'lcp', 'tbt', 'cls', 'speedIndex'];

  const coldSheet = wb.getWorksheet('Cold Cache GreenIT');
  const warmSheet = wb.getWorksheet('Warm Cache GreenIT');
  const lhSheet = wb.getWorksheet('Lighthouse Metrics');
  const summarySheet = wb.getWorksheet('Summary');

  if (!coldSheet || !warmSheet || !lhSheet || !summarySheet) {
    throw new Error(`Beklenen sheet'ler bulunamadı: ${xlsxPath}`);
  }

  const coldResults = readSheet(coldSheet, greenitCols);
  const warmResults = readSheet(warmSheet, greenitCols);
  const lhResults = readSheet(lhSheet, lhCols);

  // URL'yi Summary sheet'in B2 hücresinden al
  let url = summarySheet.getRow(2).getCell(2).value;
  if (url && typeof url === 'object' && url.text) url = url.text;
  if (!url) url = '';
  url = String(url).trim();

  const screenshot = fs.existsSync(screenshotPath) ? screenshotPath : undefined;

  await createDocxReport(
    docxPath,
    institutionName,
    url,
    coldResults,
    warmResults,
    lhResults,
    screenshot,
    formatAuthorName(authorName),
    undefined, // advisorName
    undefined, // reportDate
    authorEmail
  );

  return { institutionName, docxPath, hasScreenshot: !!screenshot, coldN: coldResults.length, warmN: warmResults.length, lhN: lhResults.length };
}

async function main() {
  const files = fs.readdirSync(RESULTS_DIR)
    .filter(f => f.endsWith('_results.xlsx'))
    .sort();

  console.log(`\n📊 Toplam ${files.length} xlsx dosyası bulundu\n`);
  console.log('═'.repeat(70));

  const results = { success: [], failed: [] };

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const xlsxPath = path.join(RESULTS_DIR, f);
    process.stdout.write(`[${i + 1}/${files.length}] ${f} ... `);
    try {
      const r = await regenerateOne(xlsxPath);
      const ssFlag = r.hasScreenshot ? '✓ss' : '✗ss';
      console.log(`✅ (cold=${r.coldN}, warm=${r.warmN}, lh=${r.lhN}, ${ssFlag})`);
      results.success.push(r.institutionName);
    } catch (err) {
      console.log(`❌ ${err.message}`);
      results.failed.push({ file: f, error: err.message });
    }
  }

  console.log('═'.repeat(70));
  console.log(`\n✅ Başarılı: ${results.success.length}`);
  console.log(`❌ Başarısız: ${results.failed.length}`);
  if (results.failed.length > 0) {
    console.log('\nHatalar:');
    results.failed.forEach(f => console.log(`  - ${f.file}: ${f.error}`));
  }
  console.log('');
}

main().catch(err => {
  console.error('\n💥 Beklenmeyen hata:', err);
  process.exit(1);
});
