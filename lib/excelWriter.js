/**
 * Excel Oluşturma Modülü
 * Template Excel dosyasından kurum bazında sonuç dosyaları oluşturur.
 *
 * Sheet yapısı:
 *   1. Cold Cache GreenIT: Measurement No, Date, Request Count, Page Size (KB), DOM Size, CO2 (gCO2e), Water (cl), EcoIndex, Grade
 *   2. Warm Cache GreenIT: Aynı yapı
 *   3. Lighthouse Metrics: Measurement No, Date, FCP (ms), LCP (ms), TBT (ms), CLS, Speed Index
 *   4. Summary: Metric, Cold Cache Average, Warm Cache Average
 */

import ExcelJS from 'exceljs';
import path from 'path';
import { getEcoIndexGrade, GRADE_RANGES } from './utils.js';

/**
 * Header stilleri
 */
const HEADER_STYLE = {
  font: { bold: true, size: 11, color: { argb: 'FFFFFFFF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } }, // Yeşil
  alignment: { horizontal: 'center', vertical: 'middle' },
  border: {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  }
};

const DATA_STYLE = {
  alignment: { horizontal: 'center', vertical: 'middle' },
  border: {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  }
};

const SUMMARY_HEADER_STYLE = {
  font: { bold: true, size: 11, color: { argb: 'FFFFFFFF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1565C0' } }, // Mavi
  alignment: { horizontal: 'center', vertical: 'middle' },
  border: {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  }
};

/**
 * Grade'e göre hücre rengi
 */
function getGradeColor(grade) {
  const colors = {
    'A': 'FF4CAF50', // Yeşil
    'B': 'FF8BC34A', // Açık yeşil
    'C': 'FFCDDC39', // Sarı-yeşil
    'D': 'FFFFEB3B', // Sarı
    'E': 'FFFF9800', // Turuncu
    'F': 'FFFF5722', // Kırmızı-turuncu
    'G': 'FFF44336'  // Kırmızı
  };
  return colors[grade] || 'FFFFFFFF';
}

/**
 * Bir sheet'e header uygula
 */
function applyHeaders(sheet, headers, style) {
  const headerRow = sheet.getRow(1);
  headers.forEach((header, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = header;
    cell.style = style;
  });
  headerRow.height = 25;
}

/**
 * Veri satırlarını yaz
 */
function writeDataRows(sheet, data, columns) {
  data.forEach((item, rowIdx) => {
    const row = sheet.getRow(rowIdx + 2);
    columns.forEach((col, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = item[col];
      cell.style = { ...DATA_STYLE };

      // Grade sütununa renk uygula
      if (col === 'grade' && typeof item[col] === 'string' && item[col] !== 'ERROR') {
        cell.style.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: getGradeColor(item[col]) }
        };
        cell.style.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      }
    });
    row.height = 20;
  });
}

/**
 * Ortalama hesapla (ERROR değerlerini atla)
 */
function calculateAverage(data, field) {
  const validValues = data
    .map(item => item[field])
    .filter(v => typeof v === 'number' && !isNaN(v));

  if (validValues.length === 0) return 'N/A';
  const avg = validValues.reduce((sum, v) => sum + v, 0) / validValues.length;
  return Math.round(avg * 100) / 100;
}

/**
 * Kurum için Excel dosyası oluştur
 * @param {string} outputPath - Çıktı dosya yolu
 * @param {string} institutionName - Kurum adı
 * @param {string} url - Test edilen URL
 * @param {Array} coldCacheResults - Cold cache GreenIT sonuçları
 * @param {Array} warmCacheResults - Warm cache GreenIT sonuçları
 * @param {Array} lighthouseResults - Lighthouse sonuçları
 */
export async function createExcelReport(outputPath, institutionName, url, coldCacheResults, warmCacheResults, lighthouseResults) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'GreenIT Automation Script';
  workbook.created = new Date();

  // ===== 1. COLD CACHE GreenIT Sheet =====
  const coldSheet = workbook.addWorksheet('Cold Cache GreenIT');
  const greenitHeaders = [
    'Measurement No', 'Measurement Date', 'Request Count',
    'Page Size (KB)', 'DOM Size', 'CO2 Emission (gCO2e)',
    'Water Consumption (cl)', 'EcoIndex', 'Grade'
  ];
  const greenitColumns = [
    'measurementNo', 'date', 'requestCount',
    'pageSizeKB', 'domSize', 'co2',
    'water', 'ecoIndex', 'grade'
  ];

  applyHeaders(coldSheet, greenitHeaders, HEADER_STYLE);
  writeDataRows(coldSheet, coldCacheResults, greenitColumns);

  // Sütun genişlikleri
  coldSheet.columns = [
    { width: 16 }, { width: 20 }, { width: 16 },
    { width: 16 }, { width: 12 }, { width: 20 },
    { width: 22 }, { width: 12 }, { width: 10 }
  ];

  // ===== 2. WARM CACHE GreenIT Sheet =====
  const warmSheet = workbook.addWorksheet('Warm Cache GreenIT');
  applyHeaders(warmSheet, greenitHeaders, HEADER_STYLE);
  writeDataRows(warmSheet, warmCacheResults, greenitColumns);
  warmSheet.columns = coldSheet.columns.map(c => ({ width: c.width }));

  // ===== 3. LIGHTHOUSE METRICS Sheet =====
  const lhSheet = workbook.addWorksheet('Lighthouse Metrics');
  const lhHeaders = [
    'Measurement No', 'Measurement Date', 'FCP (ms)',
    'LCP (ms)', 'TBT (ms)', 'CLS', 'Speed Index'
  ];
  const lhColumns = [
    'measurementNo', 'date', 'fcp',
    'lcp', 'tbt', 'cls', 'speedIndex'
  ];

  applyHeaders(lhSheet, lhHeaders, HEADER_STYLE);
  writeDataRows(lhSheet, lighthouseResults, lhColumns);

  lhSheet.columns = [
    { width: 16 }, { width: 20 }, { width: 14 },
    { width: 14 }, { width: 14 }, { width: 10 }, { width: 14 }
  ];

  // ===== 4. SUMMARY Sheet =====
  const summarySheet = workbook.addWorksheet('Summary');

  // Sütun genişlikleri (önce ayarla)
  summarySheet.columns = [
    { width: 25 }, { width: 22 }, { width: 22 }
  ];

  // ─── Kurum bilgileri (satır 1-3) ───
  const infoStyle = { font: { bold: true, size: 11 } };
  summarySheet.getRow(1).getCell(1).value = 'Institution';
  summarySheet.getRow(1).getCell(1).style = infoStyle;
  summarySheet.getRow(1).getCell(2).value = institutionName;

  summarySheet.getRow(2).getCell(1).value = 'URL';
  summarySheet.getRow(2).getCell(1).style = infoStyle;
  summarySheet.getRow(2).getCell(2).value = url;

  summarySheet.getRow(3).getCell(1).value = 'Test Date';
  summarySheet.getRow(3).getCell(1).style = infoStyle;
  summarySheet.getRow(3).getCell(2).value = new Date().toISOString().split('T')[0];

  // ─── EcoIndex Ortalama Grade (satır 5-6) ───
  const coldAvgEco = calculateAverage(coldCacheResults, 'ecoIndex');
  const warmAvgEco = calculateAverage(warmCacheResults, 'ecoIndex');
  const coldGrade = typeof coldAvgEco === 'number' ? getEcoIndexGrade(coldAvgEco) : 'N/A';
  const warmGrade = typeof warmAvgEco === 'number' ? getEcoIndexGrade(warmAvgEco) : 'N/A';
  const coldGradeInfo = GRADE_RANGES.find(g => g.grade === coldGrade);
  const warmGradeInfo = GRADE_RANGES.find(g => g.grade === warmGrade);

  const gradeHeaderStyle = {
    font: { bold: true, size: 11, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B5E20' } },
    alignment: { horizontal: 'center', vertical: 'middle' },
    border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
  };

  const gradeRow = summarySheet.getRow(5);
  gradeRow.getCell(1).value = 'Average EcoIndex Grade';
  gradeRow.getCell(1).style = gradeHeaderStyle;
  gradeRow.getCell(2).value = 'Cold Cache';
  gradeRow.getCell(2).style = gradeHeaderStyle;
  gradeRow.getCell(3).value = 'Warm Cache';
  gradeRow.getCell(3).style = gradeHeaderStyle;
  gradeRow.height = 25;

  const gradeDataRow = summarySheet.getRow(6);
  // Cold Cache Grade
  gradeDataRow.getCell(1).value = '';
  gradeDataRow.getCell(2).value = `${coldGrade} (${coldAvgEco})`;
  gradeDataRow.getCell(2).style = {
    ...DATA_STYLE,
    font: { bold: true, size: 16, color: { argb: 'FFFFFFFF' } },
    fill: coldGradeInfo
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + coldGradeInfo.color } }
      : {}
  };
  // Warm Cache Grade
  gradeDataRow.getCell(3).value = `${warmGrade} (${warmAvgEco})`;
  gradeDataRow.getCell(3).style = {
    ...DATA_STYLE,
    font: { bold: true, size: 16, color: { argb: 'FFFFFFFF' } },
    fill: warmGradeInfo
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + warmGradeInfo.color } }
      : {}
  };
  gradeDataRow.height = 35;

  // ─── Not Aralıkları Tablosu (satır 8-15) ───
  const rangeHeaderRow = summarySheet.getRow(8);
  ['Grade', 'Score Range', ''].forEach((h, i) => {
    const cell = rangeHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.style = {
      font: { bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF424242' } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
  });

  GRADE_RANGES.forEach((g, idx) => {
    const row = summarySheet.getRow(9 + idx);
    row.getCell(1).value = g.grade;
    row.getCell(1).style = {
      ...DATA_STYLE,
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + g.color } }
    };
    row.getCell(2).value = `${g.min} – ${g.max}`;
    row.getCell(2).style = DATA_STYLE;
  });

  // ─── GreenIT Ortalamaları (satır 17) ───
  const gHeaderRow = summarySheet.getRow(17);
  ['GreenIT Metric', 'Cold Cache Average', 'Warm Cache Average'].forEach((h, i) => {
    const cell = gHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.style = SUMMARY_HEADER_STYLE;
  });
  gHeaderRow.height = 25;

  const greenitMetrics = [
    { label: 'EcoIndex', field: 'ecoIndex' },
    { label: 'Grade', field: null },
    { label: 'CO2 Emission (gCO2e)', field: 'co2' },
    { label: 'Water Consumption (cl)', field: 'water' },
    { label: 'Request Count', field: 'requestCount' },
    { label: 'Page Size (KB)', field: 'pageSizeKB' },
    { label: 'DOM Size', field: 'domSize' }
  ];

  greenitMetrics.forEach((metric, idx) => {
    const row = summarySheet.getRow(18 + idx);
    row.getCell(1).value = metric.label;
    row.getCell(1).style = { ...DATA_STYLE, font: { bold: true } };

    if (metric.field === null) {
      // Grade satırı — ortalamadan hesapla
      row.getCell(2).value = coldGrade;
      row.getCell(2).style = {
        ...DATA_STYLE,
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: coldGradeInfo
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + coldGradeInfo.color } }
          : {}
      };
      row.getCell(3).value = warmGrade;
      row.getCell(3).style = {
        ...DATA_STYLE,
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: warmGradeInfo
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + warmGradeInfo.color } }
          : {}
      };
    } else {
      row.getCell(2).value = calculateAverage(coldCacheResults, metric.field);
      row.getCell(2).style = DATA_STYLE;
      row.getCell(3).value = calculateAverage(warmCacheResults, metric.field);
      row.getCell(3).style = DATA_STYLE;
    }
  });

  // ─── Lighthouse Ortalamaları (satır 26) ───
  const lhHeaderStyle = {
    ...SUMMARY_HEADER_STYLE,
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF6F00' } }
  };

  const lHeaderRow = summarySheet.getRow(26);
  ['Lighthouse Metric', 'Average Value', ''].forEach((h, i) => {
    const cell = lHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.style = lhHeaderStyle;
  });
  lHeaderRow.height = 25;

  const lhMetrics = [
    { label: 'FCP (ms)', field: 'fcp' },
    { label: 'LCP (ms)', field: 'lcp' },
    { label: 'TBT (ms)', field: 'tbt' },
    { label: 'CLS', field: 'cls' },
    { label: 'Speed Index', field: 'speedIndex' }
  ];

  lhMetrics.forEach((metric, idx) => {
    const row = summarySheet.getRow(27 + idx);
    row.getCell(1).value = metric.label;
    row.getCell(1).style = { ...DATA_STYLE, font: { bold: true } };
    row.getCell(2).value = calculateAverage(lighthouseResults, metric.field);
    row.getCell(2).style = DATA_STYLE;
  });

  // Dosyayı kaydet
  await workbook.xlsx.writeFile(outputPath);
}
