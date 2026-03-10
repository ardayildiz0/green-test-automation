/**
 * DOCX Rapor Oluşturma Modülü — Template Tabanlı
 *
 * Orijinal Writing Guideline şablonunu (.docx) temel alarak:
 *   1. Kurum adı ve URL'yi değiştirir
 *   2. Ölçüm tablolarını gerçek verilerle doldurur
 *   3. Analiz bölümlerini (6.1-6.5) sonuçlara göre otomatik yazar
 *   4. İyileştirme önerilerini (7.1-7.7) sonuçlara göre otomatik yazar
 *   5. ÖZET ve SONUÇ bölümlerini günceller
 *
 * Yaklaşım: JSZip ile .docx (zip) aç → document.xml üzerinde XML düzenleme → kaydet
 * Logolar, görseller, stiller, header/footer tamamen korunur.
 */

import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { getEcoIndexGrade, GRADE_RANGES } from './utils.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE_PATH = path.join(__dirname, '..', 'template', 'report_template.docx');

// ═══════════════════════════════════════
// YARDIMCI FONKSİYONLAR
// ═══════════════════════════════════════

function avg(arr, field) {
  const vals = arr.map(r => r[field]).filter(v => typeof v === 'number' && !isNaN(v));
  if (!vals.length) return null;
  return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length * 100) / 100;
}

function fmtNum(val, decimals = 2) {
  if (val === null || val === undefined || val === 'N/A' || val === 'ERROR') return '-';
  return typeof val === 'number' ? val.toFixed(decimals) : String(val);
}

function escXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ═══════════════════════════════════════
// XML METİN BİRLEŞTİRME + DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * Bir paragraf (w:p) bloğunun tüm w:t etiketlerinden birleşik metni çıkarır.
 */
function extractParagraphText(pXml) {
  const textParts = [];
  const re = /<w:t(?:\s[^>]*)?>(.*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(pXml)) !== null) {
    textParts.push(m[1]);
  }
  return textParts.join('');
}

/**
 * Bir paragraf (w:p) bloğundaki tüm run'ları (w:r) kaldırıp,
 * verilen metni tek bir w:r olarak yeniden oluşturur.
 * Run properties (w:rPr) ilk run'dan alınır.
 */
function replaceParagraphText(pXml, newText) {
  // İlk run'dan rPr'yi çek
  const rPrMatch = pXml.match(/<w:r>[\s\S]*?<w:rPr>([\s\S]*?)<\/w:rPr>/);
  const rPr = rPrMatch ? `<w:rPr>${rPrMatch[1]}</w:rPr>` : '';

  // Paragraf açılış ve kapanış kısımlarını bul
  // pPr'yi koru, tüm run'ları kaldır
  const pPrMatch = pXml.match(/(<w:p[^>]*>[\s\S]*?<\/w:pPr>)/);
  const pOpening = pPrMatch ? pPrMatch[1] : pXml.match(/(<w:p[^>]*>)/)[1];

  // Yeni paragraf oluştur (hyperlink vb. de temizle)
  const needsPreserve = newText.startsWith(' ') || newText.endsWith(' ') ? ' xml:space="preserve"' : '';
  return `${pOpening}\n      <w:r>${rPr}<w:t${needsPreserve}>${escXml(newText)}</w:t></w:r>\n    </w:p>`;
}

// ═══════════════════════════════════════
// TABLO SATIRI OLUŞTURMA
// ═══════════════════════════════════════

function greenitDataRow(r, colWidths) {
  const vals = [
    r.measurementNo,
    r.date || '-',
    fmtNum(r.requestCount, 0),
    fmtNum(r.pageSizeKB),
    fmtNum(r.domSize, 0),
    fmtNum(r.co2),
    fmtNum(r.water),
    fmtNum(r.ecoIndex),
    r.grade || '-'
  ];
  return makeTableRow(vals, colWidths);
}

function lighthouseDataRow(r, colWidths) {
  const vals = [
    r.measurementNo,
    r.date || '-',
    fmtNum(r.fcp, 0),
    fmtNum(r.lcp, 0),
    fmtNum(r.tbt, 0),
    fmtNum(r.cls, 3),
    fmtNum(r.speedIndex, 0)
  ];
  return makeTableRow(vals, colWidths);
}

function makeTableRow(vals, colWidths) {
  const cells = vals.map((v, i) => `
        <w:tc>
          <w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="${randomParaId()}" w14:textId="77777777" w:rsidR="00DA6648" w:rsidRDefault="00DA6648" w:rsidP="00D12359">
            <w:pPr>
              <w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>
              <w:jc w:val="center"/>
              <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="Times New Roman" w:hAnsiTheme="majorHAnsi" w:cs="Times New Roman"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsia="Times New Roman" w:hAnsiTheme="majorHAnsi" w:cs="Times New Roman"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>${escXml(String(v))}</w:t></w:r>
          </w:p>
        </w:tc>`).join('');

  return `\n      <w:tr w:rsidR="00DA6648" w14:paraId="${randomParaId()}" w14:textId="77777777"><w:trPr><w:jc w:val="center"/></w:trPr>${cells}\n      </w:tr>`;
}

function randomParaId() {
  // paraId must be < 0x7FFFFFFF
  return Math.floor(Math.random() * 0x7FFFFFFF).toString(16).toUpperCase().padStart(8, '0');
}

// ═══════════════════════════════════════
// TABLO DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * Template'deki bir tabloyu bul ve veri satırlarını değiştir.
 * Header satırı (ilk w:tr) korunur, geri kalanlar yeni verilerle değiştirilir.
 */
function replaceTableData(xml, tableStartLine, tableEndLine, dataRows) {
  const lines = xml.split('\n');

  // Tablo bloğunu bul
  let tableContent = '';
  let tblStart = -1, tblEnd = -1;

  // Satır numarası yerine tablo bloğunu regex ile bul
  // tableStartLine ve tableEndLine kullanarak aralığı belirle
  const tableBlock = lines.slice(tableStartLine - 1, tableEndLine).join('\n');

  // Header row'u koru (ilk w:tr)
  const firstTrEnd = tableBlock.indexOf('</w:tr>') + '</w:tr>'.length;
  const headerPart = tableBlock.substring(0, firstTrEnd);

  // Tablo kapanışı
  const tableClose = '\n    </w:tbl>';

  // Yeni tablo oluştur: header + data rows + close
  const newTable = headerPart + '\n' + dataRows.join('\n') + tableClose;

  // Eski tabloyu yenisiyle değiştir
  const before = lines.slice(0, tableStartLine - 1).join('\n');
  const after = lines.slice(tableEndLine).join('\n');
  return before + '\n' + newTable + '\n' + after;
}

// ═══════════════════════════════════════
// ANALİZ METNİ ÜRETME
// ═══════════════════════════════════════

function generateAnalysisTexts(institutionName, url, coldResults, warmResults, lhResults) {
  const coldAvgEco = avg(coldResults, 'ecoIndex');
  const warmAvgEco = avg(warmResults, 'ecoIndex');
  const coldGrade = coldAvgEco !== null ? getEcoIndexGrade(coldAvgEco) : 'N/A';
  const warmGrade = warmAvgEco !== null ? getEcoIndexGrade(warmAvgEco) : 'N/A';

  const coldAvgCO2 = avg(coldResults, 'co2');
  const warmAvgCO2 = avg(warmResults, 'co2');
  const coldAvgWater = avg(coldResults, 'water');
  const warmAvgWater = avg(warmResults, 'water');
  const coldAvgDom = avg(coldResults, 'domSize');
  const warmAvgDom = avg(warmResults, 'domSize');
  const coldAvgReq = avg(coldResults, 'requestCount');
  const warmAvgReq = avg(warmResults, 'requestCount');
  const coldAvgSize = avg(coldResults, 'pageSizeKB');
  const warmAvgSize = avg(warmResults, 'pageSizeKB');

  const avgFCP = avg(lhResults, 'fcp');
  const avgLCP = avg(lhResults, 'lcp');
  const avgTBT = avg(lhResults, 'tbt');
  const avgCLS = avg(lhResults, 'cls');
  const avgSI = avg(lhResults, 'speedIndex');

  const isGoodGrade = ['A', 'B', 'C'].includes(coldGrade);
  const isMediumGrade = ['D', 'E'].includes(coldGrade);
  const isBadGrade = ['F', 'G'].includes(coldGrade);

  const cacheImprovementEco = (coldAvgEco !== null && warmAvgEco !== null)
    ? Math.round((warmAvgEco - coldAvgEco) * 100) / 100 : null;
  const cacheImprovementSize = (coldAvgSize !== null && warmAvgSize !== null && coldAvgSize > 0)
    ? Math.round((1 - warmAvgSize / coldAvgSize) * 100) : null;

  // ─── ÖZET ───
  let ozet = `Bu rapor, ${institutionName} kurumuna ait ${url} web sitesinin çevresel performansını ve kullanıcı deneyimi performansını analiz etmektedir. GreenIT Analysis aracı ile cold cache ve warm cache koşullarında ${coldResults.length} adet ölçüm yapılarak EcoIndex, CO2 emisyonu ve su tüketimi metrikleri toplanmıştır. Lighthouse aracı ile ${lhResults.length} adet performans testi gerçekleştirilerek FCP, LCP, TBT, CLS ve Speed Index metrikleri ölçülmüştür. Cold cache koşullarında ortalama EcoIndex değeri ${fmtNum(coldAvgEco)} ("${coldGrade}") olarak belirlenmiştir. Warm cache koşullarında ise bu değer ${fmtNum(warmAvgEco)} ("${warmGrade}") seviyesine yükselmiştir. Lighthouse testlerinde ortalama FCP ${fmtNum(avgFCP, 0)} ms, LCP ${fmtNum(avgLCP, 0)} ms olarak ölçülmüştür.`;

  // ─── 6.1 Çevresel Performans ───
  let a61 = `${institutionName} web sitesinin (${url}) çevresel performansı GreenIT Analysis aracı ile değerlendirilmiştir. Cold cache koşullarında ortalama EcoIndex değeri ${fmtNum(coldAvgEco)} olup, bu değer "${coldGrade}" notuna karşılık gelmektedir. `;
  if (isGoodGrade) a61 += 'Bu sonuç, sitenin çevresel açıdan iyi bir performans sergilediğini göstermektedir. ';
  else if (isMediumGrade) a61 += 'Bu sonuç, sitenin çevresel performansının orta düzeyde olduğunu ve iyileştirmeye açık alanlar bulunduğunu göstermektedir. ';
  else if (isBadGrade) a61 += 'Bu sonuç, sitenin çevresel performansının düşük olduğunu ve acil iyileştirme gerektirdiğini göstermektedir. ';
  a61 += `Warm cache koşullarında ise ortalama EcoIndex ${fmtNum(warmAvgEco)} ("${warmGrade}") seviyesine yükselmiştir`;
  if (cacheImprovementEco !== null && cacheImprovementEco > 0) a61 += `, bu da cache mekanizmasının çevresel performansı ${fmtNum(cacheImprovementEco)} puan artırdığını ortaya koymaktadır`;
  a61 += '.';

  // ─── 6.2 Çevresel Ayak İzi ───
  let a62 = `Web sitesinin her ziyarette ürettiği ortalama karbon emisyonu cold cache koşullarında ${fmtNum(coldAvgCO2)} gCO2e, warm cache koşullarında ise ${fmtNum(warmAvgCO2)} gCO2e olarak ölçülmüştür. Su tüketimi açısından cold cache'de ${fmtNum(coldAvgWater)} cl, warm cache'de ${fmtNum(warmAvgWater)} cl değerleri gözlemlenmiştir. `;
  if (coldAvgCO2 !== null && coldAvgCO2 > 2.5) a62 += 'Bu değerler, sektör ortalamasının üzerinde bir çevresel ayak izi olduğuna işaret etmektedir. Sayfa optimizasyonu ile bu değerlerin düşürülmesi mümkündür.';
  else a62 += 'Bu değerler kabul edilebilir sınırlar içerisindedir ancak sürekli iyileştirme hedeflenmelidir.';

  // ─── 6.3 Sayfa Karmaşıklığı ───
  let a63 = `Sayfa DOM yapısı incelendiğinde, cold cache koşullarında ortalama ${fmtNum(coldAvgDom, 0)} DOM elemanı tespit edilmiştir. Warm cache koşullarında bu değer ${fmtNum(warmAvgDom, 0)} olarak ölçülmüştür. `;
  if (coldAvgDom !== null && coldAvgDom > 1500) a63 += 'DOM eleman sayısının yüksek olması, sayfanın karmaşık bir yapıya sahip olduğunu göstermektedir. Gereksiz HTML elemanlarının temizlenmesi, DOM ağacının sadeleştirilmesi önerilir. ';
  else if (coldAvgDom !== null && coldAvgDom > 800) a63 += 'DOM eleman sayısı orta seviyededir. Bazı optimizasyonlar ile daha verimli bir yapıya kavuşturulabilir. ';
  else a63 += 'DOM eleman sayısı makul seviyelerdedir ve sayfa yapısı nispeten sade görünmektedir. ';
  a63 += `Sayfa boyutu cold cache'de ortalama ${fmtNum(coldAvgSize)} KB olarak ölçülmüştür.`;

  // ─── 6.4 Veri transferi ───
  let a64 = `Cold cache koşullarında ortalama sayfa boyutu ${fmtNum(coldAvgSize)} KB iken, warm cache koşullarında bu değer ${fmtNum(warmAvgSize)} KB'a düşmüştür. `;
  if (cacheImprovementSize !== null && cacheImprovementSize > 0) a64 += `Cache mekanizması veri transferini yaklaşık %${cacheImprovementSize} oranında azaltmıştır. `;
  if (coldAvgSize !== null && coldAvgSize > 2000) a64 += "Sayfa boyutunun yüksek olması, büyük görseller, optimize edilmemiş CSS/JS dosyaları veya gereksiz kaynakların yüklenmesinden kaynaklanıyor olabilir. Sayfa boyutunun 1000 KB'ın altına düşürülmesi hedeflenmelidir.";
  else if (coldAvgSize !== null && coldAvgSize > 1000) a64 += 'Sayfa boyutu orta seviyededir. Görsel optimizasyonu ve gereksiz kaynak temizliği ile daha da azaltılabilir.';
  else a64 += 'Sayfa boyutu makul seviyelerdedir.';

  // ─── 6.5 HTTP istek sayısı ───
  let a65 = `Web sitesi cold cache koşullarında ortalama ${fmtNum(coldAvgReq, 0)} HTTP isteği yapmaktadır. Warm cache koşullarında bu sayı ${fmtNum(warmAvgReq, 0)} olarak gözlemlenmiştir. `;
  if (coldAvgReq !== null && coldAvgReq > 80) a65 += 'HTTP istek sayısının yüksek olması sayfa yükleme süresini ve enerji tüketimini olumsuz etkilemektedir. CSS ve JavaScript dosyalarının birleştirilmesi, sprite kullanımı ve gereksiz üçüncü parti scriptlerin kaldırılması önerilir.';
  else if (coldAvgReq !== null && coldAvgReq > 40) a65 += 'HTTP istek sayısı orta seviyededir. Kaynak birleştirme ve lazy loading teknikleri ile azaltılabilir.';
  else a65 += 'HTTP istek sayısı makul seviyelerdedir.';

  // ─── 7.1-7.7 İyileştirme Önerileri ───
  const recs = [];

  // 7.1
  let r1 = 'Sayfa boyutunun azaltılması hem performans hem de çevresel etki açısından kritik öneme sahiptir. ';
  if (coldAvgSize !== null && coldAvgSize > 1000) r1 += `Mevcut ortalama sayfa boyutu ${fmtNum(coldAvgSize)} KB olup, hedef değer olan 1000 KB'ın altına düşürülmelidir. `;
  r1 += 'CSS ve JavaScript dosyalarının minify edilmesi, kullanılmayan kodların kaldırılması (tree shaking), ve gzip/brotli sıkıştırma kullanılması önerilir. Ayrıca web fontlarının optimize edilmesi ve yalnızca gerekli karakter setlerinin yüklenmesi sayfa boyutunu önemli ölçüde azaltabilir.';
  recs.push(r1);

  // 7.2
  let r2 = 'Görsellerin HTML veya CSS ile yeniden boyutlandırılması yerine, sunucu tarafında doğru boyutlarda sunulması gerekmektedir. Tarayıcıda yeniden boyutlandırma gereksiz veri transferine neden olarak hem bant genişliği israfına hem de ek enerji tüketimine yol açar. Responsive images (<picture> elementi ve srcset attribute) kullanarak farklı ekran boyutları için uygun görsel boyutları sunulmalıdır.';
  recs.push(r2);

  // 7.3
  let r3 = 'Görsellerin modern formatlarda (WebP, AVIF) sunulması, dosya boyutlarını %25-50 oranında azaltabilir. Tüm görseller için uygun sıkıştırma seviyesi belirlenmelidir. Lazy loading uygulanarak, ekranda görünmeyen görsellerin yüklenmesi ertelenmelidir. SVG görseller için gereksiz metadata temizlenmeli ve minify edilmelidir.';
  recs.push(r3);

  // 7.4
  let r4 = 'Render-blocking CSS ve JavaScript kaynaklarının optimize edilmesi sayfa yükleme süresini önemli ölçüde iyileştirir. ';
  if (avgFCP !== null && avgFCP > 2000) r4 += `Mevcut FCP değeri (${fmtNum(avgFCP, 0)} ms) yüksek olup, critical CSS inline edilmeli ve geri kalan CSS asenkron yüklenmelidir. `;
  r4 += 'JavaScript dosyaları için defer veya async attribute kullanılmalı, code splitting uygulanmalıdır. Kullanılmayan CSS kuralları temizlenmeli ve CSS dosyaları minify edilmelidir.';
  recs.push(r4);

  // 7.5
  let r5 = 'HTTP istek sayısının azaltılması hem performans hem de enerji verimliliği açısından önemlidir. ';
  if (coldAvgReq !== null && coldAvgReq > 50) r5 += `Mevcut ortalama ${fmtNum(coldAvgReq, 0)} HTTP isteği yüksek olup azaltılmalıdır. `;
  r5 += 'CSS ve JavaScript dosyaları birleştirilmeli (bundling), CSS sprite teknikleri kullanılmalıdır. Gereksiz üçüncü parti scriptler gözden geçirilmeli ve mümkünse azaltılmalıdır.';
  recs.push(r5);

  // 7.6
  let r6 = 'Etkili bir cache stratejisi uygulanması, tekrarlayan ziyaretlerde veri transferini ve enerji tüketimini önemli ölçüde azaltır. ';
  if (cacheImprovementSize !== null) {
    r6 += `Mevcut cache mekanizması veri transferini %${cacheImprovementSize} oranında azaltmaktadır. `;
    if (cacheImprovementSize < 30) r6 += 'Bu oran düşük olup, cache politikasının gözden geçirilmesi gerekmektedir. ';
  }
  r6 += 'Statik kaynaklar için uzun süreli cache-control header\'ları ayarlanmalıdır. Service Worker kullanılarak offline deneyim ve cache yönetimi iyileştirilebilir.';
  recs.push(r6);

  // 7.7
  let r7 = 'DOM yapısının sadeleştirilmesi hem render performansını hem de enerji verimliliğini artırır. ';
  if (coldAvgDom !== null && coldAvgDom > 1000) r7 += `Mevcut ${fmtNum(coldAvgDom, 0)} DOM elemanı yüksek bir karmaşıklığa işaret etmektedir. `;
  r7 += 'Gereksiz wrapper div\'ler kaldırılmalı, semantic HTML elemanları tercih edilmelidir. Uzun listelerin virtual scrolling ile renderlanması düşünülmelidir. Inline style kullanımı minimize edilmeli, CSS class\'ları tercih edilmelidir.';
  recs.push(r7);

  // ─── 8. SONUÇ ───
  let sonuc = `${institutionName} web sitesinin çevresel performans ve kullanıcı deneyimi analizi tamamlanmıştır. GreenIT Analysis sonuçlarına göre site cold cache koşullarında "${coldGrade}" notu alırken, warm cache koşullarında "${warmGrade}" notuna yükselmiştir. Lighthouse testlerinde FCP ${fmtNum(avgFCP, 0)} ms, LCP ${fmtNum(avgLCP, 0)} ms olarak ölçülmüştür. Raporda sunulan iyileştirme önerileri uygulandığında, web sitesinin hem çevresel etkisinin azalması hem de kullanıcı deneyiminin iyileşmesi beklenmektedir. Bu iyileştirmeler uygulandıktan sonra testlerin tekrarlanarak ilerlemenin ölçülmesi tavsiye edilir.`;

  return { ozet, a61, a62, a63, a64, a65, recs, sonuc, coldGrade, warmGrade, coldAvgEco, warmAvgEco };
}

// ═══════════════════════════════════════
// ANA FONKSİYON
// ═══════════════════════════════════════

/**
 * Template DOCX'i temel alarak rapor oluştur.
 *
 * @param {string} outputPath - Çıktı dosya yolu (.docx)
 * @param {string} institutionName - Kurum adı
 * @param {string} url - Test edilen URL (ör: www.tedas.gov.tr)
 * @param {Array} coldResults - Cold cache GreenIT sonuçları
 * @param {Array} warmResults - Warm cache GreenIT sonuçları
 * @param {Array} lhResults - Lighthouse sonuçları
 * @param {string} [screenshotPath] - Web sitesi ekran görüntüsü dosya yolu (PNG)
 * @param {string} [authorName] - Raporu hazırlayan kişi adı
 */
export async function createDocxReport(outputPath, institutionName, url, coldResults, warmResults, lhResults, screenshotPath, authorName) {
  // Template'i oku
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(`Template bulunamadı: ${TEMPLATE_PATH}`);
  }

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);
  const zip = await JSZip.loadAsync(templateBuffer);

  // document.xml'i al
  let docXml = await zip.file('word/document.xml').async('string');

  // URL'den "www.xxx.yyy" formatını çıkar
  let urlClean = url.replace(/^https?:\/\//, '').replace(/\/+$/, '');

  // Metinleri üret
  const texts = generateAnalysisTexts(institutionName, urlClean, coldResults, warmResults, lhResults);

  // ═══════════════════════════════════════
  // 1. KURUM ADI DEĞİŞTİRME
  // ═══════════════════════════════════════
  // "Adalet Bakanlığı" → kurum adı
  // Türkçe karakterler farklı run'larda olduğundan, tüm w:t etiketlerini
  // birleştirip sonra tekrar ayırma yaklaşımı yerine, doğrudan metin değiştirme yapıyoruz.
  // "Adalet Bakanl" kısmı hep aynı run'da, ardından "ı", "ğ", "ı" ayrı run'larda.
  // Tüm bu run kümesini tek seferde değiştirelim.

  // Kapak sayfası text box'ları (72x72 ve 48x48 fontlu)
  // Cover text box 1 - büyük başlık (sz=72)
  docXml = replaceInstitutionRuns(docXml, institutionName, '72');
  // Cover text box 2 - fallback VML (sz=72)
  docXml = replaceInstitutionRuns(docXml, institutionName, '72');
  // İçerik başlığı (sz=40)
  docXml = replaceInstitutionRuns(docXml, institutionName, '40');

  // ═══════════════════════════════════════
  // 2. URL DEĞİŞTİRME
  // ═══════════════════════════════════════
  // www.adalet.gov.tr → yeni URL
  docXml = docXml.replace(/www\.adalet\.gov\.tr/g, escXml(urlClean));
  docXml = docXml.replace(/www\.adalet\.com\.tr/g, escXml(urlClean));

  // Hyperlink hedefini de güncelle (rels dosyasında)
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  relsXml = relsXml.replace(/https?:\/\/www\.adalet\.gov\.tr[^"]*/g, url);
  zip.file('word/_rels/document.xml.rels', relsXml);

  // ═══════════════════════════════════════
  // 2b. HAZIRLAYAN İSMİNİ DEĞİŞTİR
  // ═══════════════════════════════════════
  if (authorName) {
    docXml = docXml.replace(/Alaaddin KOYUNCU/g, escXml(authorName));
  }

  // ═══════════════════════════════════════
  // 3. EKRAN GÖRÜNTÜSÜNÜ DEĞİŞTİR
  // ═══════════════════════════════════════
  // image6.png (rId18) = GİRİŞ bölümündeki web sitesi ekran görüntüsü (Şekil 1)
  // Orijinal boyut: 2520x1279 px, Drawing extent: 6.30x3.20 inch (oran 1.97:1)
  if (screenshotPath && fs.existsSync(screenshotPath)) {
    const screenshotBuffer = fs.readFileSync(screenshotPath);
    zip.file('word/media/image6.png', screenshotBuffer);
  }

  // ═══════════════════════════════════════
  // 4. ÖZET ARALIK TABLOLARINI GÜNCELLE
  // ═══════════════════════════════════════
  // Tablo 0 (cold cache) ve Tablo 1 (warm cache): Sayfa Ağırlığı, Karmaşıklık, İstekler
  docXml = replaceRangeTable(docXml, 0, coldResults);
  docXml = replaceRangeTable(docXml, 1, warmResults);

  // ═══════════════════════════════════════
  // 5. TABLOLARI DOLDUR
  // ═══════════════════════════════════════
  const greenitColWidths = [726, 1784, 821, 1233, 986, 1423, 1319, 808, 773];
  const lhColWidths = [726, 1784, 1184, 1184, 1184, 1184, 1184];

  // Tüm büyük tabloları (15+ satır) ve özet tabloyu (5-14 satır) TEK SEFERDE bul,
  // sonra SONDAN BAŞA doğru değiştir (offset kayması olmaması için)
  docXml = replaceAllDataTables(docXml, coldResults, warmResults, lhResults, greenitColWidths, lhColWidths);

  // ═══════════════════════════════════════
  // 6. ŞEKİL BAŞLIKLARINI GÜNCELLE
  // ═══════════════════════════════════════
  // "Bakanlık internet sayfası görünümü" → kurum adı ile değiştir
  docXml = docXml.replace(
    /Bakanlık internet sayfası görünümü/g,
    escXml(`${institutionName} internet sayfası görünümü`)
  );

  // ═══════════════════════════════════════
  // 7. BÖLÜM METİNLERİNİ DEĞİŞTİR
  // ═══════════════════════════════════════
  docXml = replaceSectionText(docXml, 'ÖZET', texts.ozet);
  docXml = replaceSectionText(docXml, '6.1', texts.a61);
  docXml = replaceSectionText(docXml, '6.2', texts.a62);
  docXml = replaceSectionText(docXml, '6.3', texts.a63);
  docXml = replaceSectionText(docXml, '6.4', texts.a64);
  docXml = replaceSectionText(docXml, '6.5', texts.a65);
  docXml = replaceSectionText(docXml, '7.1', texts.recs[0]);
  docXml = replaceSectionText(docXml, '7.2', texts.recs[1]);
  docXml = replaceSectionText(docXml, '7.3', texts.recs[2]);
  docXml = replaceSectionText(docXml, '7.4', texts.recs[3]);
  docXml = replaceSectionText(docXml, '7.5', texts.recs[4]);
  docXml = replaceSectionText(docXml, '7.6', texts.recs[5]);
  docXml = replaceSectionText(docXml, '7.7', texts.recs[6]);
  docXml = replaceSectionText(docXml, '8. SONUÇ', texts.sonuc);

  // ═══════════════════════════════════════
  // 7. KAYDET
  // ═══════════════════════════════════════
  zip.file('word/document.xml', docXml);

  const outputBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });

  fs.writeFileSync(outputPath, outputBuffer);
}

// ═══════════════════════════════════════
// KURUM ADI RUN DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * "Adalet Bakanl" + "ı" + "ğ" + "ı" run kümesini tek bir run ile değiştirir.
 * fontSize ile hangi text box'ta olduğunu ayırt eder.
 */
function replaceInstitutionRuns(xml, newName, fontSize) {
  // Adalet Bakanlığı pattern: "Adalet Bakanl" run, ardından "ı", "ğ", "ı" run'ları
  // Bu pattern her üç yerde de aynı yapıda (sadece font size farklı)

  // Regex: "Adalet Bakanl" içeren run + sonraki 3 run (ı, ğ, ı)
  const pattern = new RegExp(
    '(<w:r>\\s*<w:rPr>[\\s\\S]*?<w:sz w:val="' + fontSize + '"/>[\\s\\S]*?</w:rPr>\\s*<w:t>)Adalet Bakanl(</w:t>\\s*</w:r>)' +
    '\\s*<w:r>\\s*<w:rPr>[\\s\\S]*?</w:rPr>\\s*<w:t>ı</w:t>\\s*</w:r>' +
    '\\s*<w:r>\\s*<w:rPr>[\\s\\S]*?</w:rPr>\\s*<w:t>ğ</w:t>\\s*</w:r>' +
    '\\s*<w:r>\\s*<w:rPr>[\\s\\S]*?</w:rPr>\\s*<w:t>ı</w:t>\\s*</w:r>',
    ''
  );

  return xml.replace(pattern, `$1${escXml(newName)}$2`);
}

// ═══════════════════════════════════════
// ÖZET ARALIK TABLOSU DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * N'inci küçük tabloyu (5 satır, Sayfa Ağırlığı/Karmaşıklık/İstekler) bul
 * ve Medyan değerlerini gerçek verilerle değiştir.
 * tableIndex: 0 = cold cache, 1 = warm cache
 */
function replaceRangeTable(xml, tableIndex, results) {
  const tblRegex = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
  let match;
  let smallTableIdx = 0;

  while ((match = tblRegex.exec(xml)) !== null) {
    const tblContent = match[0];
    const rowCount = (tblContent.match(/<w:tr[\s>]/g) || []).length;

    // 5 satırlık tablolar = özet aralık tabloları
    if (rowCount === 5) {
      if (smallTableIdx === tableIndex) {
        // Ortalama değerleri hesapla
        const avgPageSizeMB = avg(results, 'pageSizeKB') !== null
          ? (avg(results, 'pageSizeKB') / 1024).toFixed(3)
          : '-';
        const avgDom = avg(results, 'domSize') !== null
          ? Math.round(avg(results, 'domSize')).toString()
          : '-';
        const avgReq = avg(results, 'requestCount') !== null
          ? Math.round(avg(results, 'requestCount')).toString()
          : '-';

        // Medyan sütunundaki değerleri değiştir
        // Row 2: "2.410" → avgPageSizeMB
        // Row 3: "693" → avgDom
        // Row 4: "78" → avgReq
        let newTbl = tblContent;

        // Sayfa ağırlığı medyan değerini bul ve değiştir
        // Bu satırdaki Medyan hücresindeki sayısal değeri bul
        newTbl = replaceNthCellValue(newTbl, 2, 2, avgPageSizeMB);
        // Karmaşıklık medyan değeri
        newTbl = replaceNthCellValue(newTbl, 3, 2, avgDom);
        // İstekler medyan değeri
        newTbl = replaceNthCellValue(newTbl, 4, 2, avgReq);

        xml = xml.substring(0, match.index) + newTbl + xml.substring(match.index + match[0].length);
        return xml;
      }
      smallTableIdx++;
    }
  }

  return xml;
}

/**
 * Bir tablodaki belirli satır ve hücredeki sayısal değeri değiştirir.
 * rowIdx: satır indeksi (0-based), cellIdx: hücre indeksi (0-based)
 */
function replaceNthCellValue(tableXml, rowIdx, cellIdx, newValue) {
  const rows = [];
  const rowRegex = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
  let m;
  while ((m = rowRegex.exec(tableXml)) !== null) {
    rows.push({ start: m.index, end: m.index + m[0].length, content: m[0] });
  }

  if (rowIdx >= rows.length) return tableXml;

  const row = rows[rowIdx];
  const cells = [];
  const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/g;
  while ((m = cellRegex.exec(row.content)) !== null) {
    cells.push({ start: m.index, end: m.index + m[0].length, content: m[0] });
  }

  if (cellIdx >= cells.length) return tableXml;

  const cell = cells[cellIdx];

  // Hücredeki w:t etiketinin değerini değiştir
  // Eğer hücrede w:t varsa, içindeki sayısal değeri newValue ile değiştir
  const wtMatch = cell.content.match(/<w:t[^>]*>([^<]+)<\/w:t>/);
  if (wtMatch) {
    const oldCellXml = cell.content;
    const newCellXml = oldCellXml.replace(
      /<w:t([^>]*)>[^<]+<\/w:t>/,
      `<w:t$1>${escXml(String(newValue))}</w:t>`
    );
    const newRow = row.content.substring(0, cell.start) + newCellXml + row.content.substring(cell.end);
    tableXml = tableXml.substring(0, row.start) + newRow + tableXml.substring(row.end);
  }

  return tableXml;
}

// ═══════════════════════════════════════
// TABLO SATIRI DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * Tüm veri tablolarını TEK SEFERDE bul ve SONDAN BAŞA değiştir.
 * Bu yaklaşım, tablo boyutları değiştiğinde offset kaymasını önler.
 *
 * Template'deki tablo sırası (sabit):
 *   Tablo 0: 5 satır - cold cache aralık tablosu (replaceRangeTable ile ayrıca işleniyor)
 *   Tablo 1: 5 satır - warm cache aralık tablosu (replaceRangeTable ile ayrıca işleniyor)
 *   Tablo 2: 22 satır - Cold Cache GreenIT verileri     → bigTable[0]
 *   Tablo 3: 22 satır - Warm Cache GreenIT verileri      → bigTable[1]
 *   Tablo 4: 22 satır - Lighthouse verileri               → bigTable[2]
 *   Tablo 5: 7 satır  - Özet/Ortalama tablosu             → summaryTable
 */
function replaceAllDataTables(xml, coldResults, warmResults, lhResults, greenitColWidths, lhColWidths) {
  // 1. Önce TÜM tabloları bul ve bilgilerini kaydet
  const tblRegex = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
  let match;
  const allTables = [];

  while ((match = tblRegex.exec(xml)) !== null) {
    const content = match[0];
    const rowCount = (content.match(/<w:tr[\s>]/g) || []).length;
    allTables.push({
      index: match.index,
      length: content.length,
      content,
      rows: rowCount
    });
  }

  // 2. Büyük tabloları (15+ satır) ve özet tabloyu (5-14 satır, büyüklerden sonra) ayır
  const bigTables = [];
  let summaryTable = null;

  for (let i = 0; i < allTables.length; i++) {
    if (allTables[i].rows >= 15) {
      bigTables.push({ ...allTables[i], origIdx: i });
    }
  }

  // Özet tablo: 3 büyük tablodan sonra gelen ilk orta boy tablo (5-14 satır)
  for (let i = 0; i < allTables.length; i++) {
    if (allTables[i].rows >= 5 && allTables[i].rows < 15) {
      const bigBefore = allTables.filter((t, j) => j < i && t.rows >= 15).length;
      if (bigBefore >= 3) {
        summaryTable = { ...allTables[i], origIdx: i };
        break;
      }
    }
  }

  // 3. Değiştirme planını oluştur (position, yeni içerik)
  const replacements = [];

  // Büyük tablo 0 = Cold Cache GreenIT
  if (bigTables[0]) {
    const newContent = rebuildDataTable(bigTables[0].content, coldResults, greenitColWidths, 'greenit');
    replacements.push({ index: bigTables[0].index, length: bigTables[0].length, newContent });
  }

  // Büyük tablo 1 = Warm Cache GreenIT
  if (bigTables[1]) {
    const newContent = rebuildDataTable(bigTables[1].content, warmResults, greenitColWidths, 'greenit');
    replacements.push({ index: bigTables[1].index, length: bigTables[1].length, newContent });
  }

  // Büyük tablo 2 = Lighthouse
  if (bigTables[2]) {
    const newContent = rebuildDataTable(bigTables[2].content, lhResults, lhColWidths, 'lighthouse');
    replacements.push({ index: bigTables[2].index, length: bigTables[2].length, newContent });
  }

  // Özet tablo
  if (summaryTable) {
    const newContent = rebuildSummaryTable(summaryTable.content, coldResults, warmResults);
    replacements.push({ index: summaryTable.index, length: summaryTable.length, newContent });
  }

  // 4. SONDAN BAŞA doğru değiştir (offset kayması olmasın)
  replacements.sort((a, b) => b.index - a.index);

  for (const rep of replacements) {
    xml = xml.substring(0, rep.index) + rep.newContent + xml.substring(rep.index + rep.length);
  }

  return xml;
}

function rebuildDataTable(tableXml, results, colWidths, type) {
  // Header kısmını koru (tblPr + tblGrid + ilk w:tr)
  const firstTrEndIdx = tableXml.indexOf('</w:tr>') + '</w:tr>'.length;
  const headerPart = tableXml.substring(0, firstTrEndIdx);

  // Yeni data satırlarını oluştur
  let dataRows;
  if (type === 'greenit') {
    dataRows = results.map(r => greenitDataRow(r, colWidths));
  } else {
    dataRows = results.map(r => lighthouseDataRow(r, colWidths));
  }

  return headerPart + '\n' + dataRows.join('\n') + '\n    </w:tbl>';
}

function rebuildSummaryTable(tableXml, coldResults, warmResults) {
  const coldAvgEco = avg(coldResults, 'ecoIndex');
  const warmAvgEco = avg(warmResults, 'ecoIndex');
  const coldGrade = coldAvgEco !== null ? getEcoIndexGrade(coldAvgEco) : '-';
  const warmGrade = warmAvgEco !== null ? getEcoIndexGrade(warmAvgEco) : '-';

  const metrics = [
    ['EcoIndex', fmtNum(coldAvgEco), fmtNum(warmAvgEco)],
    ['Grade', coldGrade, warmGrade],
    ['CO2 (gCO2e)', fmtNum(avg(coldResults, 'co2')), fmtNum(avg(warmResults, 'co2'))],
    ['Water (cl)', fmtNum(avg(coldResults, 'water')), fmtNum(avg(warmResults, 'water'))],
    ['Request Count', fmtNum(avg(coldResults, 'requestCount'), 0), fmtNum(avg(warmResults, 'requestCount'), 0)],
    ['Page Size (KB)', fmtNum(avg(coldResults, 'pageSizeKB')), fmtNum(avg(warmResults, 'pageSizeKB'))],
    ['DOM Size', fmtNum(avg(coldResults, 'domSize'), 0), fmtNum(avg(warmResults, 'domSize'), 0)]
  ];

  // Header koru
  const firstTrEndIdx = tableXml.indexOf('</w:tr>') + '</w:tr>'.length;
  const headerPart = tableXml.substring(0, firstTrEndIdx);

  const sumColWidths = [2400, 2200, 2200];
  const dataRows = metrics.map(([label, cold, warm]) =>
    makeTableRow([label, cold, warm], sumColWidths)
  );

  return headerPart + '\n' + dataRows.join('\n') + '\n    </w:tbl>';
}

// ═══════════════════════════════════════
// BÖLÜM METNİ DEĞİŞTİRME
// ═══════════════════════════════════════

/**
 * Bir başlık (Heading) altındaki paragrafların metnini değiştirir.
 * Başlığı tanıyıp, bir sonraki başlığa kadar olan paragrafları bulur
 * ve ilk body paragrafın metnini yeni metinle değiştirir, diğerlerini kaldırır.
 */
function replaceSectionText(xml, sectionId, newText) {
  // Bölüm başlığını bul
  let headingPattern;
  if (sectionId === 'ÖZET') {
    headingPattern = /<w:t>ÖZET<\/w:t>/;
  } else if (sectionId === '8. SONUÇ') {
    headingPattern = /<w:t>8\. SONUÇ<\/w:t>/;
  } else {
    headingPattern = new RegExp(`<w:t>${sectionId.replace('.', '\\.')}[^<]*</w:t>`);
  }

  const headingMatch = headingPattern.exec(xml);
  if (!headingMatch) return xml;

  const headingPos = headingMatch.index;
  const headingPEnd = xml.indexOf('</w:p>', headingPos) + '</w:p>'.length;

  const afterHeading = xml.substring(headingPEnd);

  // Boundary markers: sonraki heading, tablo, sectPr veya body sonu
  const markers = [];
  const nextSection = /<w:pStyle w:val="Balk[12]"\/>/.exec(afterHeading);
  const nextTable = /<w:tbl>/.exec(afterHeading);
  const nextSectPr = /<w:sectPr/.exec(afterHeading);
  const nextBodyEnd = /<\/w:body>/.exec(afterHeading);

  if (nextSection) markers.push(nextSection.index);
  if (nextTable) markers.push(nextTable.index);
  if (nextSectPr) markers.push(nextSectPr.index);
  if (nextBodyEnd) markers.push(nextBodyEnd.index);

  let endOffset;
  if (markers.length > 0) {
    const markerOffset = Math.min(...markers);
    // CRITICAL: Marker'dan geriye giderek <w:p başlangıcını bul
    const beforeMarker = afterHeading.substring(0, markerOffset);
    let lastWP = beforeMarker.lastIndexOf('<w:p ');
    if (lastWP === -1) lastWP = beforeMarker.lastIndexOf('<w:p>');
    endOffset = lastWP >= 0 ? lastWP : markerOffset;
  } else {
    endOffset = afterHeading.length;
  }

  const actualEnd = headingPEnd + endOffset;

  const newParagraph = `
    <w:p w14:paraId="${randomParaId()}" w14:textId="${randomParaId()}" w:rsidR="00530778" w:rsidRDefault="00530778" w:rsidP="00A23D3B">
      <w:pPr>
        <w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/>
        <w:jc w:val="both"/>
        <w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">${escXml(newText)}</w:t></w:r>
    </w:p>`;

  xml = xml.substring(0, headingPEnd) + newParagraph + xml.substring(actualEnd);
  return xml;
}
