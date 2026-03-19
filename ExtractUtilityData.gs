/**
 * ════════════════════════════════════════════════════════════════
 *  J.Residence — Extract MEA/MWA Data from PDFs in Google Drive
 *  อ่านข้อมูลจาก PDF ใบแจ้งหนี้/ใบเสร็จ แล้ว output เป็น JSON
 *  สำหรับ import เข้า Dashboard (sec-utility)
 * ════════════════════════════════════════════════════════════════
 *
 *  วิธีใช้:
 *  1. วางโค้ดนี้ใน Google Apps Script project เดียวกับ runAll()
 *  2. รัน runAll() ก่อน (ให้ PDFs ถูกบันทึกใน Drive แล้ว)
 *  3. เลือก "extractAll" → ▶ Run
 *  4. เปิด Execution Log → copy JSON → วางใน Dashboard "📥 Import JSON"
 *
 *  หรือรัน "extractToSheet" เพื่อ export เป็น Google Sheet
 * ════════════════════════════════════════════════════════════════
 */

// ── Root folder name (ต้องตรงกับ ROOT_NAME ใน runAll script) ──
var EXTRACT_ROOT_NAME = '📁 J.Residence — สาธารณูปโภค';

// ══════════════════════════════════════════════════════════════
//  MAIN — Extract all PDF data and output JSON
// ══════════════════════════════════════════════════════════════
function extractAll() {
  Logger.log('══════════════════════════════════════');
  Logger.log(' Extract MEA/MWA Data from PDFs');
  Logger.log('══════════════════════════════════════\n');

  var result = {};
  var stats = { total: 0, success: 0, fail: 0 };

  // Find root folder
  var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
  if (!roots.hasNext()) {
    Logger.log('❌ ไม่พบ folder: ' + EXTRACT_ROOT_NAME);
    Logger.log('   รัน runAll() ก่อนเพื่อสร้าง folder structure');
    return;
  }
  var rootFolder = roots.next();

  // Process MEA
  var meaFolders = rootFolder.getFoldersByName('⚡ MEA — การไฟฟ้านครหลวง (012076042)');
  if (meaFolders.hasNext()) {
    var meaRoot = meaFolders.next();
    Logger.log('⚡ Processing MEA...');
    processUtilityFolder_(meaRoot, 'mea', result, stats);
  }

  // Process MWA
  var mwaFolders = rootFolder.getFoldersByName('💧 MWA — การประปานครหลวง (74475021)');
  if (mwaFolders.hasNext()) {
    var mwaRoot = mwaFolders.next();
    Logger.log('\n💧 Processing MWA...');
    processUtilityFolder_(mwaRoot, 'mwa', result, stats);
  }

  // Output
  Logger.log('\n══════════════════════════════════════');
  Logger.log(' สรุป: สำเร็จ ' + stats.success + '/' + stats.total + ' ไฟล์');
  Logger.log('══════════════════════════════════════');

  var json = JSON.stringify(result, null, 2);
  Logger.log('\n══════════════════════════════════════');
  Logger.log(' JSON สำหรับ Import (copy ทั้งหมดด้านล่าง)');
  Logger.log('══════════════════════════════════════');
  Logger.log(json);

  return result;
}

// ══════════════════════════════════════════════════════════════
//  Process utility folder tree
// ══════════════════════════════════════════════════════════════
function processUtilityFolder_(utilityFolder, type, result, stats) {
  var yearFolders = utilityFolder.getFolders();
  while (yearFolders.hasNext()) {
    var yearFolder = yearFolders.next();
    var monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      var monthFolder = monthFolders.next();
      var files = monthFolder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        if (file.getMimeType() !== 'application/pdf') continue;
        stats.total++;
        try {
          var data = extractPdfData_(file, type);
          if (data && data.key) {
            // Merge with existing (invoice + receipt for same month)
            var existing = result[data.key] || {};
            for (var k in data.values) {
              if (data.values[k] && !existing[k]) {
                existing[k] = data.values[k];
              }
              // Prefer receipt amount over invoice amount
              if (k === 'amount' && data.isReceipt && data.values[k]) {
                existing[k] = data.values[k];
              }
            }
            result[data.key] = existing;
            stats.success++;
            Logger.log('  ✅ ' + file.getName());
          } else {
            stats.fail++;
            Logger.log('  ⚠ ไม่สามารถ parse: ' + file.getName());
          }
        } catch (e) {
          stats.fail++;
          Logger.log('  ❌ ' + file.getName() + ': ' + e.message);
        }
      }
    }
  }
}

// ══════════════════════════════════════════════════════════════
//  Extract text from PDF via Drive OCR conversion
// ══════════════════════════════════════════════════════════════
function extractPdfText_(file) {
  // Convert PDF to Google Doc (OCR) temporarily
  var blob = file.getBlob();
  var resource = {
    title: '__temp_extract_' + file.getId(),
    mimeType: 'application/pdf'
  };

  // Use Drive API v2 to insert with OCR
  var docFile = Drive.Files.insert(resource, blob, {
    ocr: true,
    ocrLanguage: 'th'
  });

  // Read text from the created doc
  var doc = DocumentApp.openById(docFile.id);
  var text = doc.getBody().getText();

  // Delete temp doc
  DriveApp.getFileById(docFile.id).setTrashed(true);

  return text;
}

// ══════════════════════════════════════════════════════════════
//  Extract data from PDF file
// ══════════════════════════════════════════════════════════════
function extractPdfData_(file, type) {
  var fileName = file.getName();
  var text = extractPdfText_(file);
  if (!text || text.length < 50) return null;

  var isReceipt = /Receipt|ใบเสร็จ/i.test(fileName) ||
                  /ใบเสร็จรับเงิน|RECEIPT/i.test(text);

  // Detect month from filename (e.g., MEA_Invoice_ม.ค._2569_...)
  var monthKey = detectMonthKey_(fileName, text, type);
  if (!monthKey) return null;

  var values = {};

  if (type === 'mea') {
    values = parseMeaText_(text);
  } else {
    values = parseMwaText_(text);
  }

  return {
    key: monthKey,
    values: values,
    isReceipt: isReceipt
  };
}

// ══════════════════════════════════════════════════════════════
//  Parse MEA PDF text
// ══════════════════════════════════════════════════════════════
function parseMeaText_(text) {
  var data = {};

  // ── ยอดเงิน (Amount) ──
  // Pattern: จำนวนเงินรวม / ยอดเงินรวม / TOTAL
  var amtPatterns = [
    /(?:จำนวนเงินรวม|ยอดเงินรวม|ยอดรวม(?:ทั้งสิ้น)?|จำนวนเงินที่ต้องชำระ|TOTAL\s*AMOUNT|NET\s*AMOUNT|AMOUNT\s*DUE)[^\d]{0,30}([\d,]+\.?\d*)/i,
    /(?:รวมภาษีมูลค่าเพิ่ม|รวมเงิน|ยอดชำระ)[^\d]{0,20}([\d,]+\.\d{2})/i
  ];
  for (var i = 0; i < amtPatterns.length; i++) {
    var m = text.match(amtPatterns[i]);
    if (m) {
      var amt = m[1].replace(/,/g, '');
      if (parseFloat(amt) > 1000) { // MEA bills should be > 1000
        data.amount = amt;
        break;
      }
    }
  }

  // ── Off Peak units ──
  var offPatterns = [
    /(?:OFF[\s\-]*PEAK|นอกช่วง(?:ความต้องการสูง|เวลาพีค))[^\d]{0,40}([\d,]+(?:\.\d+)?)\s*(?:หน่วย|unit|kWh)?/i,
    /OFF[\s\-]*PEAK[^\d]{0,15}([\d,]+)/i
  ];
  for (var i = 0; i < offPatterns.length; i++) {
    var m = text.match(offPatterns[i]);
    if (m) { data.offpeak = m[1].replace(/,/g, ''); break; }
  }

  // ── On Peak units ──
  var onPatterns = [
    /(?:ON[\s\-]*PEAK|ช่วง(?:ความต้องการสูง|เวลาพีค))[^\d]{0,40}([\d,]+(?:\.\d+)?)\s*(?:หน่วย|unit|kWh)?/i,
    /ON[\s\-]*PEAK[^\d]{0,15}([\d,]+)/i
  ];
  for (var i = 0; i < onPatterns.length; i++) {
    var m = text.match(onPatterns[i]);
    if (m) { data.onpeak = m[1].replace(/,/g, ''); break; }
  }

  // ── Total units (auto-calc if off+on available) ──
  if (data.offpeak && data.onpeak) {
    data.totalUnits = String(parseFloat(data.offpeak) + parseFloat(data.onpeak));
  }

  // ── วันที่จดมิเตอร์ ──
  var mDatePatterns = [
    /(?:วัน(?:ที่)?จดมิเตอร์|วันที่อ่าน(?:มิเตอร์)?|METER\s*READ(?:ING)?\s*DATE)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i,
    /(?:วันที่อ่าน)[^\d]{0,10}(\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4})/i
  ];
  for (var i = 0; i < mDatePatterns.length; i++) {
    var m = text.match(mDatePatterns[i]);
    if (m) { data.meterDate = m[1].replace(/\s/g, ''); break; }
  }

  // ── รอบบิล (Billing period) ──
  var periodPatterns = [
    /(?:งวดที่|รอบ(?:การ)?(?:ใช้|บิล)|BILLING\s*PERIOD)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})\s*[-–ถึง\s]+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i,
    /(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})\s*[-–]\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/
  ];
  for (var i = 0; i < periodPatterns.length; i++) {
    var m = text.match(periodPatterns[i]);
    if (m) { data.period = m[1] + '–' + m[2]; break; }
  }

  // ── วันที่ออกบิล ──
  var billPatterns = [
    /(?:วันที่(?:ออก)?บิล|วันที่(?:ออก|พิมพ์)|BILL\s*DATE|ISSUE\s*DATE|PRINT\s*DATE)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i
  ];
  for (var i = 0; i < billPatterns.length; i++) {
    var m = text.match(billPatterns[i]);
    if (m) { data.paidDate = m[1]; break; }
  }

  // ── เลขมิเตอร์เก่า/ใหม่ (demand) ──
  var demandM = text.match(/(?:DEMAND|ดีมานด์|กำลังไฟสูงสุด)[^\d]{0,15}([\d,]+(?:\.\d+)?)\s*(?:kW)?/i);
  if (demandM) data.demand = demandM[1].replace(/,/g, '');

  // ── Power Factor ──
  var pfM = text.match(/(?:POWER\s*FACTOR|ตัวประกอบกำลัง|PF)[^\d]{0,15}([\d.]+)\s*%?/i);
  if (pfM) data.powerFactor = pfM[1];

  return data;
}

// ══════════════════════════════════════════════════════════════
//  Parse MWA PDF text
// ══════════════════════════════════════════════════════════════
function parseMwaText_(text) {
  var data = {};

  // ── ยอดเงิน ──
  var amtPatterns = [
    /(?:จำนวนเงินรวม|ยอดเงินรวม|ยอดรวม(?:ทั้งสิ้น)?|TOTAL|รวมเงิน)[^\d]{0,30}([\d,]+\.?\d*)/i,
    /(?:รวมภาษีมูลค่าเพิ่ม|ยอดชำระ)[^\d]{0,20}([\d,]+\.\d{2})/i,
    /([\d,]+\.\d{2})\s*(?:บาท|BAHT)/i
  ];
  for (var i = 0; i < amtPatterns.length; i++) {
    var m = text.match(amtPatterns[i]);
    if (m) {
      var amt = m[1].replace(/,/g, '');
      if (parseFloat(amt) > 100) { // MWA bills usually > 100
        data.amount = amt;
        break;
      }
    }
  }

  // ── หน่วยน้ำ (ลบ.ม.) ──
  var unitPatterns = [
    /(?:ปริมาณ(?:น้ำ)?(?:ที่)?ใช้|จำนวน(?:น้ำ)?ที่ใช้|CONSUMPTION|USAGE|ปริมาณการใช้น้ำ)[^\d]{0,25}([\d,]+(?:\.\d+)?)\s*(?:ลบ\.?ม\.?|ลูกบาศก์|cu\.?\s*m|m3)?/i,
    /(?:จำนวน|ปริมาณ)\s*(?:\(ลบ\.ม\.\))?\s*[^\d]{0,10}([\d,]+)\s*(?:ลบ\.?ม\.?|ลูกบาศก์)?/i,
    /([\d]+)\s*ลบ\.?ม/i
  ];
  for (var i = 0; i < unitPatterns.length; i++) {
    var m = text.match(unitPatterns[i]);
    if (m) { data.units = m[1].replace(/,/g, ''); break; }
  }

  // ── มิเตอร์เก่า ──
  var oldPatterns = [
    /(?:เลข(?:มิเตอร์|มาตร)เก่า|มาตรเดิม|(?:เลข)?อ่านครั้งก่อน|PREVIOUS\s*(?:READ(?:ING)?|METER))[^\d]{0,20}([\d,]+)/i,
    /(?:ครั้งก่อน|เดิม)\s*[^\d]{0,10}([\d,]+)/i
  ];
  for (var i = 0; i < oldPatterns.length; i++) {
    var m = text.match(oldPatterns[i]);
    if (m) { data.meterOld = m[1].replace(/,/g, ''); break; }
  }

  // ── มิเตอร์ใหม่ ──
  var newPatterns = [
    /(?:เลข(?:มิเตอร์|มาตร)(?:ใหม่|ปัจจุบัน)|มาตร(?:ใหม่|ปัจจุบัน)|(?:เลข)?อ่านครั้ง(?:นี้|ล่าสุด)|CURRENT\s*(?:READ(?:ING)?|METER)|PRESENT\s*(?:READ(?:ING)?|METER))[^\d]{0,20}([\d,]+)/i,
    /(?:ครั้ง(?:นี้|หลัง)|ปัจจุบัน)\s*[^\d]{0,10}([\d,]+)/i
  ];
  for (var i = 0; i < newPatterns.length; i++) {
    var m = text.match(newPatterns[i]);
    if (m) { data.meterNew = m[1].replace(/,/g, ''); break; }
  }

  // ── Auto-calc units ──
  if (data.meterOld && data.meterNew && !data.units) {
    var diff = parseFloat(data.meterNew) - parseFloat(data.meterOld);
    if (diff > 0) data.units = String(diff);
  }

  // ── วันจดมิเตอร์ ──
  var mDatePatterns = [
    /(?:วัน(?:ที่)?จดมิเตอร์|วัน(?:ที่)?อ่าน(?:มิเตอร์|มาตร)?|METER\s*READ(?:ING)?\s*DATE)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i
  ];
  for (var i = 0; i < mDatePatterns.length; i++) {
    var m = text.match(mDatePatterns[i]);
    if (m) { data.meterDate = m[1]; break; }
  }

  // ── รอบบิล ──
  var periodPatterns = [
    /(?:งวด(?:ที่)?|รอบ(?:บิล|การใช้)?|BILLING\s*PERIOD)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})\s*[-–ถึง\s]+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i
  ];
  for (var i = 0; i < periodPatterns.length; i++) {
    var m = text.match(periodPatterns[i]);
    if (m) { data.period = m[1] + '–' + m[2]; break; }
  }

  // ── วันออกใบเสร็จ ──
  var rcDatePatterns = [
    /(?:วันที่(?:ออก)?(?:ใบเสร็จ|รับชำระ|ชำระ)|RECEIPT\s*DATE|PAYMENT\s*DATE)[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i,
    /(?:วันที่)\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i
  ];
  for (var i = 0; i < rcDatePatterns.length; i++) {
    var m = text.match(rcDatePatterns[i]);
    if (m) { data.paidDate = m[1]; break; }
  }

  return data;
}

// ══════════════════════════════════════════════════════════════
//  Detect month key from filename / text
// ══════════════════════════════════════════════════════════════
function detectMonthKey_(fileName, text, type) {
  var TH_SHORT = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                      'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

  // Try from filename: MEA_Invoice_ม.ค._2569_...
  for (var m = 1; m <= 12; m++) {
    var short = TH_SHORT[m];
    var esc = short.replace(/\./g, '\\.');
    var re = new RegExp(esc + '\\s*(25\\d{2})');
    var match = fileName.match(re);
    if (match) {
      return type + '_' + short + '_' + match[1];
    }
  }

  // Try from filename: patterns like _2569_01_ or _2568_12_
  var numMatch = fileName.match(/_(25\d{2})_(0[1-9]|1[0-2])_/);
  if (numMatch) {
    var yr = numMatch[1];
    var mo = parseInt(numMatch[2]);
    return type + '_' + TH_SHORT[mo] + '_' + yr;
  }

  // Try from text: Thai full month names
  var TH_FULL = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  for (var m = 1; m <= 12; m++) {
    var reF = new RegExp(TH_FULL[m] + '\\s*(25\\d{2})');
    var reS = new RegExp(TH_SHORT[m].replace(/\./g, '\\.') + '\\s*(25\\d{2})');
    var mF = text.match(reF);
    var mS = text.match(reS);
    var found = mF || mS;
    if (found) {
      return type + '_' + TH_SHORT[m] + '_' + found[1];
    }
  }

  return null;
}

// ══════════════════════════════════════════════════════════════
//  OPTIONAL: Extract to Google Sheet
// ══════════════════════════════════════════════════════════════
function extractToSheet() {
  var result = extractAll();
  if (!result) return;

  var ss = SpreadsheetApp.create('J.Residence — Utility Data Extract');

  // MEA Sheet
  var meaSheet = ss.getActiveSheet();
  meaSheet.setName('MEA');
  meaSheet.appendRow(['เดือน', 'ยอดรวม (฿)', 'Off Peak', 'On Peak', 'หน่วยรวม',
                      'วันจดมิเตอร์', 'รอบบิล', 'วันออกบิล', 'Demand', 'PF']);
  var meaKeys = Object.keys(result).filter(function(k) { return k.startsWith('mea_'); }).sort();
  meaKeys.forEach(function(k) {
    var d = result[k];
    var month = k.replace('mea_', '').replace(/_/g, ' ');
    meaSheet.appendRow([month, d.amount||'', d.offpeak||'', d.onpeak||'', d.totalUnits||'',
                        d.meterDate||'', d.period||'', d.paidDate||'', d.demand||'', d.powerFactor||'']);
  });

  // MWA Sheet
  var mwaSheet = ss.insertSheet('MWA');
  mwaSheet.appendRow(['เดือน', 'ยอดรวม (฿)', 'หน่วย (ลบ.ม.)', 'มิเตอร์เก่า', 'มิเตอร์ใหม่',
                      'วันจดมิเตอร์', 'รอบบิล', 'วันออกใบเสร็จ']);
  var mwaKeys = Object.keys(result).filter(function(k) { return k.startsWith('mwa_'); }).sort();
  mwaKeys.forEach(function(k) {
    var d = result[k];
    var month = k.replace('mwa_', '').replace(/_/g, ' ');
    mwaSheet.appendRow([month, d.amount||'', d.units||'', d.meterOld||'', d.meterNew||'',
                        d.meterDate||'', d.period||'', d.paidDate||'']);
  });

  Logger.log('\n📊 Google Sheet: ' + ss.getUrl());
  Logger.log('   ชื่อ: ' + ss.getName());
}

// ══════════════════════════════════════════════════════════════
//  OPTIONAL: Extract single file (for testing)
// ══════════════════════════════════════════════════════════════
function testExtractSingle() {
  // เปลี่ยน fileId เป็น ID ของ PDF ที่ต้องการทดสอบ
  var fileId = 'PASTE_FILE_ID_HERE';
  var file = DriveApp.getFileById(fileId);
  var text = extractPdfText_(file);

  Logger.log('═══ Raw Text ═══');
  Logger.log(text);

  Logger.log('\n═══ Parsed MEA ═══');
  Logger.log(JSON.stringify(parseMeaText_(text), null, 2));

  Logger.log('\n═══ Parsed MWA ═══');
  Logger.log(JSON.stringify(parseMwaText_(text), null, 2));
}

// ══════════════════════════════════════════════════════════════
//  OPTIONAL: Get all Drive File IDs for index.html
//  (อัปเดต _UT_DRIVE_FILE_IDS ใน dashboard)
// ══════════════════════════════════════════════════════════════
function getDriveFileIds() {
  var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
  if (!roots.hasNext()) { Logger.log('❌ ไม่พบ root folder'); return; }
  var root = roots.next();
  var ids = {};

  // MEA
  var meaFolders = root.getFoldersByName('⚡ MEA — การไฟฟ้านครหลวง (012076042)');
  if (meaFolders.hasNext()) {
    scanFileIds_(meaFolders.next(), 'MEA', ids);
  }

  // MWA
  var mwaFolders = root.getFoldersByName('💧 MWA — การประปานครหลวง (74475021)');
  if (mwaFolders.hasNext()) {
    scanFileIds_(mwaFolders.next(), 'MWA', ids);
  }

  Logger.log('══════════════════════════════════════');
  Logger.log(' Drive File IDs');
  Logger.log(' วาง JSON นี้ใน localStorage key: jres_utility_fids_v1');
  Logger.log('══════════════════════════════════════');
  Logger.log(JSON.stringify(ids, null, 2));
  return ids;
}

function scanFileIds_(utilityFolder, prefix, ids) {
  var yearFolders = utilityFolder.getFolders();
  while (yearFolders.hasNext()) {
    var yearFolder = yearFolders.next();
    var yearName = yearFolder.getName().replace(/[^\d]/g, '');
    var monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      var mf = monthFolders.next();
      var mfName = mf.getName();
      var mm = mfName.match(/^(\d{2})/);
      if (!mm) continue;
      var files = mf.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        var fn = f.getName().toLowerCase();
        var key = prefix + '_' + yearName + '_' + mm[1];
        if (fn.indexOf('invoice') !== -1) {
          ids[key + '_inv'] = f.getId();
        } else if (fn.indexOf('receipt') !== -1) {
          ids[key + '_rec'] = f.getId();
        }
      }
    }
  }
}
