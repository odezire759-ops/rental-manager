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
 *  3. Deploy → Manage deployments → New → Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Copy Web App URL → วางใน Dashboard (ระบบ sync อัตโนมัติ)
 *
 *  หรือรัน "extractAll" ด้วยมือเพื่อดู JSON ใน Log
 *  หรือรัน "extractToSheet" เพื่อ export เป็น Google Sheet
 * ════════════════════════════════════════════════════════════════
 */

// ── Root folder name (ต้องตรงกับ ROOT_NAME ใน runAll script) ──
var EXTRACT_ROOT_NAME = '📁 J.Residence — สาธารณูปโภค';

// Cache expiry: 30 minutes (ไม่ต้อง OCR ใหม่ทุกครั้ง)
var CACHE_KEY = 'jres_utility_extract_v1';
var CACHE_TTL = 1800; // seconds

// ══════════════════════════════════════════════════════════════
//  WEB APP ENDPOINT — Dashboard เรียกผ่าน fetch()
//  Deploy → Web app → Execute as: Me → Access: Anyone
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  var forceRefresh = e && e.parameter && e.parameter.refresh === '1';
  var data;

  if (!forceRefresh) {
    // Try cache first
    var cache = CacheService.getScriptCache();
    var cached = cache.get(CACHE_KEY);
    if (cached) {
      try {
        data = JSON.parse(cached);
      } catch(ex) {
        data = null;
      }
    }
  }

  if (!data) {
    // Extract fresh data from PDFs
    data = extractAll();
    // Save to cache (split if > 100KB — CacheService limit)
    try {
      var json = JSON.stringify(data);
      if (json.length < 100000) {
        CacheService.getScriptCache().put(CACHE_KEY, json, CACHE_TTL);
      }
    } catch(ex) {
      Logger.log('Cache save failed: ' + ex.message);
    }
  }

  // Return JSON with CORS headers
  var output = ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      ts: new Date().toISOString(),
      count: data ? Object.keys(data).length : 0,
      data: data || {}
    }))
    .setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ══════════════════════════════════════════════════════════════
//  Clear cache (run manually after adding new PDFs)
// ══════════════════════════════════════════════════════════════
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  Logger.log('✅ Cache cleared — next doGet() will re-extract from PDFs');
}

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
  if (!text || text.length < 50) {
    Logger.log('    ⚠ OCR text too short (' + (text ? text.length : 0) + ' chars): ' + fileName);
    return null;
  }

  var isReceipt = /Receipt|ใบเสร็จ/i.test(fileName) ||
                  /ใบเสร็จรับเงิน|RECEIPT/i.test(text);

  // Detect month from filename (e.g., MEA_Invoice_ม.ค._2569_...)
  var monthKey = detectMonthKey_(fileName, text, type);
  if (!monthKey) {
    Logger.log('    ⚠ ไม่สามารถระบุเดือน: ' + fileName);
    Logger.log('      OCR text (first 300 chars): ' + text.substring(0, 300));
    return null;
  }

  var values = {};

  if (type === 'mea') {
    values = parseMeaText_(text);
  } else {
    values = parseMwaText_(text);
  }

  // Log extracted values for verification
  var fields = Object.keys(values).filter(function(k) { return values[k]; });
  Logger.log('    → ' + monthKey + ': ' + fields.join(', '));

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
  // MEA PDF: "รวมเงินที่ต้องชำระทั้งสิ้น (Amount) 40,520.97 บาท"
  // หรือ "รวมค่าไฟฟ้าเดือนปัจจุบัน 40,520.97 บาท"
  var amtPatterns = [
    /(?:รวมเงินที่ต้องชำระทั้งสิ้น|Amount)[^\d]{0,20}([\d,]+\.\d{2})/i,
    /(?:รวมค่าไฟฟ้าเดือนปัจจุบัน)[^\d]{0,20}([\d,]+\.\d{2})/i,
    /(?:จำนวนเงินรวม|ยอดเงินรวม|ยอดรวม(?:ทั้งสิ้น)?)[^\d]{0,30}([\d,]+\.\d{2})/i,
    /(?:TOTAL\s*AMOUNT|NET\s*AMOUNT|AMOUNT\s*DUE)[^\d]{0,20}([\d,]+\.\d{2})/i
  ];
  for (var i = 0; i < amtPatterns.length; i++) {
    var m = text.match(amtPatterns[i]);
    if (m) {
      var amt = m[1].replace(/,/g, '');
      if (parseFloat(amt) > 1000) {
        data.amount = amt;
        break;
      }
    }
  }

  // ── On Peak / Off Peak units ──
  // MEA PDF: "จำนวน On Peak 2,284 หน่วย" + "จำนวน Off Peak 7,120 หน่วย"
  // ต้อง match "หน่วย" เพื่อไม่จับ demand (กิโลวัตต์/กิโลวาร์)
  // Pattern 1: "จำนวน On/Off Peak XXX หน่วย" (from the right side box)
  var onM = text.match(/จำนวน\s*On\s*Peak\s*([\d,]+)\s*หน่วย/i);
  if (!onM) onM = text.match(/On\s*Peak\s*([\d,]+)\s*หน่วย/i);
  if (onM) data.onpeak = onM[1].replace(/,/g, '');

  var offM = text.match(/จำนวน\s*Off\s*Peak\s*([\d,]+)\s*หน่วย/i);
  if (!offM) offM = text.match(/Off\s*Peak\s*([\d,]+)\s*หน่วย/i);
  if (offM) data.offpeak = offM[1].replace(/,/g, '');

  // Fallback: "On Peak 2,284 หน่วย 9,889.03 บาท" (from detail table)
  if (!data.onpeak) {
    var onF = text.match(/On\s*Peak\s*([\d,]+)\s*หน่วย\s*([\d,]+\.\d{2})\s*บาท/i);
    if (onF) data.onpeak = onF[1].replace(/,/g, '');
  }
  if (!data.offpeak) {
    var offF = text.match(/Off\s*Peak\s*([\d,]+)\s*หน่วย\s*([\d,]+\.\d{2})\s*บาท/i);
    if (offF) data.offpeak = offF[1].replace(/,/g, '');
  }

  // ── Total units (auto-calc) ──
  if (data.offpeak && data.onpeak) {
    data.totalUnits = String(parseFloat(data.offpeak) + parseFloat(data.onpeak));
  }

  // ── จำนวนหน่วย (kWh) — from header table ──
  // MEA PDF header: "จำนวนหน่วย kWh 9,404"
  if (!data.totalUnits) {
    var kwhM = text.match(/(?:จำนวนหน่วย|kWh)[^\d]{0,10}([\d,]+)/i);
    if (kwhM) data.totalUnits = kwhM[1].replace(/,/g, '');
  }

  // ── วันที่จดมิเตอร์ (Meter Reading Date) ──
  // MEA PDF: "วันที่อ่านมิเตอร์ Meter Reading Date 31/01/69"
  // หรือ header table: "วันที่อ่านมิเตอร์ล่าสุด Meter Reading Date 31/01/69"
  var mDatePatterns = [
    /Meter\s*Reading\s*Date[^\d]{0,15}(\d{1,2}\/\d{1,2}\/\d{2,4})/i,
    /วันที่อ่านมิเตอร์[^\d]{0,30}(\d{1,2}\/\d{1,2}\/\d{2,4})/i,
    /วัน(?:ที่)?จดมิเตอร์[^\d]{0,20}(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i
  ];
  for (var i = 0; i < mDatePatterns.length; i++) {
    var m = text.match(mDatePatterns[i]);
    if (m) { data.meterDate = m[1]; break; }
  }

  // ── รอบบิล (Billing period) ──
  // MEA PDF: "บิลประจำเดือน 01/69"
  var billMonthM = text.match(/บิล(?:ประจำ)?เดือน[^\d]{0,10}(\d{1,2}\/\d{2,4})/i);
  if (billMonthM) data.period = billMonthM[1];
  // Or full period range
  if (!data.period) {
    var periodM = text.match(/(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})\s*[-–ถึง\s]+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/);
    if (periodM) data.period = periodM[1] + '–' + periodM[2];
  }

  // ── วันครบกำหนดชำระ (Payment Due Date) ──
  // MEA PDF: "โปรดชำระภายในวันที่ Payment Due Date 16/02/69"
  var dueM = text.match(/(?:Payment\s*Due\s*Date|โปรดชำระภายใน(?:วันที่)?)[^\d]{0,20}(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
  if (dueM) data.paidDate = dueM[1];

  // ── Demand (kW) ──
  var demandM = text.match(/(?:DEMAND|ดีมานด์|กำลังไฟสูงสุด)[^\d]{0,15}([\d,]+(?:\.\d+)?)\s*(?:kW)?/i);
  if (demandM) data.demand = demandM[1].replace(/,/g, '');

  // ── Power Factor ──
  var pfM = text.match(/(?:POWER\s*FACTOR|ตัวประกอบกำลัง|PF|Pf)[^\d]{0,15}([\d.]+)/i);
  if (pfM) data.powerFactor = pfM[1];

  // ── Multiplier ──
  var multM = text.match(/(?:Multiplier|ตัวคูณ)[^\d]{0,10}(\d+)/i);
  if (multM) data.multiplier = multM[1];

  // ── เลขมิเตอร์ ──
  // MEA PDF header: "เลขอ่านก่อน Previous Meter Reading 594345"
  var prevM = text.match(/(?:Previous\s*Meter\s*Reading|เลขอ่านก่อน)[^\d]{0,15}([\d,]+)/i);
  if (prevM) data.meterOld = prevM[1].replace(/,/g, '');
  var lastM = text.match(/(?:Last\s*Meter\s*Reading|เลขอ่านล่าสุด|เลขอ่านครั้งหลัง)[^\d]{0,15}([\d,]+)/i);
  if (lastM) data.meterNew = lastM[1].replace(/,/g, '');

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
  // ใช้ [\s_]* แทน \s* เพราะ filename ใช้ _ เป็น separator
  for (var m = 1; m <= 12; m++) {
    var short = TH_SHORT[m];
    var esc = short.replace(/\./g, '\\.');
    var re = new RegExp(esc + '[\\s_]*(25\\d{2})');
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

  // Try from text: Thai full month names + short month names
  // ใช้ [\s_.]* เพื่อรองรับ OCR ที่อาจมีช่องว่าง/จุด/underscore แทรก
  var TH_FULL = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  for (var m = 1; m <= 12; m++) {
    var reF = new RegExp(TH_FULL[m] + '[\\s_.]{0,5}(25\\d{2})');
    var reS = new RegExp(TH_SHORT[m].replace(/\./g, '\\.') + '[\\s_.]{0,5}(25\\d{2})');
    var mF = text.match(reF);
    var mS = text.match(reS);
    var found = mF || mS;
    if (found) {
      return type + '_' + TH_SHORT[m] + '_' + found[1];
    }
  }

  // Fallback: try DD/MM/YYYY patterns in text
  var dateMatch = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](25\d{2})/);
  if (dateMatch) {
    var dm = parseInt(dateMatch[2]); // month from date
    if (dm >= 1 && dm <= 12) {
      return type + '_' + TH_SHORT[dm] + '_' + dateMatch[3];
    }
  }
  // Try Gregorian year → convert to Buddhist
  var gregMatch = text.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})/);
  if (gregMatch) {
    var gm = parseInt(gregMatch[2]);
    var gy = parseInt(gregMatch[3]) + 543;
    if (gm >= 1 && gm <= 12) {
      return type + '_' + TH_SHORT[gm] + '_' + gy;
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
