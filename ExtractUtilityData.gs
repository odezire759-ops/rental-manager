/**
 * ================================================================
 *  J.Residence — Extract MEA/MWA Data from PDFs in Google Drive
 *  อ่านข้อมูลจาก PDF ใบแจ้งหนี้/ใบเสร็จ แล้ว output เป็น JSON
 *  สำหรับ import เข้า Dashboard (sec-utility)
 * ================================================================
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
 * ================================================================
 */

// ── Root folder name (ต้องตรงกับ ROOT_NAME ใน runAll script) ──
var EXTRACT_ROOT_NAME = '📁 J.Residence — สาธารณูปโภค';
var MEA_FOLDER_NAME   = '⚡ MEA — การไฟฟ้านครหลวง (012076042)';
var MWA_FOLDER_NAME   = '💧 MWA — การประปานครหลวง (74475021)';

// Cache expiry: 30 minutes (ไม่ต้อง OCR ใหม่ทุกครั้ง)
var CACHE_KEY = 'jres_utility_extract_v1';
var CACHE_TTL = 1800; // seconds

// ==============================================================
//  DOWNLOAD FROM GMAIL → DRIVE
//  แก้ปัญหา Gmail threading: วน ALL messages ใน ALL threads
//  รัน downloadFromGmail() เพื่อดึง PDF ที่ขาดหายจาก Gmail
// ==============================================================
function downloadFromGmail() {
  Logger.log('======================================');
  Logger.log(' Download MEA/MWA PDFs: Gmail → Drive');
  Logger.log(' (iterates ALL messages in ALL threads)');
  Logger.log('======================================\n');

  var rootFolder = getOrCreateFolder_(null, EXTRACT_ROOT_NAME);
  var meaRoot    = getOrCreateFolder_(rootFolder, MEA_FOLDER_NAME);
  var mwaRoot    = getOrCreateFolder_(rootFolder, MWA_FOLDER_NAME);

  var stats = { saved: 0, skipped: 0, noDate: 0, err: 0 };

  // ── Search configs ──────────────────────────────────────────
  // จำกัดปี 2567-2569 (ค.ศ. 2024-01-01 ถึง 2027-01-01) เพื่อหลีกเลี่ยง timeout
  var searches = [
    { query: '012076042 has:attachment after:2023/12/31 before:2027/01/01', type: 'mea', root: meaRoot, label: 'MEA' },
    { query: '74475021 has:attachment after:2023/12/31 before:2027/01/01',  type: 'mwa', root: mwaRoot, label: 'MWA' }
  ];

  var startTime = Date.now();
  var MAX_MS = 5 * 60 * 1000; // หยุดก่อน 5 นาที (limit 6 นาที)

  searches.forEach(function(cfg) {
    Logger.log('🔍 Searching Gmail: "' + cfg.query + '"');
    var threads = GmailApp.search(cfg.query, 0, 500);
    Logger.log('   พบ ' + threads.length + ' threads\n');

    threads.forEach(function(thread) {
      // หยุดถ้าใกล้ timeout
      if (Date.now() - startTime > MAX_MS) {
        Logger.log('  ⏱ หยุดก่อน timeout — รัน downloadFromGmail() อีกครั้งเพื่อต่อ');
        return;
      }
      // ── วน EVERY message ใน thread (แก้ปัญหา threading) ──
      var messages = thread.getMessages();
      messages.forEach(function(msg) {
        var subject = msg.getSubject();
        var msgDate = msg.getDate();

        msg.getAttachments().forEach(function(att) {
          // PDF เท่านั้น
          if (att.getContentType().indexOf('pdf') === -1 &&
              !att.getName().match(/\.pdf$/i)) return;

          // Detect month/year จาก subject + filename + email date
          var info = detectMonthFromGmail_(subject, att.getName(), msgDate);
          if (!info) {
            Logger.log('  ⚠ ไม่พบเดือน: ' + att.getName().substring(0,40)
                       + ' | subject: ' + subject.substring(0,50));
            stats.noDate++;
            return;
          }

          // ตัดสินใจ Invoice vs Receipt
          // MWA = Receipt เสมอ (subject มีคำ "Invoice" แต่จริงๆ คือ e-Receipt)
          var fileType = (cfg.type === 'mwa') ? 'Receipt' : detectFileType_(subject, att.getName());

          // สร้างชื่อไฟล์
          var mo2 = info.month < 10 ? '0'+info.month : String(info.month);
          var fileName = cfg.label + '_' + fileType + '_'
                       + info.thShort.replace(/\./g,'') + '_' + info.year
                       + '_' + mo2 + '.pdf';

          // หาหรือสร้าง folder: root/year/MM_thShort
          var yearFolder  = getOrCreateFolder_(cfg.root, String(info.year));
          var moFolderName = mo2 + '_' + info.thShort;
          var monthFolder  = getOrCreateFolder_(yearFolder, moFolderName);

          // Skip ถ้ามีไฟล์ชื่อเดียวกันอยู่แล้ว
          if (monthFolder.getFilesByName(fileName).hasNext()) {
            stats.skipped++;
            return;
          }

          // บันทึกลง Drive
          try {
            monthFolder.createFile(att.copyBlob().setName(fileName));
            Logger.log('  ✅ [' + info.thShort + ' ' + info.year + '] ' + fileName);
            stats.saved++;
          } catch(e) {
            Logger.log('  ❌ ' + fileName + ': ' + e.message);
            stats.err++;
          }
        });
      });
    });
    Logger.log('');
  });

  Logger.log('======================================');
  Logger.log(' สรุป: บันทึก ' + stats.saved
    + ' | ข้าม (มีอยู่แล้ว) ' + stats.skipped
    + ' | ไม่พบเดือน ' + stats.noDate
    + ' | ผิดพลาด ' + stats.err);
  Logger.log('======================================');
  Logger.log('\n▶ รัน extractAll() หรือ clearCache() + doGet เพื่ออัปเดต Dashboard');
}

// ── ตรวจจับเดือน/ปี จาก email subject + attachment filename + email date ──
function detectMonthFromGmail_(subject, filename, msgDate) {
  var TH_SHORT = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                      'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

  // Pattern 1: MWA subject "(วันที่ 15/10/2568)"
  var m = subject.match(/วันที่\s+(\d{1,2})\/(\d{1,2})\/(25\d{2})/);
  if (m) {
    var mo = parseInt(m[2]), yr = parseInt(m[3]);
    if (mo >= 1 && mo <= 12) return { month: mo, year: yr, thShort: TH_SHORT[mo] };
  }

  // Pattern 2: attachment filename "25681015_xxx.pdf" → YYYYMMDD
  var fn = filename.match(/^(25\d{2})(\d{2})(\d{2})/);
  if (fn) {
    var yr = parseInt(fn[1]), mo = parseInt(fn[2]);
    if (mo >= 1 && mo <= 12) return { month: mo, year: yr, thShort: TH_SHORT[mo] };
  }

  // Pattern 3: any DD/MM/YYYY (Buddhist) in subject
  var m2 = subject.match(/(\d{1,2})\/(\d{1,2})\/(25\d{2})/);
  if (m2) {
    var mo = parseInt(m2[2]), yr = parseInt(m2[3]);
    if (mo >= 1 && mo <= 12) return { month: mo, year: yr, thShort: TH_SHORT[mo] };
  }

  // Pattern 4: MEA subject "ประจำเดือน MM/YY" or "MM/2568"
  var m3 = subject.match(/(\d{1,2})\/(25\d{2})/);
  if (m3) {
    var mo = parseInt(m3[1]), yr = parseInt(m3[2]);
    if (mo >= 1 && mo <= 12) return { month: mo, year: yr, thShort: TH_SHORT[mo] };
  }

  // Fallback: ใช้ email date (แปลง ค.ศ. → พ.ศ.)
  if (msgDate) {
    var mo = msgDate.getMonth() + 1;
    var yr = msgDate.getFullYear() + 543;
    return { month: mo, year: yr, thShort: TH_SHORT[mo] };
  }

  return null;
}

// ── Invoice vs Receipt ──────────────────────────────────────────
function detectFileType_(subject, filename) {
  var txt = (subject + ' ' + filename).toLowerCase();
  if (/invoice|ใบแจ้งหนี้|tax\s*invoice/.test(txt)) return 'Invoice';
  return 'Receipt'; // MWA e-Tax = Receipt by default
}

// ── Get or create Drive folder ──────────────────────────────────
function getOrCreateFolder_(parent, name) {
  var iter = parent ? parent.getFoldersByName(name)
                    : DriveApp.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

// ==============================================================
//  WEB APP ENDPOINT — อ่านจาก Drive JSON file (เร็ว, ไม่ run OCR)
//  OCR ทำแยกต่างหากด้วย extractAndSave() แล้วเก็บไว้เป็นไฟล์
// ==============================================================
var DATA_FILE_NAME = 'jres_utility_data.json';

function doGet(e) {
  // 1. ลอง CacheService ก่อน (30 นาที)
  var data = cacheLoad_();

  // 2. ถ้าไม่มี cache → อ่าน Drive JSON file
  if (!data) {
    data = loadFromDriveFile_();
    if (data) cacheSave_(data); // warm up cache
  }

  // 3. ถ้ายังไม่มีเลย → แจ้งให้รัน extractAndSave()
  if (!data) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: 'ยังไม่มีข้อมูล — รัน extractAndSave() ใน Apps Script ก่อน',
        data: {}
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      ts: new Date().toISOString(),
      count: Object.keys(data).length,
      data: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
//  EXTRACT + SAVE — รัน manually หลังได้รับ PDF ใหม่
//  OCR ทุกไฟล์แล้วบันทึกเป็น JSON ไว้ใน Drive
//  (อาจใช้เวลานาน ถ้า timeout ให้รันซ้ำ — ผลลัพธ์สะสม)
// ==============================================================
function extractAndSave() {
  Logger.log('== extractAndSave: start ==');

  // โหลดข้อมูลเก่าจาก Drive ไว้ก่อน (เพื่อ merge ไม่ใช่ overwrite)
  var existing = loadFromDriveFile_() || {};
  var startTime = Date.now();
  var MAX_MS = 5 * 60 * 1000; // หยุดก่อน timeout ที่ 5 นาที

  var stats = { total: 0, success: 0, fail: 0, skipped: 0, timedOut: false };

  var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
  if (!roots.hasNext()) {
    Logger.log('ไม่พบ folder: ' + EXTRACT_ROOT_NAME);
    return;
  }
  var rootFolder = roots.next();

  var configs = [
    { folderName: MEA_FOLDER_NAME, type: 'mea', label: 'MEA' }
    // MWA: ปิดชั่วคราว รัน extractAndSaveMWA() แยกต่างหาก
  ];

  configs.forEach(function(cfg) {
    var folders = rootFolder.getFoldersByName(cfg.folderName);
    if (!folders.hasNext()) return;
    var utilRoot = folders.next();
    Logger.log('\nProcessing ' + cfg.label + '...');

    var yearFolders = utilRoot.getFolders();
    while (yearFolders.hasNext()) {
      var yearFolder = yearFolders.next();
      var monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        var monthFolder = monthFolders.next();
        var files = monthFolder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          if (file.getMimeType() !== 'application/pdf') continue;

          // หยุดถ้าใกล้ timeout
          if (Date.now() - startTime > MAX_MS) {
            Logger.log('\nหยุดก่อน timeout — บันทึกผลที่ได้แล้ว รันซ้ำเพื่อ process ไฟล์ที่เหลือ');
            stats.timedOut = true;
            break;
          }

          // Skip ถ้าเดือนนี้มี amount อยู่แล้ว (ไม่ต้อง OCR ซ้ำ)
          var quickKey = detectMonthKey_(file.getName(), '', cfg.type);
          if (quickKey && existing[quickKey] && existing[quickKey].amount) {
            stats.skipped++;
            continue;
          }

          stats.total++;
          try {
            var parsed = extractPdfData_(file, cfg.type);
            if (parsed && parsed.key) {
              if (parsed.values.offpeak && parsed.values.onpeak &&
                  parsed.values.offpeak === parsed.values.onpeak) {
                delete parsed.values.offpeak;
                delete parsed.values.onpeak;
                delete parsed.values.totalUnits;
              }
              var rec = existing[parsed.key] || {};
              for (var k in parsed.values) {
                if (parsed.values[k]) rec[k] = parsed.values[k];
              }
              existing[parsed.key] = rec;
              stats.success++;
              Logger.log('  OK ' + file.getName());
            } else {
              stats.fail++;
              Logger.log('  SKIP (parse fail): ' + file.getName());
            }
          } catch(e) {
            stats.fail++;
            Logger.log('  ERR ' + file.getName() + ': ' + e.message);
          }
        }
        if (stats.timedOut) break;
      }
      if (stats.timedOut) break;
    }
  });

  // บันทึกลง Drive file + warm cache
  saveToDriveFile_(existing);
  cacheSave_(existing);

  Logger.log('\n== extractAndSave: done ==');
  Logger.log('success=' + stats.success + ' fail=' + stats.fail
    + ' total=' + stats.total
    + (stats.timedOut ? ' [PARTIAL — run again]' : ' [COMPLETE]'));
}

// ==============================================================
//  EXTRACT MWA ONLY (รันแยก หลังจาก extractAndSave() MEA เสร็จ)
// ==============================================================
function extractAndSaveMWA() {
  Logger.log('== extractAndSaveMWA: start ==');
  var existing = loadFromDriveFile_() || {};
  var startTime = Date.now();
  var MAX_MS = 5 * 60 * 1000;
  var stats = { total: 0, success: 0, fail: 0, timedOut: false };

  var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
  if (!roots.hasNext()) { Logger.log('ไม่พบ root folder'); return; }
  var rootFolder = roots.next();

  var mwaFolders = rootFolder.getFoldersByName(MWA_FOLDER_NAME);
  if (!mwaFolders.hasNext()) { Logger.log('ไม่พบ MWA folder'); return; }
  var mwaRoot = mwaFolders.next();

  var yearFolders = mwaRoot.getFolders();
  while (yearFolders.hasNext()) {
    var yearFolder = yearFolders.next();
    var monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      var monthFolder = monthFolders.next();
      var files = monthFolder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        if (file.getMimeType() !== 'application/pdf') continue;
        if (Date.now() - startTime > MAX_MS) {
          Logger.log('หยุดก่อน timeout — รันซ้ำ');
          stats.timedOut = true;
          break;
        }
        var qk = detectMonthKey_(file.getName(), '', 'mwa');
        if (qk && existing[qk] && existing[qk].amount) { stats.skipped++; continue; }
        stats.total++;
        try {
          var parsed = extractPdfData_(file, 'mwa');
          if (parsed && parsed.key) {
            var rec = existing[parsed.key] || {};
            for (var k in parsed.values) {
              if (parsed.values[k]) rec[k] = parsed.values[k];
            }
            existing[parsed.key] = rec;
            stats.success++;
            Logger.log('  OK ' + file.getName());
          } else {
            stats.fail++;
            Logger.log('  SKIP: ' + file.getName());
          }
        } catch(e) {
          stats.fail++;
          Logger.log('  ERR ' + file.getName() + ': ' + e.message);
        }
      }
      if (stats.timedOut) break;
    }
    if (stats.timedOut) break;
  }

  saveToDriveFile_(existing);
  cacheSave_(existing);
  Logger.log('== extractAndSaveMWA: done == success=' + stats.success
    + ' fail=' + stats.fail + (stats.timedOut ? ' [PARTIAL]' : ' [COMPLETE]'));
}

// ==============================================================
//  Drive JSON file helpers
// ==============================================================
function loadFromDriveFile_() {
  try {
    var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
    if (!roots.hasNext()) return null;
    var folder = roots.next();
    var files = folder.getFilesByName(DATA_FILE_NAME);
    if (!files.hasNext()) return null;
    var content = files.next().getBlob().getDataAsString();
    return JSON.parse(content);
  } catch(ex) {
    return null;
  }
}

function saveToDriveFile_(data) {
  try {
    var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
    if (!roots.hasNext()) return;
    var folder = roots.next();
    var json = JSON.stringify(data, null, 2);
    // ลบไฟล์เก่าแล้วสร้างใหม่
    var existing = folder.getFilesByName(DATA_FILE_NAME);
    while (existing.hasNext()) existing.next().setTrashed(true);
    folder.createFile(DATA_FILE_NAME, json, 'application/json');
    Logger.log('Saved ' + DATA_FILE_NAME + ' (' + json.length + ' bytes)');
  } catch(ex) {
    Logger.log('saveToDriveFile_ failed: ' + ex.message);
  }
}

// ── CacheService helpers (L1 cache, 30 min) ──────────────────
var CACHE_CHUNK = 90000;
var CACHE_META  = CACHE_KEY + '_meta';

function cacheSave_(data) {
  try {
    var json = JSON.stringify(data);
    var cache = CacheService.getScriptCache();
    if (json.length <= CACHE_CHUNK) {
      cache.put(CACHE_KEY, json, CACHE_TTL);
      cache.remove(CACHE_META);
    } else {
      var chunks = [];
      for (var i = 0; i < json.length; i += CACHE_CHUNK) {
        chunks.push(json.substring(i, i + CACHE_CHUNK));
      }
      var puts = {};
      chunks.forEach(function(c, idx) { puts[CACHE_KEY + '_' + idx] = c; });
      puts[CACHE_META] = String(chunks.length);
      cache.putAll(puts, CACHE_TTL);
      cache.remove(CACHE_KEY);
    }
  } catch(ex) {}
}

function cacheLoad_() {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get(CACHE_META);
    if (meta) {
      var n = parseInt(meta);
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
      var vals = cache.getAll(keys);
      return JSON.parse(keys.map(function(k) { return vals[k]||''; }).join(''));
    }
    var single = cache.get(CACHE_KEY);
    return single ? JSON.parse(single) : null;
  } catch(ex) { return null; }
}

// ==============================================================
//  Clear cache + Drive file (force re-extract next time)
// ==============================================================
function clearCache() {
  var cache = CacheService.getScriptCache();
  var meta = cache.get(CACHE_META);
  if (meta) {
    var n = parseInt(meta);
    var keys = [CACHE_META];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    cache.removeAll(keys);
  }
  cache.remove(CACHE_KEY);
  Logger.log('Cache cleared');
}

function clearAll() {
  clearCache();
  try {
    var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
    if (roots.hasNext()) {
      var f = roots.next().getFilesByName(DATA_FILE_NAME);
      while (f.hasNext()) f.next().setTrashed(true);
      Logger.log('Drive file deleted');
    }
  } catch(ex) {}
  Logger.log('clearAll done — run extractAndSave() to rebuild');
}

// ==============================================================
//  MAIN — Extract all PDF data and output JSON
// ==============================================================
function extractAll() {
  Logger.log('======================================');
  Logger.log(' Extract MEA/MWA Data from PDFs');
  Logger.log('======================================\n');

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
  Logger.log('\n======================================');
  Logger.log(' สรุป: สำเร็จ ' + stats.success + '/' + stats.total + ' ไฟล์');
  Logger.log('======================================');

  var json = JSON.stringify(result, null, 2);
  Logger.log('\n======================================');
  Logger.log(' JSON สำหรับ Import (copy ทั้งหมดด้านล่าง)');
  Logger.log('======================================');
  Logger.log(json);

  return result;
}

// ==============================================================
//  Process utility folder tree
// ==============================================================
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
            // Validate: reject if offpeak === onpeak (OCR bug — likely total/2)
            if (data.values.offpeak && data.values.onpeak && data.values.offpeak === data.values.onpeak) {
              Logger.log('    ⚠ offpeak === onpeak (' + data.values.offpeak + ') — rejected (OCR bug)');
              delete data.values.offpeak;
              delete data.values.onpeak;
              delete data.values.totalUnits;
            }

            // Merge with existing — OVERWRITE with new values
            var existing = result[data.key] || {};
            for (var k in data.values) {
              if (data.values[k]) {
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

// ==============================================================
//  Extract text from PDF via Drive OCR conversion
// ==============================================================
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

// ==============================================================
//  Extract data from PDF file
// ==============================================================
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

// ==============================================================
//  Parse MEA PDF text
// ==============================================================
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
  // MEA OCR: column headers appear on one line, values on next line
  // "Meter Reading Date Last Meter Reading ... ตัวคูณ\n20016690834 28/02/69 603749 ..."
  // → ต้องใช้ [\s\S] เพื่อข้ามบรรทัด + invoice number (digits) ได้
  var mDatePatterns = [
    /Meter\s*Reading\s*Date[\s\S]{1,600}?(\d{1,2}\/\d{1,2}\/\d{2})\b/i,
    /วันที่อ่านมิเตอร์[\s\S]{1,300}?(\d{1,2}\/\d{1,2}\/\d{2,4})/i,
    /วัน(?:ที่)?จดมิเตอร์[\s\S]{1,200}?(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i,
    /\b2\d{9,11}\s+(\d{1,2}\/\d{1,2}\/\d{2,4})/
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

// ==============================================================
//  Parse MWA PDF text
// ==============================================================
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

// ==============================================================
//  Detect month key from filename / text
// ==============================================================
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

// ==============================================================
//  OPTIONAL: Extract to Google Sheet
// ==============================================================
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

// ==============================================================
//  TEST: ทดสอบ OCR + parse กับ PDF ตัวแรกที่เจอ
//  รัน function นี้แล้วดู Execution Log
// ==============================================================
function testExtractSingle() {
  // หาไฟล์ PDF จาก Drive อัตโนมัติ (ไม่ต้องวาง File ID)
  var roots = DriveApp.getFoldersByName(EXTRACT_ROOT_NAME);
  if (!roots.hasNext()) { Logger.log('❌ ไม่พบ root folder'); return; }
  var root = roots.next();

  // หา MEA Invoice ตัวแรก
  var file = null;
  var meaFolders = root.getFoldersByName('⚡ MEA — การไฟฟ้านครหลวง (012076042)');
  if (meaFolders.hasNext()) {
    var meaRoot = meaFolders.next();
    var yrs = meaRoot.getFolders();
    while (yrs.hasNext() && !file) {
      var yr = yrs.next();
      var mos = yr.getFolders();
      while (mos.hasNext() && !file) {
        var mo = mos.next();
        var fs = mo.getFiles();
        while (fs.hasNext()) {
          var f = fs.next();
          if (f.getName().toLowerCase().indexOf('invoice') !== -1 && f.getMimeType() === 'application/pdf') {
            file = f;
            break;
          }
        }
      }
    }
  }

  if (!file) { Logger.log('❌ ไม่พบ MEA Invoice PDF'); return; }

  Logger.log('📄 File: ' + file.getName());
  Logger.log('   ID: ' + file.getId());
  Logger.log('   Size: ' + file.getSize() + ' bytes\n');

  var text = extractPdfText_(file);

  Logger.log('=======================================');
  Logger.log('  RAW OCR TEXT (' + text.length + ' chars)');
  Logger.log('=======================================');
  Logger.log(text);

  Logger.log('\n=======================================');
  Logger.log('  PARSED MEA RESULT');
  Logger.log('=======================================');
  var parsed = parseMeaText_(text);
  Logger.log(JSON.stringify(parsed, null, 2));

  Logger.log('\n=======================================');
  Logger.log('  MONTH DETECTION');
  Logger.log('=======================================');
  var monthKey = detectMonthKey_(file.getName(), text, 'mea');
  Logger.log('monthKey: ' + (monthKey || 'NULL — ไม่สามารถระบุเดือน'));
}

// ==============================================================
//  OPTIONAL: Get all Drive File IDs for index.html
//  (อัปเดต _UT_DRIVE_FILE_IDS ใน dashboard)
// ==============================================================
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

  Logger.log('======================================');
  Logger.log(' Drive File IDs');
  Logger.log(' วาง JSON นี้ใน localStorage key: jres_utility_fids_v1');
  Logger.log('======================================');
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
