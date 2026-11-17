/******************** CONFIG ********************/
const SPREADSHEET_ID = '1yF2jIPcKXgGLDQFcsDABY-zQeoj87XquGaENwTgN8xQ';
const ROOT_FOLDER_NAME = 'AnimalSurvey_OwnerPhotos';
const TIME_ZONE = 'Asia/Bangkok';

// ปีของโฟลเดอร์รูปภาพใน Drive (บังคับให้เป็น 2569 เสมอ)
const PHOTO_FOLDER_YEAR = '2569';

// รูปแบบการแสดง timestamp ในชีต (วันที่ + เวลา)
const TIMESTAMP_FORMAT = 'dd/MM/yyyy, HH:mm:ss';

/******************** MASTER HEADER ********************/
const MASTER_HEADERS = [
  "timestamp","ปีงบประมาณ","ศูนย์บริการ","ชุมชน",
  "ประเภทสัตว์","ชนิดสัตว์ (อื่น)","จำนวนสัตว์ (อื่น)",
  "ลำดับสัตว์","ชื่อสัตว์","เพศ","อายุ (ปี)","อายุ (เดือน)",
  "สี / ตำหนิ","สถานะทำหมัน","วันที่ฉีดยาคุม","สัตวแพทย์ผู้ฉีดยาคุม",
  "สถานะวัคซีนพิษสุนัขบ้า","วันที่ฉีดวัคซีน","สัตวแพทย์ผู้ฉีดวัคซีน",
  "ลักษณะการเลี้ยง","สถานที่เลี้ยง","พื้นที่การเลี้ยง",
  "ชื่อเจ้าของสัตว์","เลขบัตรประชาชน","เบอร์โทรศัพท์มือถือ",
  "เบอร์โทรศัพท์บ้าน","บ้านเลขที่","ถนน","ซอย",
  "ตำบล","อำเภอ","จังหวัด",
  "พิกัดที่เลือก",
  "ผู้บันทึก","ตำแหน่ง","เลขบัตรผู้บันทึก",
  "ลิงก์รูปเจ้าของ","ลิงก์รูปสัตว์"
];

/******************** MAIN ********************/
function doPost(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    return output_('ระบบกำลังประมวลผล กรุณาลองใหม่');
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return output_('No POST data');
    }

    let data = {};

    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return output_('JSON ไม่ถูกต้อง');
    }

    // คำสั่งลบข้อมูลบ้าน (จากปุ่มลบหมุดในฟอร์ม)
    if (data && data.action === 'deleteHouse') {
      return deleteHouseData_(data);
    }

    if (data && data.action === 'getHouse') {
      return getHouseData_(data);
    }

    if (data && data.action === 'getSurveyorStats') {
      return getSurveyorStats_(data);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date();

    const year = safe_(data.year) || '2569';
    const centerName = safe_(data.centerName || data.center);
    const communityName = safe_(data.communityName || data.community);

    if (!centerName || !communityName) {
      return output_('กรุณาระบุชื่อศูนย์บริการและชุมชน');
    }

    const addr = data.address || {};

    // โหมดแก้ไข: ลบแถวเดิมของบ้านหลังนี้ก่อน แล้วบันทึกรอบใหม่ทับ (กันแถวซ้ำ)
    const isUpdate = data && data.action === 'updateHouse';
    if (isUpdate) {
      const removed = deleteHouseRows_(ss, centerName, communityName, formatHouseNo_(addr.addrNo));
      Logger.log('UPDATE removed rows: ' + removed);
    }
    const keepPhotos = isUpdate ? (data.existingPhotos || {}) : null;

    /************ PET STRUCTURE ************/
    let pets = [];

    if (Array.isArray(data.pets) && data.pets.length > 0) {
      pets = data.pets;
    } else if (data.pet) {
      pets = [data.pet];
    } else {
      pets = [data];
    }

    /************ SHEETS ************/
    const centerSheet = getOrCreateSheet_(ss, centerName);
    const communitySheet = getOrCreateSheet_(ss, communityName);

    ensureHeaders_(centerSheet);
    ensureHeaders_(communitySheet);

    // ตั้งคอลัมน์บ้านเลขที่เป็นข้อความก่อนบันทึก
    setHouseNumberColumnText_(centerSheet);
    setHouseNumberColumnText_(communitySheet);

    /************ DRIVE ************/
    const rootFolder = getOrCreateFolder_(null, ROOT_FOLDER_NAME);
    // บังคับให้รูปภาพทั้งหมดบันทึกลงโฟลเดอร์ปี 2569
    const yearFolder = getOrCreateFolder_(rootFolder, PHOTO_FOLDER_YEAR);
    const centerFolder = getOrCreateFolder_(
      yearFolder, cleanName_(centerName)
    );
    const communityFolder = getOrCreateFolder_(
      centerFolder, cleanName_(communityName)
    );

    /************ OWNER FOLDER ************/
    const ownerNameClean = cleanName_(
      safe_(data.ownerName) || 'ไม่ระบุ'
    );

    const addrNoClean = cleanName_(
      formatHouseNo_(addr.addrNo)
    );

    const roadClean = cleanName_(addr.road || '');

    let ownerFolderName = 'เจ้าของ_' + ownerNameClean;

    if (addrNoClean) {
      ownerFolderName += '_บ้านเลขที่' + addrNoClean;
    }

    if (roadClean) {
      ownerFolderName += '_ถนน' + roadClean;
    }

    const ownerFolder = getOrCreateFolder_(
      communityFolder, ownerFolderName
    );

    let ownerPhotoUrl = savePhotoSmart_(
      data.ownerPhotoBase64 || data.ownerPhoto || data.ownerImage,
      ownerFolder,
      'รูปเจ้าของ'
    );
    if (!ownerPhotoUrl && keepPhotos && keepPhotos.owner) {
      ownerPhotoUrl = keepPhotos.owner;
    }

    const petMainFolder = getOrCreateFolder_(
      ownerFolder, 'สัตว์เลี้ยง'
    );

    /************ LOOP PET ************/
    pets.forEach((pet, index) => {
      let petPhotoUrl = '';

      const petPhotoData =
        pet.petPhotoBase64 ||
        pet.photoBase64 ||
        pet.image ||
        pet.photo ||
        pet["รูปภาพ"] ||
        pet["รูปสัตว์"] ||
        pet["รูปภาพสัตว์เลี้ยง"] ||
        pet["รูปภาพสัตว์เลี้ยง (อื่นๆโปรดระบุ)"] ||
        pet["รูปภาพสัตว์เลี้ยง (อื่น ๆ โปรดระบุ)"] ||
        pet["รูปภาพอื่นๆ"] ||
        data["รูปภาพ"] ||
        data["รูปสัตว์"] ||
        data["รูปภาพสัตว์เลี้ยง"] ||
        data["รูปภาพสัตว์เลี้ยง (อื่นๆโปรดระบุ)"] ||
        data["รูปภาพสัตว์เลี้ยง (อื่น ๆ โปรดระบุ)"] ||
        data["รูปภาพอื่นๆ"] ||
        '';

      if (petPhotoData) {
        const typeName = cleanName_(
          pet.animalType || data.animalType || 'ไม่ระบุ'
        );

        const petName = cleanName_(pet.name || 'ไม่ระบุ');

        petPhotoUrl = savePhotoSmart_(
          petPhotoData,
          petMainFolder,
          'ตัวที่' + (index + 1) + '_' + typeName + '_' + petName
        );
      }

      if (!petPhotoUrl && pet.imageUrl) {
        petPhotoUrl = saveImageFromUrl_(
          pet.imageUrl,
          petMainFolder,
          'สัตว์_' + (index + 1)
        );
      }
      if (!petPhotoUrl && keepPhotos && keepPhotos.pets && keepPhotos.pets[index]) {
        petPhotoUrl = keepPhotos.pets[index];
      }

      const rowObj = buildRowObject_(
        timestamp, year, centerName, communityName,
        data, pet, index, addr, ownerPhotoUrl, petPhotoUrl
      );

      appendSafe_(centerSheet, rowObj);
      appendSafe_(communitySheet, rowObj);
    });

    // บันทึกข้อมูลหลักเสร็จแล้ว
    // อัปเดต SUMMARY_AUTO + CHARTS_AUTO ทันทีในรอบเดียวกัน (ไม่พึ่ง trigger เบื้องหลัง)
    let summaryNote = '';
    try {
      buildSummaryAndDashboardCore_();
      SpreadsheetApp.flush();
      try {
        createHouseholdCharts_();
        SpreadsheetApp.flush();
      } catch (chartErr) {
        summaryNote = ' (สร้างกราฟไม่สำเร็จ: ' + (chartErr.message || chartErr) + ')';
        Logger.log('CHART ERROR: ' + (chartErr.stack || chartErr));
      }
    } catch (sumErr) {
      Logger.log('SUMMARY ERROR: ' + (sumErr.stack || sumErr));
      return output_(
        'บันทึกข้อมูลลงชีตศูนย์/ชุมชนแล้ว แต่อัปเดต SUMMARY_AUTO ไม่สำเร็จ: ' + (sumErr.message || sumErr),
        true
      );
    }

    return output_(
      (isUpdate ? 'แก้ไขข้อมูลบ้านเลขที่ ' + safe_(addr.addrNo) + ' สำเร็จ อัปเดต SUMMARY_AUTO และ CHARTS_AUTO แล้ว'
                : 'บันทึกข้อมูลสำเร็จ อัปเดต SUMMARY_AUTO และ CHARTS_AUTO แล้ว') + summaryNote,
      true
    );
  } catch (err) {
    Logger.log(err.stack || err);
    return output_(err.message || String(err));
  } finally {
    lock.releaseLock();
  }
}



/******************** BUILD ROW ********************/

/*
 * ลบข้อมูลบ้านออกจากชีตศูนย์ + ชีตชุมชน (เรียกจากปุ่มลบหมุดในฟอร์ม)
 * คืนจำนวนแถวที่ลบ แล้วสร้าง SUMMARY_AUTO + CHARTS_AUTO ใหม่ทันที
 */
function deleteHouseData_(data) {
  if (data.passcode !== 'babydog4321') {
    return output_('รหัสไม่ถูกต้อง ไม่สามารถลบข้อมูลได้');
  }

  const centerName = safe_(data.centerName || data.center);
  const communityName = safe_(data.communityName || data.community);
  const addrNo = safe_(data.addrNo);

  if (!centerName || !communityName || !addrNo) {
    return output_('ข้อมูลไม่ครบ (ต้องมีศูนย์ ชุมชน บ้านเลขที่)');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let total = 0;

  [centerName, communityName].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    const ci = headers.indexOf('ชุมชน');
    const hi = headers.indexOf('บ้านเลขที่');
    if (ci === -1 || hi === -1) return;

    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    const delRows = [];
    vals.forEach(function(r, i) {
      const comOk = String(r[ci]).trim() === communityName;
      const houseOk = formatHouseNo_(r[hi]) === addrNo;
      if (comOk && houseOk) delRows.push(i + 2);
    });

    delRows.sort(function(a, b) { return b - a; }).forEach(function(ro) {
      sheet.deleteRow(ro);
    });
    total += delRows.length;
  });

  if (total === 0) {
    return output_('ไม่พบข้อมูลบ้านเลขที่ ' + addrNo + ' ในชีต (อาจถูกลบไปแล้ว)');
  }

  // สร้างสรุปใหม่ทันทีเพื่อให้ข้อมูลตรงจริง
  try {
    buildSummaryAndDashboardCore_();
    SpreadsheetApp.flush();
    try {
      createHouseholdCharts_();
      SpreadsheetApp.flush();
    } catch (chartErr) {
      Logger.log('CHART ERROR: ' + (chartErr.stack || chartErr));
    }
  } catch (sumErr) {
    Logger.log('SUMMARY ERROR: ' + (sumErr.stack || sumErr));
  }

  return output_('ลบข้อมูลบ้านเลขที่ ' + addrNo + ' จำนวน ' + total + ' แถว และอัปเดตสรุปแล้ว', true);
}

function buildRowObject_(
  timestamp, year, centerName, communityName,
  data, pet, index, addr, ownerPhotoUrl, petPhotoUrl
) {
  function pick_() {
    for (let i = 0; i < arguments.length; i++) {
      const v = arguments[i];

      if (v !== undefined && v !== null && v !== '') {
        return String(v).trim();
      }
    }

    return '';
  }

  /************ ประเภทสัตว์ ************/
  const rawType = pick_(pet.animalType, data.animalType);

  const otherType = pick_(
    pet.otherAnimalType,
    data.otherAnimalType,
    pet["ชนิดสัตว์ (อื่น)"],
    data["ชนิดสัตว์ (อื่น)"],
    pet["อื่นๆโปรดระบุ"],
    data["อื่นๆโปรดระบุ"]
  );

  let animalType = 'ไม่ระบุ';

  if (rawType) {
    const t = rawType.trim();
    const low = t.toLowerCase();

    if (['สุนัข', 'หมา', 'dog'].includes(low)) {
      animalType = 'สุนัข';
    } else if (['แมว', 'cat'].includes(low)) {
      animalType = 'แมว';
    } else if (['โค', 'วัว', 'cow'].includes(low)) {
      animalType = 'โค';
    } else if (t === 'เป็ด') {
      animalType = 'เป็ด';
    } else if (t === 'ไก่') {
      animalType = 'ไก่';
    } else if (['สุกร', 'หมู'].includes(t)) {
      animalType = 'สุกร';
    } else if (t === 'นก') {
      animalType = 'นก';
    } else if (t === 'แพะ') {
      animalType = 'แพะ';
    } else if (t === 'ม้า') {
      animalType = 'ม้า';
    } else if (t.startsWith('อื่น') && otherType) {
      animalType = 'อื่นๆ (' + otherType + ')';
    } else {
      animalType = t;
    }
  }

  /************ ลักษณะการเลี้ยง ************/
  const raisingType = pick_(
    pet["ลักษณะการเลี้ยง"], data["ลักษณะการเลี้ยง"],
    pet["วิธีการเลี้ยง"], data["วิธีการเลี้ยง"],
    pet.raisingStyle, data.raisingStyle,
    pet.raisingType, data.raisingType,
    pet.raising, data.raising,
    pet.raising_type, data.raising_type
  ) || 'ไม่ระบุ';

  /************ พื้นที่การเลี้ยง ************/
  const raisingArea = pick_(
    pet["พื้นที่การเลี้ยง"], data["พื้นที่การเลี้ยง"],
    pet.raisingArea, data.raisingArea,
    pet.raising_area, data.raising_area
  ) || 'ไม่ระบุ';

  /************ สถานที่เลี้ยง ************/
  const raisingLocation = pick_(
    pet["สถานที่เลี้ยง"], data["สถานที่เลี้ยง"],
    pet.raisingLocation, data.raisingLocation
  ) || 'ไม่ระบุ';

  /************ วัคซีน ************/
  const rabiesStatus = pick_(
    pet.rabiesStatus, pet.rabies, data.rabiesStatus
  ) || 'ไม่ระบุ';

  return {
    "timestamp": timestamp,
    "ปีงบประมาณ": year,
    "ศูนย์บริการ": centerName,
    "ชุมชน": communityName,

    "ประเภทสัตว์": animalType,
    "ชนิดสัตว์ (อื่น)": otherType || '',

    "จำนวนสัตว์ (อื่น)": (() => {
      const val = pet.otherAnimalQty || data.otherAnimalQty;
      const num = Number(val);
      return num > 0 ? num : '';
    })(),

    "ลำดับสัตว์": pick_(pet.no, index + 1),
    "ชื่อสัตว์": pick_(pet.name) || 'ไม่ระบุ',
    "เพศ": pick_(pet.gender) || 'ไม่ระบุ',
    "อายุ (ปี)": pick_(pet.ageYear),
    "อายุ (เดือน)": pick_(pet.ageMonth),
    "สี / ตำหนิ": pick_(pet.color),

    "สถานะทำหมัน": pick_(pet.sterilization) || 'ไม่ระบุ',
    "วันที่ฉีดยาคุม": pick_(pet.contraceptiveDate),
    "สัตวแพทย์ผู้ฉีดยาคุม": pick_(pet.contraceptiveVet),

    "สถานะวัคซีนพิษสุนัขบ้า": rabiesStatus,
    "วันที่ฉีดวัคซีน": pick_(pet.rabiesDate),
    "สัตวแพทย์ผู้ฉีดวัคซีน": pick_(pet.rabiesVet),

    "ลักษณะการเลี้ยง": raisingType,
    "สถานที่เลี้ยง": raisingLocation,
    "พื้นที่การเลี้ยง": raisingArea,

    "ชื่อเจ้าของสัตว์": pick_(data.ownerName) || 'ไม่ระบุ',
    "เลขบัตรประชาชน": pick_(data.citizenId),
    "เบอร์โทรศัพท์มือถือ": pick_(data.phone),
    "เบอร์โทรศัพท์บ้าน": pick_(data.homePhone),

    // บ้านเลขที่บันทึกเป็นข้อความเสมอ
    "บ้านเลขที่": formatHouseNo_(addr.addrNo),

    "ถนน": pick_(addr.road),
    "ซอย": pick_(addr.soi),
    "ตำบล": pick_(addr.subdistrict),
    "อำเภอ": pick_(addr.district),
    "จังหวัด": pick_(addr.province),

    // พิกัดบ้านที่เลือกจากแผนที่ เช่น 15.228664, 104.861098
    "พิกัดที่เลือก": (() => {
      const hl = data.houseLocation || {};
      const latS = String(
        hl.latitude || hl.lat || data.houseLat || data.latitude || ''
      ).trim();
      const lngS = String(
        hl.longitude || hl.lng || data.houseLng || data.longitude || ''
      ).trim();
      if (!latS || !lngS) return '';
      return latS + ', ' + lngS;
    })(),

    "ผู้บันทึก": pick_(data.recorderName),
    "ตำแหน่ง": pick_(data.recorderRole),
    "เลขบัตรผู้บันทึก": pick_(data.surveyorId),

    "ลิงก์รูปเจ้าของ": ownerPhotoUrl || '',
    "ลิงก์รูปสัตว์": petPhotoUrl || ''
  };
}

/******************** HOUSE NUMBER FIX ********************/
/*
 * ป้องกัน Date object ถูกแปลงเป็นข้อความ GMT+0700
 *
 * หมายเหตุ:
 * หากข้อมูลเดิมถูกแปลงเป็นวันที่แล้ว จะไม่สามารถทราบ
 * บ้านเลขที่เดิมได้อย่างแน่นอนจากค่า Date เพียงอย่างเดียว
 */
function formatHouseNo_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';

    // บ้านเลขที่ลักษณะ 23/3 มักถูกชีตแปลงเป็นวันที่ 23/3/2026
    // จึงคืนค่าเฉพาะวัน/เดือน เพื่อไม่ให้ปีถูกเติมเข้ามาเป็นส่วนหนึ่งของบ้านเลขที่
    // หมายเหตุ: หากเป็นวันที่จริงที่ตั้งใจกรอก จะไม่สามารถแยกจากบ้านเลขที่ได้แน่นอน
    return Utilities.formatDate(
      value,
      TIME_ZONE,
      'd/M'
    );
  }

  return String(value).trim();
}

/*
 * ตั้งคอลัมน์บ้านเลขที่เป็น Plain text
 * เพื่อป้องกัน Google Sheets ตีความเป็นวันที่
 */
function setHouseNumberColumnText_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return;

  const headers = sheet.getRange(
    1, 1, 1, lastColumn
  ).getDisplayValues()[0].map(h => String(h).trim());

  const houseIndex = headers.indexOf('บ้านเลขที่');

  if (houseIndex === -1) return;

  const column = houseIndex + 1;

  sheet.getRange(
    2, column, Math.max(sheet.getMaxRows() - 1, 1), 1
  ).setNumberFormat('@');
}

/*
 * เพิ่มคอลัมน์ "พิกัดที่เลือก" ให้ชีตศูนย์ 1-7 และทุกชีตชุมชน
 * วิธีใช้: วางไฟล์นี้ทับของเดิม แล้วกด Run → addHouseLocationColumnNow หนึ่งครั้ง
 */
function addHouseLocationColumnNow() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const skip = ['SUMMARY_AUTO', 'DASHBOARD_AUTO', 'CHARTS_AUTO'];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (skip.indexOf(name) !== -1) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0].map(h => String(h).trim());

    // ข้ามชีตที่ไม่ใช่ชีตข้อมูล
    if (headers.indexOf('จังหวัด') === -1) return;

    // มีคอลัมน์แล้ว ข้าม
    if (headers.indexOf('พิกัดที่เลือก') !== -1) {
      Logger.log('มีแล้ว: ' + name);
      return;
    }

    // แทรกหลังคอลัมน์ "จังหวัด"
    const provinceCol = headers.indexOf('จังหวัด') + 1; // 1-based
    sheet.insertColumnAfter(provinceCol);
    sheet.getRange(1, provinceCol + 1).setValue('พิกัดที่เลือก');

    // ตั้งคอลัมน์เป็นข้อความ กันชีตแปลงค่าพิกัด
    sheet.getRange(
      2, provinceCol + 1,
      Math.max(sheet.getMaxRows() - 1, 1), 1
    ).setNumberFormat('@');

    Logger.log('เพิ่มแล้ว: ' + name);
  });
}


/************************************************************
 * BULK APPEND
 * บันทึกสัตว์หลายตัวพร้อมกัน
 ************************************************************/

function appendManySafe_(sheet, rowObjects) {

  if (!rowObjects || rowObjects.length === 0) {
    return;
  }

  // เพิ่มคอลัมน์ที่ขาดก่อน
  rowObjects.forEach(rowObj => {
    autoAddMissingColumns_(sheet, rowObj);
  });

  const lastColumn = sheet.getLastColumn();

  const headers = sheet.getRange(
    1, 1, 1, lastColumn
  ).getValues()[0].map(h => String(h).trim());

  const houseIndex = headers.indexOf('บ้านเลขที่');
  const timestampIndex = headers.indexOf('timestamp');

  const rows = rowObjects.map(rowObj => {

    return headers.map(header => {

      const value = rowObj[header];

      if (header === 'บ้านเลขที่') {
        return formatHouseNo_(value);
      }

      return value !== undefined && value !== null
        ? value
        : '';

    });

  });

  const startRow = sheet.getLastRow() + 1;

  // ตั้งบ้านเลขที่เป็นข้อความก่อนเขียน
  if (houseIndex >= 0) {

    sheet.getRange(
      startRow,
      houseIndex + 1,
      rows.length,
      1
    ).setNumberFormat('@');

  }

  // เขียนหลายแถวด้วยคำสั่งเดียว
  sheet.getRange(
    startRow,
    1,
    rows.length,
    headers.length
  ).setValues(rows);

  // ตั้ง timestamp
  if (timestampIndex >= 0) {

    sheet.getRange(
      startRow,
      timestampIndex + 1,
      rows.length,
      1
    ).setNumberFormat(TIMESTAMP_FORMAT);

  }

  // ย้ำบ้านเลขที่
  if (houseIndex >= 0) {

    sheet.getRange(
      startRow,
      houseIndex + 1,
      rows.length,
      1
    ).setNumberFormat('@');

  }

}

function fixTimestampFormatAllSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const lastColumn = sheet.getLastColumn();

    if (lastColumn < 1 || sheet.getLastRow() < 1) {
      Logger.log('ข้าม (ชีตว่าง): ' + name);
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0].map(h => String(h).trim().toLowerCase());

    const tsIndex = headers.indexOf('timestamp');

    if (tsIndex === -1) {
      Logger.log('ข้าม (ไม่พบหัวคอลัมน์ timestamp): ' + name);
      return;
    }

    // ตั้งรูปแบบทั้งคอลัมน์ (รวมแถวว่างด้านล่าง)
    sheet.getRange(2, tsIndex + 1, Math.max(sheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat(TIMESTAMP_FORMAT);

    Logger.log('แก้แล้ว: ' + name);
  });
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MASTER_HEADERS);
  }
}

function autoAddMissingColumns_(sheet, rowObj) {
  let headers = sheet.getRange(
    1, 1, 1, sheet.getLastColumn()
  ).getValues()[0].map(h => String(h).trim());

  Object.keys(rowObj).forEach(key => {
    if (headers.indexOf(key) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(key);
      headers.push(key);
    }
  });
}

/******************** HELPERS ********************/
function safe_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}

function getOrCreateFolder_(parent, name) {
  name = (name || 'Unknown').toString().trim();

  if (!parent) {
    const it = DriveApp.getFoldersByName(name);
    return it.hasNext() ? it.next() : DriveApp.createFolder(name);
  }

  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function cleanName_(text) {
  return (text || '').toString().trim()
    .replace(/\s+/g, '_')
    .replace(/[\\/:*?"<>|]/g, '');
}

/******************** SAVE PHOTO ********************/
function savePhotoSmart_(base64Data, folder, fileName) {
  if (!base64Data || typeof base64Data !== 'string') {
    return '';
  }

  try {
    let contentType = 'image/jpeg';

    base64Data = base64Data.trim();

    if (base64Data.indexOf('data:') === 0) {
      const match = base64Data.match(/^data:(.*?);base64,/);

      if (match && match[1]) {
        contentType = match[1];
      }

      base64Data = base64Data.split('base64,')[1];
    }

    base64Data = base64Data.replace(/\s/g, '');

    if (!base64Data || base64Data.length < 100) {
      Logger.log('BASE64 TOO SHORT');
      return '';
    }

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    return file.getUrl();

  } catch (err) {
    Logger.log('SAVE ERROR: ' + err);
    return '';
  }
}

function saveImageFromUrl_(url, folder, fileName) {
  try {
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob();

    const file = folder.createFile(blob).setName(fileName + '.jpg');

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    return file.getUrl();

  } catch (err) {
    Logger.log(err);
    return '';
  }
}

/******************** OUTPUT ********************/
function output_(msg, ok) {
  return ContentService.createTextOutput(
    JSON.stringify({
      status: ok ? 'success' : 'error',
      message: msg
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/******************** ZONE RESPONSIBILITY ********************/
/**
 * กำหนดเขตพื้นที่รับผิดชอบตามศูนย์บริการสาธารณสุข
 * เขต 1: ศูนย์ที่ 2 และ 6
 * เขต 2: ศูนย์ที่ 7
 * เขต 3: ศูนย์ที่ 1 และ 3
 * เขต 4: ศูนย์ที่ 4 และ 5
 */
function getResponsibilityZone_(centerName) {
  const name = String(centerName || '').trim();
  const match = name.match(/ศูนย์ที่\s*([1-7])/);
  if (!match) return 'ไม่ระบุเขต';

  const centerNo = Number(match[1]);
  const zoneMap = {
    2: 'เขต 1',
    6: 'เขต 1',
    7: 'เขต 2',
    1: 'เขต 3',
    3: 'เขต 3',
    4: 'เขต 4',
    5: 'เขต 4'
  };
  return zoneMap[centerNo] || 'ไม่ระบุเขต';
}

function responsibilityZoneOrder_(zone) {
  const order = { 'เขต 1': 1, 'เขต 2': 2, 'เขต 3': 3, 'เขต 4': 4, 'ไม่ระบุเขต': 99 };
  return order[zone] || 99;
}

/******************** NORMALIZE ANIMAL TYPE ********************/
function normalizeSummaryAnimalType_(rawType) {
  const original = String(rawType || '').trim();
  const t = original.toLowerCase();

  if (/^(สุนัข|หมา|dog)(มีเจ้าของ|ไม่มีเจ้าของ)?$/.test(t)) {
    return 'สุนัข';
  }

  if (/^(แมว|cat)(มีเจ้าของ|ไม่มีเจ้าของ)?$/.test(t)) {
    return 'แมว';
  }

  if (/^(โค|วัว|cow)$/.test(t)) return 'โค';
  if (t === 'เป็ด') return 'เป็ด';
  if (t === 'ไก่') return 'ไก่';
  if (/^(สุกร|หมู|pig)$/.test(t)) return 'สุกร';
  if (t === 'นก') return 'นก';
  if (t === 'แพะ') return 'แพะ';
  if (t === 'ม้า') return 'ม้า';

  if (t.startsWith('อื่น')) return 'อื่นๆ';

  return '';
}

/******************** SUMMARY AND DASHBOARD ********************/
function buildSummaryAndDashboardCore_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  let allData = [];

  /************ อ่านข้อมูลจากชีตชุมชน ************/
  // ต้องบันทึก log ภายใน forEach เท่านั้น เพราะ name และ sheet
  // จะถูกกำหนดค่าภายใน callback ของแต่ละชีต
  sheets.forEach(sheet => {
    const name = sheet.getName();

    if (name === 'SUMMARY_AUTO' || name === 'DASHBOARD_AUTO' || name === 'CHARTS_AUTO') {
      Logger.log('ข้ามชีตระบบ: ' + name);
      return;
    }

    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {
      Logger.log('ข้าม (มีแค่หัวตารางหรือชีตว่าง): ' + name);
      return;
    }

    const headers = values[0].map(h => String(h).trim());
    const communityIndex = headers.indexOf('ชุมชน');

    if (communityIndex === -1) {
      Logger.log('ข้าม (ไม่พบหัวคอลัมน์ "ชุมชน"): ' + name + ' | หัวที่มี: ' + headers.slice(0, 8).join(','));
      return;
    }

    let sampleCommunity = '';

    for (let i = 1; i < values.length; i++) {
      if (values[i][communityIndex]) {
        sampleCommunity = String(values[i][communityIndex]).trim();
        break;
      }
    }

    // ป้องกันการนับข้อมูลจากชีตศูนย์ซ้ำ
    if (!sampleCommunity || name !== sampleCommunity) {
      Logger.log('ข้าม (ชื่อชีตไม่ตรงค่าชุมชนในแถว): ชีต=' + name + ' ค่าชุมชน=' + (sampleCommunity || '(ว่าง)'));
      return;
    }

    Logger.log('อ่านชีต: ' + name + ' (' + (values.length - 1) + ' แถวข้อมูล)');

    values.slice(1).forEach(row => {
      const obj = {};

      headers.forEach((header, i) => {
        obj[header] = row[i];
      });

      // บ้านเลขที่จากเซลล์ที่เป็น Date จะถูกจัดการตรงนี้ด้วย
      obj['บ้านเลขที่'] = formatHouseNo_(obj['บ้านเลขที่']);

      allData.push(obj);
    });
  });

  Logger.log('รวมแถวข้อมูลทั้งหมด: ' + allData.length + ' แถว');
  if (allData.length === 0) {
    Logger.log('ไม่มีข้อมูล: ไม่สร้าง SUMMARY_AUTO (ตรวจว่าแถวในชีตชุมชนมีค่าคอลัมน์ "ชุมชน" ตรงกับชื่อชีตหรือไม่)');
    return;
  }

  /************ ตัวเก็บค่าสรุป ************/
  let animalGender = {};
  let steril = {};
  let vaccine = {};
  let center = {};
  let community = {};
  let recorder = {};
  let byCenter = {};
  let byCommunity = {};
  let centerCommunityAnimal = {};

  /************ รวมข้อมูล ************/
  allData.forEach(obj => {
    let type = String(obj['ประเภทสัตว์'] || 'ไม่ระบุ').trim();
    let qty = 1;

    if (type.startsWith('อื่น')) {
      const otherQty = Number(obj['จำนวนสัตว์ (อื่น)']);
      qty = otherQty > 0 ? otherQty : 1;
      type = 'อื่นๆ';
    }

    const gender = obj['เพศ'] || 'ไม่ระบุ';
    const ster = obj['สถานะทำหมัน'] || 'ไม่ระบุ';
    const vac = obj['สถานะวัคซีนพิษสุนัขบ้า'] || 'ไม่ระบุ';
    const cen = obj['ศูนย์บริการ'] || 'ไม่ระบุ';
    const com = obj['ชุมชน'] || 'ไม่ระบุ';
    const rec = obj['ผู้บันทึก'] || 'ไม่ระบุ';

    if (!animalGender[type]) {
      animalGender[type] = { ผู้: 0, เมีย: 0, อื่นๆ: 0 };
    }

    if (gender === 'ผู้') {
      animalGender[type]['ผู้'] += qty;
    } else if (gender === 'เมีย') {
      animalGender[type]['เมีย'] += qty;
    } else {
      animalGender[type]['อื่นๆ'] += qty;
    }

    steril[ster] = (steril[ster] || 0) + qty;
    vaccine[vac] = (vaccine[vac] || 0) + qty;
    center[cen] = (center[cen] || 0) + qty;
    community[com] = (community[com] || 0) + qty;
    recorder[rec] = (recorder[rec] || 0) + qty;

    const key = cen + '|' + com + '|' + type;

    centerCommunityAnimal[key] =
      (centerCommunityAnimal[key] || 0) + qty;

    if (!byCenter[cen]) {
      byCenter[cen] = {
        animalGender: {},
        steril: {},
        vaccine: {},
        recorder: {}
      };
    }

    if (!byCenter[cen].animalGender[type]) {
      byCenter[cen].animalGender[type] = {
        ผู้: 0, เมีย: 0, อื่นๆ: 0
      };
    }

    if (gender === 'ผู้') {
      byCenter[cen].animalGender[type]['ผู้'] += qty;
    } else if (gender === 'เมีย') {
      byCenter[cen].animalGender[type]['เมีย'] += qty;
    } else {
      byCenter[cen].animalGender[type]['อื่นๆ'] += qty;
    }

    byCenter[cen].steril[ster] =
      (byCenter[cen].steril[ster] || 0) + qty;

    byCenter[cen].vaccine[vac] =
      (byCenter[cen].vaccine[vac] || 0) + qty;

    byCenter[cen].recorder[rec] =
      (byCenter[cen].recorder[rec] || 0) + qty;

    if (!byCommunity[com]) {
      byCommunity[com] = {
        animalGender: {},
        steril: {},
        vaccine: {},
        recorder: {}
      };
    }

    if (!byCommunity[com].animalGender[type]) {
      byCommunity[com].animalGender[type] = {
        ผู้: 0, เมีย: 0, อื่นๆ: 0
      };
    }

    if (gender === 'ผู้') {
      byCommunity[com].animalGender[type]['ผู้'] += qty;
    } else if (gender === 'เมีย') {
      byCommunity[com].animalGender[type]['เมีย'] += qty;
    } else {
      byCommunity[com].animalGender[type]['อื่นๆ'] += qty;
    }

    byCommunity[com].steril[ster] =
      (byCommunity[com].steril[ster] || 0) + qty;

    byCommunity[com].vaccine[vac] =
      (byCommunity[com].vaccine[vac] || 0) + qty;

    byCommunity[com].recorder[rec] =
      (byCommunity[com].recorder[rec] || 0) + qty;
  });

  /************ เรียงลำดับ ************/
  function sortObj(obj) {
    return Object.fromEntries(
      Object.entries(obj).sort((a, b) => b[1] - a[1])
    );
  }

  steril = sortObj(steril);
  vaccine = sortObj(vaccine);
  center = sortObj(center);
  community = sortObj(community);
  recorder = sortObj(recorder);

  /************ สร้าง SUMMARY_AUTO ************/
  let sumSheet = ss.getSheetByName('SUMMARY_AUTO');

  if (!sumSheet) {
    sumSheet = ss.insertSheet('SUMMARY_AUTO');
  } else {
    sumSheet.clear();
  }

  /************ ประเภทสัตว์ + เพศ ************/
  sumSheet.getRange(1, 1, 1, 4)
    .setValues([['ประเภทสัตว์', 'ผู้', 'เมีย', 'อื่นๆ']]);

  let r = 2;

  Object.keys(animalGender).forEach(type => {
    const g = animalGender[type];

    sumSheet.getRange(r, 1, 1, 4).setValues([[
      type, g['ผู้'], g['เมีย'], g['อื่นๆ']
    ]]);

    r++;
  });

  /************ ตารางสรุปทั่วไป ************/
  function writeTable(title, obj, startRow) {
    sumSheet.getRange(startRow, 1).setValue(title);

    sumSheet.getRange(startRow + 1, 1, 1, 2)
      .setValues([['รายการ', 'จำนวน']]);

    let rr = startRow + 2;

    Object.entries(obj).forEach(([key, value]) => {
      sumSheet.getRange(rr, 1, 1, 2)
        .setValues([[key, value]]);
      rr++;
    });

    return { next: rr + 2 };
  }

  let sectionStart = r + 2;

  const t1 = writeTable('สถานะทำหมัน', steril, sectionStart);
  const t2 = writeTable('วัคซีนพิษสุนัขบ้า', vaccine, t1.next);
  const t3 = writeTable('ศูนย์บริการ', center, t2.next);
  const t4 = writeTable('ชุมชน', community, t3.next);
  const t5 = writeTable('ผู้บันทึก', recorder, t4.next);

  let nextRow = t5.next;

  /************ ศูนย์ + ชุมชน + ประเภทสัตว์ ************/
  sumSheet.getRange(nextRow, 1)
    .setValue('ศูนย์ + ชุมชน + ประเภทสัตว์');

  sumSheet.getRange(nextRow + 1, 1, 1, 4)
    .setValues([['ศูนย์', 'ชุมชน', 'ประเภทสัตว์', 'จำนวน']]);

  let rcc = nextRow + 2;

  Object.entries(centerCommunityAnimal).forEach(([key, value]) => {
    const [cen, com, type] = key.split('|');

    sumSheet.getRange(rcc, 1, 1, 4)
      .setValues([[cen, com, type, value]]);

    rcc++;
  });

  nextRow = rcc + 3;

  /************ สรุปแยกตามศูนย์ ************/
  Object.keys(byCenter).forEach(cen => {
    sumSheet.getRange(nextRow, 1).setValue('ศูนย์: ' + cen);
    nextRow++;

    Object.entries(byCenter[cen].steril).forEach(([key, value]) => {
      sumSheet.getRange(nextRow, 1, 1, 2)
        .setValues([[key, value]]);
      nextRow++;
    });

    nextRow += 2;
  });

  /************ สรุปแยกตามชุมชน ************/
  Object.keys(byCommunity).forEach(com => {
    sumSheet.getRange(nextRow, 1).setValue('ชุมชน: ' + com);
    nextRow++;

    Object.entries(byCommunity[com].steril).forEach(([key, value]) => {
      sumSheet.getRange(nextRow, 1, 1, 2)
        .setValues([[key, value]]);
      nextRow++;
    });

    nextRow += 2;
  });

  /****************************************************************
   * ตารางใหม่: สรุปสัตว์รายศูนย์-ชุมชน-บ้านเลขที่-ถนน
   ****************************************************************/
  const animalColumns = [
    'สุนัข', 'แมว', 'โค', 'เป็ด', 'ไก่',
    'สุกร', 'นก', 'แพะ', 'ม้า', 'อื่นๆ'
  ];

  const addressSummary = {};

  allData.forEach(obj => {
    const cen = String(obj['ศูนย์บริการ'] || 'ไม่ระบุ').trim();
    const com = String(obj['ชุมชน'] || 'ไม่ระบุ').trim();

    // อ่านบ้านเลขที่อย่างระมัดระวัง
    const houseNo = formatHouseNo_(obj['บ้านเลขที่']) || 'ไม่ระบุ';

    const road = String(obj['ถนน'] || 'ไม่ระบุ').trim();

    const key = JSON.stringify([cen, com, houseNo, road]);

    if (!addressSummary[key]) {
      addressSummary[key] = {
        zone: getResponsibilityZone_(cen),
        center: cen,
        community: com,
        houseNo: houseNo,
        road: road,
        counts: {
          'สุนัข': 0,
          'แมว': 0,
          'โค': 0,
          'เป็ด': 0,
          'ไก่': 0,
          'สุกร': 0,
          'นก': 0,
          'แพะ': 0,
          'ม้า': 0,
          'อื่นๆ': 0
        }
      };
    }

    const rawType = String(obj['ประเภทสัตว์'] || '').trim();
    const normalizedType = normalizeSummaryAnimalType_(rawType);

    // ไม่รวมประเภทที่ไม่ตรงกับคอลัมน์ที่กำหนด
    if (!normalizedType) return;

    let qty = 1;

    if (normalizedType === 'อื่นๆ') {
      const otherQty = Number(obj['จำนวนสัตว์ (อื่น)']);
      qty = otherQty > 0 ? otherQty : 1;
    }

    addressSummary[key].counts[normalizedType] += qty;
  });

  /************ เขียนตารางบ้านเลขที่ ************/
  nextRow += 1;

  sumSheet.getRange(nextRow, 1)
    .setValue('สรุปสัตว์รายศูนย์-ชุมชน-บ้านเลขที่-ถนน');

  const addressHeaders = [
    'เขตพื้นที่รับผิดชอบ', 'ศูนย์', 'ชุมชน', 'บ้านเลขที่', 'ถนน',
    'สุนัข', 'แมว', 'โค', 'เป็ด', 'ไก่',
    'สุกร', 'นก', 'แพะ', 'ม้า', 'อื่นๆ'
  ];

  sumSheet.getRange(
    nextRow + 1, 1, 1, addressHeaders.length
  ).setValues([addressHeaders]);

  const addressRows = Object.values(addressSummary)
    .sort((a, b) => {
      const z = responsibilityZoneOrder_(a.zone) - responsibilityZoneOrder_(b.zone);
      if (z !== 0) return z;

      const c1 = a.center.localeCompare(b.center, 'th');
      if (c1 !== 0) return c1;

      const c2 = a.community.localeCompare(b.community, 'th');
      if (c2 !== 0) return c2;

      const c3 = a.houseNo.localeCompare(b.houseNo, 'th');
      if (c3 !== 0) return c3;

      return a.road.localeCompare(b.road, 'th');
    })
    .map(item => [
      item.zone,
      item.center,
      item.community,
      item.houseNo,
      item.road,
      ...animalColumns.map(animal => item.counts[animal])
    ]);

  if (addressRows.length > 0) {
    // ตั้งคอลัมน์บ้านเลขที่ในตารางสรุปเป็นข้อความก่อนเขียน
    sumSheet.getRange(
      nextRow + 2, 4, addressRows.length, 1
    ).setNumberFormat('@');

    sumSheet.getRange(
      nextRow + 2, 1,
      addressRows.length,
      addressHeaders.length
    ).setValues(addressRows);

    // ย้ำรูปแบบข้อความหลังเขียน
    sumSheet.getRange(
      nextRow + 2, 4, addressRows.length, 1
    ).setNumberFormat('@');
  }

  /************ จัดรูปแบบ ************/
  sumSheet.getRange(
    nextRow + 1, 1, 1, addressHeaders.length
  ).setFontWeight('bold');

  sumSheet.getRange(
    nextRow + 1, 1,
    Math.max(addressRows.length + 1, 1),
    addressHeaders.length
  ).setVerticalAlignment('middle');

  /************ ตารางรวมจำนวนสัตว์ตามเขตและศูนย์บริการ ************/
  const zoneCenterTotals = {};
  Object.values(addressSummary).forEach(item => {
    const key = JSON.stringify([item.zone, item.center]);
    if (!zoneCenterTotals[key]) {
      zoneCenterTotals[key] = { zone: item.zone, center: item.center, counts: {} };
      animalColumns.forEach(animal => zoneCenterTotals[key].counts[animal] = 0);
    }
    animalColumns.forEach(animal => {
      zoneCenterTotals[key].counts[animal] += item.counts[animal] || 0;
    });
  });

  const zoneTableStart = nextRow + Math.max(addressRows.length, 1) + 4;
  sumSheet.getRange(zoneTableStart, 1).setValue('สรุปจำนวนสัตว์ตามเขตพื้นที่รับผิดชอบและศูนย์บริการ');
  const zoneHeaders = ['เขตพื้นที่รับผิดชอบ', 'ศูนย์', ...animalColumns, 'รวมทั้งหมด'];
  sumSheet.getRange(zoneTableStart + 1, 1, 1, zoneHeaders.length).setValues([zoneHeaders]);

  const zoneRows = Object.values(zoneCenterTotals)
    .sort((a, b) => {
      const z = responsibilityZoneOrder_(a.zone) - responsibilityZoneOrder_(b.zone);
      return z !== 0 ? z : a.center.localeCompare(b.center, 'th');
    })
    .map(item => {
      const counts = animalColumns.map(animal => item.counts[animal] || 0);
      const total = counts.reduce((sum, value) => sum + value, 0);
      return [item.zone, item.center, ...counts, total];
    });

  if (zoneRows.length) {
    sumSheet.getRange(zoneTableStart + 2, 1, zoneRows.length, zoneHeaders.length).setValues(zoneRows);
  }

  sumSheet.getRange(zoneTableStart + 1, 1, 1, zoneHeaders.length).setFontWeight('bold');
  sumSheet.getRange(zoneTableStart + 1, 1, Math.max(zoneRows.length + 1, 1), zoneHeaders.length).setVerticalAlignment('middle');

  sumSheet.setFrozenRows(1);
  sumSheet.autoResizeColumns(1, Math.max(addressHeaders.length, zoneHeaders.length));

  Logger.log(
    'สร้าง SUMMARY_AUTO สำเร็จ: ' +
    addressRows.length + ' รายการบ้าน และสรุปตามเขตพื้นที่รับผิดชอบ ' + zoneRows.length + ' รายการศูนย์' );


/************************************************************
 * HOUSEHOLD SURVEY SUMMARY
 * สรุปครัวเรือนที่สำรวจแยกตามศูนย์บริการ 1-7
 * พร้อมรายชื่อบ้านที่สำรวจและจำนวนสัตว์
 ************************************************************/

const householdTargets = {
  '1': 1783,
  '2': 2691,
  '3': 2440,
  '4': 1794,
  '5': 3168,
  '6': 1658,
  '7': 4768
};

const animalTypes = [
  'สุนัข', 'แมว', 'โค', 'เป็ด', 'ไก่',
  'สุกร', 'นก', 'แพะ', 'ม้า', 'อื่นๆ'
];

const householdMap = {};
const householdCounts = {
  '1': new Set(),
  '2': new Set(),
  '3': new Set(),
  '4': new Set(),
  '5': new Set(),
  '6': new Set(),
  '7': new Set()
};

// อ่านข้อมูลทุกครัวเรือนจากข้อมูลที่รวบรวมไว้
allData.forEach(obj => {

  const centerName = String(
    obj['ศูนย์บริการ'] || ''
  ).trim();

  const communityName = String(
    obj['ชุมชน'] || ''
  ).trim();

  const houseNo = formatHouseNo_(
    obj['บ้านเลขที่']
  );

  const road = String(
    obj['ถนน'] || ''
  ).trim();

  // รองรับชื่อศูนย์หลายรูปแบบ
  const match = centerName.match(
    /ศูนย์ที่\s*([1-7])/
  );

  if (!match) return;
  if (!communityName || !houseNo) return;

  const centerNo = match[1];

  // บ้านหนึ่งหลังนับเพียงครั้งเดียวในชุมชนเดียวกัน
  const householdKey = JSON.stringify([
  centerNo,
  communityName,
  houseNo,
  road
]);

  householdCounts[centerNo].add(householdKey);

  // เก็บรายละเอียดครัวเรือน
  if (!householdMap[householdKey]) {

    householdMap[householdKey] = {
      center: 'ศูนย์บริการสาธารณสุขศูนย์ที่ ' + centerNo,
      community: communityName,
      houseNo: houseNo,
      road: road || 'ไม่ระบุ',
      animals: {
        'สุนัข': 0,
        'แมว': 0,
        'โค': 0,
        'เป็ด': 0,
        'ไก่': 0,
        'สุกร': 0,
        'นก': 0,
        'แพะ': 0,
        'ม้า': 0,
        'อื่นๆ': 0
      }
    };
  }

  // นับจำนวนสัตว์ที่พบในบ้านหลังนี้
  const type = normalizeSummaryAnimalType_(
    obj['ประเภทสัตว์']
  );

  if (type) {
    let qty = 1;

    if (type === 'อื่นๆ') {
      const otherQty = Number(
        obj['จำนวนสัตว์ (อื่น)']
      );

      qty = otherQty > 0 ? otherQty : 1;
    }

    householdMap[householdKey].animals[type] += qty;
  }

});

// เรียงลำดับครัวเรือนตามศูนย์และชุมชน
const householdDetails = Object.values(householdMap)
  .sort((a, b) => {

    const centerA = Number(
      a.center.match(/([1-7])$/)[1]
    );

    const centerB = Number(
      b.center.match(/([1-7])$/)[1]
    );

    if (centerA !== centerB) {
      return centerA - centerB;
    }

    const communityCompare =
      a.community.localeCompare(b.community, 'th');

    if (communityCompare !== 0) {
      return communityCompare;
    }

    return a.houseNo.localeCompare(b.houseNo, 'th');
  });

// ตำแหน่งตารางต่อจากตารางก่อนหน้า
let householdTableStart =
  zoneTableStart + Math.max(zoneRows.length, 1) + 4;

// ======================================
// ตารางที่ 1 รายชื่อบ้านที่สำรวจ
// ======================================

sumSheet.getRange(
  householdTableStart, 1
).setValue('รายละเอียดครัวเรือนที่สำรวจข้อมูลสัตว์');

const detailHeaders = [
  'ศูนย์บริการ',
  'ชุมชน',
  'บ้านเลขที่',
  'ถนน',
  ...animalTypes,
  'รวมสัตว์'
];

sumSheet.getRange(
  householdTableStart + 1,
  1,
  1,
  detailHeaders.length
).setValues([detailHeaders]);

const detailRows = householdDetails.map(item => {

  const counts = animalTypes.map(
    type => item.animals[type] || 0
  );

  const totalAnimals = counts.reduce(
    (sum, value) => sum + value,
    0
  );

  return [
    item.center,
    item.community,
    item.houseNo,
    item.road,
    ...counts,
    totalAnimals
  ];
});

if (detailRows.length > 0) {

  // บ้านเลขที่เป็นข้อความ
  sumSheet.getRange(
    householdTableStart + 2,
    3,
    detailRows.length,
    1
  ).setNumberFormat('@');

  sumSheet.getRange(
    householdTableStart + 2,
    1,
    detailRows.length,
    detailHeaders.length
  ).setValues(detailRows);
}

// จัดรูปแบบตารางรายละเอียด
sumSheet.getRange(
  householdTableStart + 1,
  1,
  1,
  detailHeaders.length
).setFontWeight('bold');

// ======================================
// ตารางที่ 2 สรุปจำนวนครัวเรือนและร้อยละ
// ======================================

const summaryStart =
  householdTableStart + Math.max(detailRows.length, 1) + 4;

sumSheet.getRange(
  summaryStart, 1
).setValue('สรุปจำนวนครัวเรือนที่สำรวจ แยกตามศูนย์บริการสาธารณสุข');

const summaryHeaders = [
  'ศูนย์บริการสาธารณสุข',
  'จำนวนครัวเรือนเป้าหมาย',
  'จำนวนครัวเรือนที่สำรวจ',
  'จำนวนครัวเรือนคงเหลือ',
  'ร้อยละความครอบคลุม'
];

sumSheet.getRange(
  summaryStart + 1,
  1,
  1,
  summaryHeaders.length
).setValues([summaryHeaders]);

const summaryRows = [];

let totalTarget = 0;
let totalSurveyed = 0;

for (let i = 1; i <= 7; i++) {

  const centerNo = String(i);

  const target = householdTargets[centerNo];

  const surveyed = householdCounts[centerNo].size;

  const remaining = target - surveyed;

  const percentage = target > 0
    ? surveyed / target
    : 0;

  totalTarget += target;
  totalSurveyed += surveyed;

  summaryRows.push([
    'ศูนย์บริการสาธารณสุขศูนย์ที่ ' + i,
    target,
    surveyed,
    remaining,
    percentage
  ]);
}

// แถวรวมทั้ง 7 ศูนย์
summaryRows.push([
  'รวมทั้งหมด',
  totalTarget,
  totalSurveyed,
  totalTarget - totalSurveyed,
  totalTarget > 0 ? totalSurveyed / totalTarget : 0
]);

sumSheet.getRange(
  summaryStart + 2,
  1,
  summaryRows.length,
  summaryHeaders.length
).setValues(summaryRows);

// จัดรูปแบบหัวตาราง
sumSheet.getRange(
  summaryStart + 1,
  1,
  1,
  summaryHeaders.length
).setFontWeight('bold');

// เน้นแถวรวม
const totalRow =
  summaryStart + 2 + summaryRows.length - 1;

sumSheet.getRange(
  totalRow,
  1,
  1,
  summaryHeaders.length
).setFontWeight('bold');

// แสดงเปอร์เซ็นต์
sumSheet.getRange(
  summaryStart + 2,
  5,
  summaryRows.length,
  1
).setNumberFormat('0.00%');

// จัดรูปแบบ
sumSheet.getRange(
  summaryStart + 1,
  1,
  summaryRows.length + 1,
  summaryHeaders.length
).setVerticalAlignment('middle');

sumSheet.autoResizeColumns(
  1,
  Math.max(detailHeaders.length, summaryHeaders.length)
);

Logger.log(
  'สร้างตารางครัวเรือนสำเร็จ: ' +
  totalSurveyed + ' ครัวเรือน จาก ' +
  totalTarget + ' ครัวเรือน'
);
}

/************************************************************
 * BACKGROUND SUMMARY + CHART QUEUE (worker)
 ************************************************************/

const SUMMARY_WORKER_NAME = 'processSummaryQueue';

/**
 * คงไว้เผื่อ trigger เก่าที่ยังเหลืออยู่เรียกมา
 * (ปัจจุบัน doPost อัปเดตสรุปทันทีแล้ว ไม่สร้าง trigger ใหม่)
 */
function processSummaryQueue() {
  buildSummaryAndDashboardCore_();
  SpreadsheetApp.flush();
  try {
    createHouseholdCharts_();
    SpreadsheetApp.flush();
  } catch (chartErr) {
    Logger.log('CHART ERROR: ' + (chartErr.stack || chartErr));
  }
  Logger.log('อัปเดต SUMMARY_AUTO และ CHARTS_AUTO สำเร็จ');
}

/************************************************************
 * หาตำแหน่งวางกราฟใหม่
 ************************************************************/

function getNextHouseholdChartRow_(sheet) {

  const charts = sheet.getCharts();

  if (charts.length === 0) {
    return 2;
  }

  let maxRow = 1;

  charts.forEach(chart => {

    const info = chart.getContainerInfo();

    const anchorRow = info.getAnchorRow();

    const offsetHeight = info.getOffsetHeight();

    const estimatedRows = Math.ceil(
      offsetHeight / 20
    );

    maxRow = Math.max(
      maxRow,
      anchorRow + Math.max(estimatedRows, 20) + 3
    );

  });

  return maxRow + 2;
}


function rebuildHouseholdChartsNow() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('ระบบกำลังประมวลผล');
  }

  try {

    createHouseholdCharts_();

    SpreadsheetApp.flush();

    Logger.log('สร้างกราฟครัวเรือนสำเร็จ');

  } finally {

    lock.releaseLock();

  }

}

/************************************************************
 * HOUSEHOLD CHARTS AUTO - FINAL VERSION
 ************************************************************/

const HOUSEHOLD_CHART_TITLE_1 =
  'จำนวนครัวเรือนเป้าหมาย เทียบกับที่สำรวจ';

const HOUSEHOLD_CHART_TITLE_2 =
  'ร้อยละความครอบคลุมการสำรวจครัวเรือน';


function createHouseholdCharts_() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const summarySheet = ss.getSheetByName('SUMMARY_AUTO');

  if (!summarySheet) {
    throw new Error('ไม่พบ SUMMARY_AUTO');
  }

  let chartSheet = ss.getSheetByName('CHARTS_AUTO');

  if (!chartSheet) {
    chartSheet = ss.insertSheet('CHARTS_AUTO');
  }

  // --------------------------------
  // ค้นหาตารางสรุปครัวเรือน
  // --------------------------------

  const lastRow = summarySheet.getLastRow();

  const values = summarySheet
    .getRange(1, 1, lastRow, 5)
    .getDisplayValues();

  const title =
    'สรุปจำนวนครัวเรือนที่สำรวจ แยกตามศูนย์บริการสาธารณสุข';

  let titleIndex = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === title) {
      titleIndex = i;
      break;
    }
  }

  if (titleIndex === -1) {
    throw new Error('ไม่พบหัวข้อตารางครัวเรือน');
  }

  const headerRow = titleIndex + 2;
  const dataStartRow = headerRow + 1;

  // --------------------------------
  // ตรวจสอบหัวตาราง
  // --------------------------------

  const headers = summarySheet
    .getRange(headerRow, 1, 1, 5)
    .getDisplayValues()[0];

  if (headers[0] !== 'ศูนย์บริการสาธารณสุข') {
    throw new Error('หัวตารางไม่ถูกต้อง');
  }

  // --------------------------------
  // อ่านข้อมูล 7 ศูนย์
  // --------------------------------

  const data = summarySheet
    .getRange(dataStartRow, 1, 7, 5)
    .getValues();

  if (data.some(row => !row[0])) {
    throw new Error('ข้อมูลครัวเรือนทั้ง 7 ศูนย์ไม่ครบ');
  }

  // --------------------------------
  // สร้างตารางช่วยในชีตกราฟ
  // ใช้คอลัมน์ X:AA
  // --------------------------------

  const helperCol = 24;

  const helper = [
    [
      'ศูนย์บริการ',
      'เป้าหมาย',
      'สำรวจแล้ว',
      'ความครอบคลุม'
    ],
    ...data.map(row => [
      row[0],
      Number(row[1]) || 0,
      Number(row[2]) || 0,
      Number(row[4]) || 0
    ])
  ];

  // ตรวจสอบพื้นที่ช่วยก่อนเขียน
  // เพื่อไม่เขียนทับข้อมูลเดิม
  const helperRange = chartSheet.getRange(
    1, helperCol, 8, 4
  );

  const existingHelper = helperRange.getValues();

  const helperIsOurs =
    existingHelper[0][0] === 'ศูนย์บริการ';

  if (!helperIsOurs &&
      existingHelper.some(row =>
        row.some(cell => cell !== '' && cell !== null)
      )) {
    throw new Error(
      'พื้นที่ X:AA มีข้อมูลอื่นอยู่ กรุณาเลือกคอลัมน์ช่วยใหม่'
    );
  }

  helperRange.setValues(helper);

  chartSheet.hideColumns(helperCol, 4);

  SpreadsheetApp.flush();

  // --------------------------------
  // ลบเฉพาะกราฟครัวเรือนเดิม
  // --------------------------------

  chartSheet.getCharts().forEach(chart => {

    const title = String(
      chart.getOptions().get('title') || ''
    );

    if (
      title === HOUSEHOLD_CHART_TITLE_1 ||
      title === HOUSEHOLD_CHART_TITLE_2
    ) {
      chartSheet.removeChart(chart);
    }

  });

  SpreadsheetApp.flush();

  // --------------------------------
  // หาตำแหน่งวางกราฟ
  // --------------------------------

  let nextRow = 2;

  chartSheet.getCharts().forEach(chart => {

    const info = chart.getContainerInfo();

    const row = info.getAnchorRow() + 26;

    nextRow = Math.max(nextRow, row);

  });

  // --------------------------------
  // กราฟที่ 1
  // เป้าหมายเทียบกับสำรวจ
  // --------------------------------

  const chart1 = chartSheet.newChart()

    .setChartType(Charts.ChartType.COLUMN)

    .addRange(
      chartSheet.getRange(1, helperCol, 8, 3)
    )

    .setNumHeaders(1)

    .setOption(
      'title',
      HOUSEHOLD_CHART_TITLE_1
    )

    .setOption(
      'width',
      1100
    )

    .setOption(
      'height',
      450
    )

    .setOption(
      'legend',
      { position: 'top' }
    )

    .setOption(
      'hAxis',
      { title: 'ศูนย์บริการสาธารณสุข' }
    )

    .setOption(
      'vAxis',
      {
        title: 'จำนวนครัวเรือน',
        minValue: 0
      }
    )

    .setPosition(nextRow, 1, 0, 0)

    .build();

  chartSheet.insertChart(chart1);

  // --------------------------------
  // กราฟที่ 2
  // ร้อยละความครอบคลุม
  // --------------------------------

  const chart2 = chartSheet.newChart()

    .setChartType(Charts.ChartType.BAR)

    .addRange(
      chartSheet.getRange(1, helperCol, 8, 1)
    )

    .addRange(
      chartSheet.getRange(1, helperCol + 3, 8, 1)
    )

    .setNumHeaders(1)

    .setOption(
      'title',
      HOUSEHOLD_CHART_TITLE_2
    )

    .setOption(
      'width',
      1100
    )

    .setOption(
      'height',
      450
    )

    .setOption(
      'legend',
      { position: 'none' }
    )

    .setOption(
      'hAxis',
      {
        title: 'ร้อยละความครอบคลุม',
        format: '0.00%',
        minValue: 0
      }
    )

    .setPosition(nextRow + 25, 1, 0, 0)

    .build();

  chartSheet.insertChart(chart2);

  SpreadsheetApp.flush();

  Logger.log(
    'สร้างกราฟครัวเรือนสำเร็จ จำนวน 2 กราฟ'
  );
}


function appendSafe_(sheet, rowObj) {
  appendManySafe_(sheet, [rowObj]);
}


/************************************************************
 * REBUILD SUMMARY_AUTO + CHARTS_AUTO (RUN MANUALLY)
 * หลังวางโค้ดและบันทึก ให้เลือกฟังก์ชันนี้แล้วกด Run 1 ครั้ง
 ************************************************************/
function rebuildSummaryAndChartsNow() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('ระบบกำลังประมวลผลอยู่ กรุณารอสักครู่แล้วลองใหม่');
  }

  try {
    Logger.log('เริ่มสร้าง SUMMARY_AUTO');
    buildSummaryAndDashboardCore_();
    SpreadsheetApp.flush();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summary = ss.getSheetByName('SUMMARY_AUTO');
    if (!summary || summary.getLastRow() < 2) {
      throw new Error('SUMMARY_AUTO ยังไม่มีข้อมูล กรุณาตรวจสอบว่าชีตชุมชนมีข้อมูลและชื่อชีตตรงกับคอลัมน์ชุมชน');
    }

    Logger.log('สร้าง SUMMARY_AUTO แล้ว จำนวน ' + summary.getLastRow() + ' แถว');
    Logger.log('เริ่มสร้าง CHARTS_AUTO');
    createHouseholdCharts_();
    SpreadsheetApp.flush();

    const charts = ss.getSheetByName('CHARTS_AUTO');
    Logger.log('สร้าง CHARTS_AUTO แล้ว จำนวนกราฟ ' + (charts ? charts.getCharts().length : 0));
    Logger.log('เสร็จสมบูรณ์: SUMMARY_AUTO และ CHARTS_AUTO');
  } catch (err) {
    Logger.log('REBUILD ERROR: ' + (err.stack || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/******************** UPDATE / GET HOUSE ********************/
/* ลบแถวของบ้านหลังหนึ่งออกจากชีต (ใช้ร่วมกับลบและแก้ไข) คืนจำนวนแถวที่ลบ */
function deleteHouseRows_(ss, centerName, communityName, addrNo) {
  let total = 0;
  [centerName, communityName].forEach(function(name) {
    if (!name) return;
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    const ci = headers.indexOf('ชุมชน');
    const hi = headers.indexOf('บ้านเลขที่');
    if (ci === -1 || hi === -1) return;
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    const delRows = [];
    vals.forEach(function(r, i) {
      const comOk = String(r[ci]).trim() === communityName;
      const houseOk = formatHouseNo_(r[hi]) === addrNo;
      if (comOk && houseOk) delRows.push(i + 2);
    });
    delRows.sort(function(a, b) { return b - a; }).forEach(function(ro) {
      sheet.deleteRow(ro);
    });
    total += delRows.length;
  });
  return total;
}

/* ดึงข้อมูลบ้านหลังหนึ่ง คืนรูป payload เดียวกับที่หน้าเว็บส่ง */
function getHouseData_(data) {
  const centerName = safe_(data.centerName || data.center);
  const communityName = safe_(data.communityName || data.community);
  const addrNo = safe_(data.addrNo);
  if (!centerName || !communityName || !addrNo) {
    return output_('ข้อมูลไม่ครบ (ต้องมีศูนย์ ชุมชน บ้านเลขที่)');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(communityName) || ss.getSheetByName(centerName);
  if (!sheet || sheet.getLastRow() < 2) {
    return output_('ไม่พบข้อมูลบ้านเลขที่ ' + addrNo);
  }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0].map(function(h) { return String(h).trim(); });
  const ci = headers.indexOf('ชุมชน');
  const hi = headers.indexOf('บ้านเลขที่');
  if (ci === -1 || hi === -1) {
    return output_('ชีตยังไม่มีโครงข้อมูลที่อ่านได้');
  }
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  const rows = [];
  vals.forEach(function(r) {
    if (String(r[ci]).trim() === communityName && formatHouseNo_(r[hi]) === addrNo) {
      const o = {};
      headers.forEach(function(h, i) { o[h] = r[i]; });
      rows.push(o);
    }
  });
  if (!rows.length) {
    return output_('ไม่พบข้อมูลบ้านเลขที่ ' + addrNo);
  }
  const first = rows[0];
  const gv = function(row, k) { return row[k] !== undefined ? String(row[k]) : ''; };
  const g = function(k) { return gv(first, k); };
  let animalType = g('ประเภทสัตว์');
  let otherAnimalType = g('ชนิดสัตว์ (อื่น)');
  if (animalType.indexOf('อื่นๆ (') === 0 && animalType.slice(-1) === ')') {
    otherAnimalType = otherAnimalType || animalType.slice(5, -1);
    animalType = 'อื่นๆโปรดระบุ';
  }
  const pets = rows.map(function(r) {
    return {
      no: gv(r, 'ลำดับสัตว์'),
      name: gv(r, 'ชื่อสัตว์') === 'ไม่ระบุ' ? '' : gv(r, 'ชื่อสัตว์'),
      gender: gv(r, 'เพศ') || 'ไม่ระบุ',
      ageYear: gv(r, 'อายุ (ปี)'),
      ageMonth: gv(r, 'อายุ (เดือน)'),
      color: gv(r, 'สี / ตำหนิ'),
      sterilization: gv(r, 'สถานะทำหมัน') || 'ยังไม่ทำหมัน',
      contraceptiveDate: gv(r, 'วันที่ฉีดยาคุม'),
      contraceptiveVet: gv(r, 'สัตวแพทย์ผู้ฉีดยาคุม'),
      rabies: gv(r, 'สถานะวัคซีนพิษสุนัขบ้า') || 'ไม่เคยฉีดวัคซีนพิษสุนัขบ้า',
      rabiesDate: gv(r, 'วันที่ฉีดวัคซีน'),
      rabiesVet: gv(r, 'สัตวแพทย์ผู้ฉีดวัคซีน'),
      raisingStyle: gv(r, 'ลักษณะการเลี้ยง') || 'เลี้ยงในพื้นที่จำกัดตลอดเวลา',
      raisingLocation: gv(r, 'สถานที่เลี้ยง') || 'บ้านพักอาศัย',
      photoUrl: gv(r, 'ลิงก์รูปสัตว์')
    };
  });
  const hasPets = pets.some(function(p) { return p.name || p.ageYear || p.ageMonth || p.photoUrl; });
  const coords = g('พิกัดที่เลือก').split(',');
  const record = {
    centerName: centerName,
    community: communityName,
    animalType: animalType,
    otherAnimalType: otherAnimalType,
    otherAnimalQty: g('จำนวนสัตว์ (อื่น)'),
    ownerName: g('ชื่อเจ้าของสัตว์') === 'ไม่ระบุ' ? '' : g('ชื่อเจ้าของสัตว์'),
    recorderName: g('ผู้บันทึก'),
    recorderRole: g('ตำแหน่ง'),
    surveyorId: g('เลขบัตรผู้บันทึก'),
    citizenId: g('เลขบัตรประชาชน'),
    phone: g('เบอร์โทรศัพท์มือถือ'),
    homePhone: g('เบอร์โทรศัพท์บ้าน'),
    raisingArea: g('พื้นที่การเลี้ยง'),
    address: {
      addrNo: g('บ้านเลขที่'),
      road: g('ถนน'),
      soi: g('ซอย'),
      subdistrict: g('ตำบล') || 'ในเมือง',
      district: g('อำเภอ') || 'เมือง',
      province: g('จังหวัด') || 'อุบลราชธานี'
    },
    houseLocation: {
      latitude: (coords[0] || '').trim(),
      longitude: (coords[1] || '').trim(),
      status: 'surveyed'
    },
    ownerPhotoUrl: g('ลิงก์รูปเจ้าของ'),
    pets: hasPets ? pets : [],
    existingPhotos: {
      owner: g('ลิงก์รูปเจ้าของ'),
      pets: pets.map(function(p) { return p.photoUrl; })
    }
  };
  return output_({found: true, rows: rows.length, record: record}, true);
}

/******************** SURVEYOR STATS ********************/
/* นับบ้านเลขที่ที่ไม่ซ้ำของเลขบัตรผู้สำรวจ จากทุกชีต (ตัดชีตสรุปออก) */
function getSurveyorStats_(data) {
  const sid = String(safe_(data.recorderName)).replace(/\s+/g, ' ').trim();
  if (!sid) {
    return output_('กรุณาระบุชื่อ-นามสกุลผู้บันทึก');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const skip = ['SUMMARY_AUTO', 'DASHBOARD_AUTO', 'CHARTS_AUTO'];
  const seen = {};
  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    if (skip.indexOf(name) >= 0) return;
    if (sheet.getLastRow() < 2) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    const ii = headers.indexOf('ผู้บันทึก');
    const ci = headers.indexOf('ชุมชน');
    const hi = headers.indexOf('บ้านเลขที่');
    if (ii === -1 || ci === -1 || hi === -1) return;
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    vals.forEach(function(r) {
      if (String(r[ii]).replace(/\s+/g, ' ').trim() !== sid) return;
      const community = String(r[ci]).trim();
      const addrNo = formatHouseNo_(r[hi]);
      if (!community || !addrNo) return;
      seen[community + '|' + addrNo] = community;
    });
  });
  const houses = Object.keys(seen).map(function(k) {
    const p = k.split('|');
    return {community: p[0], addrNo: p[1]};
  });
  houses.sort(function(a, b) {
    return a.community < b.community ? -1 : (a.community > b.community ? 1 : (a.addrNo < b.addrNo ? -1 : 1));
  });
  return output_({totalHouses: houses.length, houses: houses}, true);
}
