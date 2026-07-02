// ═══════════════════════════════════════════
//  Smart Farmer Database — Apps Script
//  รองรับเจ้าหน้าที่หลายจังหวัด
// ═══════════════════════════════════════════

const SHEET_MASTER  = 'Master';
const SHEET_PLOT    = 'Plot';
const SHEET_HISTORY = 'History';
const SPREADSHEET_ID = '1LQwVM9Ybq2woYpokDkavPoA-EqukLaLiqNFoFfDQq6E';

// ── สร้างเมนูเมื่อเปิด Sheets ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Smart Farmer')
    .addItem('👤 ลงทะเบียนเกษตรกรใหม่', 'openRegister')
    .addItem('📋 บันทึกประวัติประจำปี',  'openHistory')
    .addSeparator()
    .addItem('📥 นำเข้าข้อมูลจาก CSV', 'openImport')
    .addItem('🔍 ค้นหาเกษตรกร',         'openSearch')
    .addToUi();
}

function openRegister() {
  const html = HtmlService.createHtmlOutputFromFile('Register')
    .setTitle('ลงทะเบียนเกษตรกรใหม่')
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openHistory() {
  const html = HtmlService.createHtmlOutputFromFile('History')
    .setTitle('บันทึกประวัติการพัฒนา')
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openSearch() {
  const html = HtmlService.createHtmlOutputFromFile('Search')
    .setTitle('ค้นหาเกษตรกร')
    .setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── ตรวจสอบเลขบัตรซ้ำ ──
function checkDuplicate(id) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const ids    = master.getRange('A2:A').getValues().flat()
                       .map(v => String(v).trim());
  const found  = ids.includes(String(id).trim());
  if (!found) return { duplicate: false };
  // หาชื่อ
  const idx = ids.indexOf(String(id).trim());
  const row = master.getRange(idx + 2, 1, 1, 4).getValues()[0];
  return { duplicate: true, name: row[1] + row[2] + ' ' + row[3] };
}

// ── VLOOKUP ดึงชื่อจาก Master ──
function lookupName(id) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const data   = master.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      return {
        found:    true,
        title:    data[i][1],
        fname:    data[i][2],
        lname:    data[i][3],
        province: data[i][9],
        status:   data[i][24]
      };
    }
  }
  return { found: false };
}
// ═══════════════════════════════════════════
//  Web App Entry Point
// ═══════════════════════════════════════════

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';
  const id   = (e && e.parameter && e.parameter.id)   || '';
  const html = HtmlService.createTemplateFromFile('WebApp');
  html.page     = page;
  html.farmerId = id;
  return html.evaluate()
    .setTitle('Smart Farmer — ส.ป.ก.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันช่วย include ไฟล์ HTML อื่น
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── ดึงข้อมูล Dashboard ──
function getDashboardData() {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const master  = ss.getSheetByName(SHEET_MASTER);
    const history = ss.getSheetByName(SHEET_HISTORY);

    const mData = master.getDataRange().getValues().slice(1).filter(r => r[0]);
    const hData = history.getDataRange().getValues().slice(1).filter(r => r[0]);

    // คำนวณสถานะสูงสุดของแต่ละคนจาก History
    const statusByFarmer = {};
    hData.forEach(r => {
      const id = String(r[0]).trim();
      if (!statusByFarmer[id]) statusByFarmer[id] = [];
      statusByFarmer[id].push(r[39]); // column AN = ผลประเมิน
    });

    const getStatus = (id) => getHighestStatus(statusByFarmer[id] || []);

    const total = mData.length;
    const tb    = mData.filter(r => getStatus(String(r[0]).trim()) === 'ต้นแบบ').length;
    const ex    = mData.filter(r => getStatus(String(r[0]).trim()) === 'Existing').length;
    const dv    = mData.filter(r => getStatus(String(r[0]).trim()) === 'Developing').length;

    const byProv = {};
    mData.forEach(r => {
      const p = r[9] || 'ไม่ระบุ';
      const s = getStatus(String(r[0]).trim());
      if (!byProv[p]) byProv[p] = { tb:0, ex:0, dv:0 };
      if (s==='ต้นแบบ') byProv[p].tb++;
      else if (s==='Existing') byProv[p].ex++;
      else byProv[p].dv++;
    });

    const incRows = hData.filter(r => r[19] && r[20]);
    const avgInc1 = incRows.length
      ? Math.round(incRows.reduce((s,r) => s + Number(String(r[19]).replace(/,/g,'')), 0) / incRows.length) : 0;
    const avgInc2 = incRows.length
      ? Math.round(incRows.reduce((s,r) => s + Number(String(r[20]).replace(/,/g,'')), 0) / incRows.length) : 0;

    const incByProv = {};
    const byYear   = {};

    hData.forEach(r => {
      const rid  = String(r[0]).trim();
      const mRow = mData.find(m => String(m[0]).trim() === rid);
      const prov = mRow ? mRow[9] : 'ไม่ระบุ';
      if (!incByProv[prov]) incByProv[prov] = { before:[], after:[] };
      if (r[19]) incByProv[prov].before.push(Number(String(r[19]).replace(/,/g,'')));
      if (r[20]) incByProv[prov].after.push(Number(String(r[20]).replace(/,/g,'')));

      const y = String(r[2]) || 'ไม่ระบุ';
      if (!byYear[y]) byYear[y] = { count:0, tb:0, ex:0, dv:0 };
      byYear[y].count++;
      if (r[39]==='ต้นแบบ')        byYear[y].tb++;
      else if (r[39]==='Existing')  byYear[y].ex++;
      else if (r[39]==='Developing') byYear[y].dv++;
    });

    const recent = hData.slice(-10).reverse().map(r => {
      const rid  = String(r[0]).trim();
      const mRow = mData.find(m => String(m[0]).trim() === rid);
      return {
        id:     rid,
        name:   mRow ? mRow[1]+mRow[2]+' '+mRow[3] : '—',
        prov:   mRow ? mRow[9] : '—',
        year:   r[2],
        result: r[39],
        date:   r[43] ? String(r[43]).substring(0,10) : '—'
      };
    });

    return { total, tb, ex, dv, byProv, avgInc1, avgInc2,
             incByProv, byYear, recent };

  } catch(err) {
    return { 
      error: err.message,
      total:0, tb:0, ex:0, dv:0,
      byProv:{}, avgInc1:0, avgInc2:0,
      incByProv:{}, byYear:{}, recent:[]
    };
  }
}
// ── บันทึกเกษตรกรใหม่ → Master ──
function saveMaster(form) {
  const dup = checkDuplicate(form.id);
  if (dup.duplicate) return { ok: false, msg: 'เลขบัตรนี้มีในระบบแล้ว: ' + dup.name };

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const user   = Session.getActiveUser().getEmail();
  const now    = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

  master.appendRow([
    form.id,         // A  เลขบัตร
    form.title,      // B  คำนำหน้า
    form.fname,      // C  ชื่อ
    form.lname,      // D  นามสกุล
    form.bday,       // E  วันเกิด
    form.hno,        // F  บ้านเลขที่
    form.moo,        // G  หมู่ที่
    form.sub,        // H  ตำบล
    form.dis,        // I  อำเภอ
    form.prov,       // J  จังหวัด
    form.mem,        // K  สมาชิกครัวเรือน
    form.lab,        // L  แรงงานเกษตร
    form.edu,        // M  ระดับการศึกษา (ย้ายมา)
    form.degree,     // N  วุฒิการศึกษา (ย้ายมา)
    form.major,      // O  สาขา (ย้ายมา)
    form.own,        // P  เจ้าของแปลง (ย้ายมา)
    form.ownerRel,   // Q  ความสัมพันธ์ (ย้ายมา)
    form.ownerName,  // R  ชื่อเจ้าของแปลง (ย้ายมา)
    form.ownerID,    // S  เลขบัตรเจ้าของแปลง (ย้ายมา)
    form.regAgri,    // T  ขึ้นทะเบียนเกษตรอำเภอ
    form.regRubber,  // U  ขึ้นทะเบียนการยาง
    form.coop,       // V  สหกรณ์
    form.bigfarm,    // W  แปลงใหญ่
    form.smce,       // X  วิสาหกิจ
    'Developing',    // Y  สถานะเริ่มต้น
    '',              // Z  จำนวนปี (สูตร)
    '',              // AA ปีล่าสุด (สูตร)
    user,            // AB บันทึกโดย
    now              // AC วันที่บันทึก
  ]);

  const lastRow = master.getLastRow();
  master.getRange(lastRow, 26).setFormula(
    `=IFERROR(COUNTIF(History!A:A,A${lastRow}),0)`);
  master.getRange(lastRow, 27).setFormula(
    `=IFERROR(MAXIFS(History!C:C,History!A:A,A${lastRow}),"—")`);

  return { ok: true, msg: 'บันทึก ' + form.title + form.fname + ' ' + form.lname + ' สำเร็จ' };
}

// ── บันทึกข้อมูลแปลง → Plot ──
function savePlot(form) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const plot = ss.getSheetByName(SHEET_PLOT);

  plot.appendRow([
    form.id,          // A FK
    '',               // B ชื่อ (สูตร VLOOKUP)
    form.arai,        // C ไร่ทั้งหมด
    form.angan,       // D งาน
    form.awa,         // E วา
    form.jRai,        // F ร่วม ไร่
    form.jNgan,       // G ร่วม งาน
    form.jWa,         // H ร่วม วา
    form.plotNo,      // I แปลงเลขที่
    String(form.rawang), // J กลุ่ม/ระวาง ← บังคับเป็น String
    form.plotMoo,     // K หมู่ที่แปลง
    form.plotSub,     // L ตำบลแปลง
    form.plotDis,     // M อำเภอแปลง
    form.plotProv,    // N จังหวัดแปลง
    form.water,       // O แหล่งน้ำ
    form.power,       // P ไฟฟ้า
    form.signal,      // Q สัญญาณ
    form.road,        // R ถนน
    form.transport,   // S คมนาคม
  ]);

  // ใส่ VLOOKUP ชื่อ
  const lastRow = plot.getLastRow();
  plot.getRange(lastRow, 2).setFormula(
    `=IFERROR(VLOOKUP(A${lastRow},Master!A:D,2,FALSE)&VLOOKUP(A${lastRow},Master!A:D,3,FALSE)&" "&VLOOKUP(A${lastRow},Master!A:D,4,FALSE),"— ไม่พบ —")`);

  return { ok: true };
}

// ── บันทึกประวัติ → History ──
function saveHistory(form) {
  const check = lookupName(form.id);
  if (!check.found) return { ok: false, msg: 'ไม่พบเลขบัตรในระบบ — กรุณาลงทะเบียนเกษตรกรก่อน' };

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hist = ss.getSheetByName(SHEET_HISTORY);
  const rows = hist.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(form.id).trim() &&
        String(rows[i][2]) === String(form.year)) {
      return { ok: false, msg: 'มีประวัติปี ' + form.year + ' ของเกษตรกรรายนี้แล้ว' };
    }
  }

  const user = Session.getActiveUser().getEmail();
  const now  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');

  hist.appendRow([
    form.id,          // A  [0]  FK
    '',               // B  [1]  ชื่อ VLOOKUP
    form.year,        // C  [2]  ปีงบประมาณ
    form.act1,        // D  [3]  ทำนา
    form.act2,        // E  [4]  ปลูกผัก
    form.act3,        // F  [5]  ผลไม้
    form.act4,        // G  [6]  สมุนไพร
    form.act5,        // H  [7]  พืชไร่
    form.act6,        // I  [8]  ปศุสัตว์
    form.act7,        // J  [9]  แปรรูป
    form.act8,        // K  [10] อื่นๆ
    form.areaRai,     // L  [11] พื้นที่ไร่
    form.yieldAmt,    // M  [12] ผลผลิต จำนวน
    form.yieldUnit,   // N  [13] ผลผลิต หน่วย
    form.ch1,         // O  [14] พ่อค้า/ล้ง
    form.ch2,         // P  [15] ตลาดชุมชน
    form.ch3,         // Q  [16] ห้าง
    form.ch4,         // R  [17] ออนไลน์
    form.page,        // S  [18] ชื่อเพจ
    form.inc1,        // T  [19] รายได้ก่อน
    form.inc2,        // U  [20] รายได้หลัง
    form.incSummary,  // V  [21] สรุปรายได้
    '',               // W  [22] ผลต่าง บาท (สูตร)
    '',               // X  [23] ผลต่าง % (สูตร)
    form.k1,          // Y  [24] ความรู้ — บริหารแปลง
    form.k2,          // Z  [25] ความรู้ — พัฒนาผลิตภัณฑ์
    form.k3,          // AA [26] ความรู้ — ช่องทางขาย
    form.k4,          // AB [27] ความรู้ — วางแผนอนาคต
    form.k5,          // AC [28] ความรู้ — เผยแพร่
    form.kNot,        // AD [29] ไม่ได้นำไปใช้ เนื่องจาก
    form.c1,          // AE [30] เปลี่ยนแปลง — รายได้เพิ่ม
    form.c2,          // AF [31] เปลี่ยนแปลง — รายจ่ายลด
    form.c3,          // AG [32] เปลี่ยนแปลง — ต้นทุนลด
    form.c4,          // AH [33] เปลี่ยนแปลง — วางแผนผลิต
    form.c5,          // AI [34] เปลี่ยนแปลง — วางแผนการตลาด
    form.c6,          // AJ [35] เปลี่ยนแปลง — ผลผลิตเพิ่ม
    form.c7,          // AK [36] เปลี่ยนแปลง — ผลผลิตมีคุณภาพ
    form.cOther,      // AL [37] เปลี่ยนแปลง — อื่นๆ ระบุ
    form.problem,     // AM [38] ปัญหา/อุปสรรค
    form.result,      // AN [39] ผลการประเมิน SF
    form.sfType,      // AO [40] ประเภทเกษตรกร
    form.staff,       // AP [41] ผู้จัดเก็บ
    form.staffTel,    // AQ [42] เบอร์
    form.recordDate,  // AR [43] วันที่ลงข้อมูล
    user,             // AS [44] บันทึกโดย email
    now               // AT [45] วันที่บันทึกในระบบ
  ]);

  const lastRow = hist.getLastRow();
  hist.getRange(lastRow, 2).setFormula(
`=IFERROR(VLOOKUP(A${lastRow},Master!A:D,2,FALSE)&VLOOKUP(A${lastRow},Master!A:D,3,FALSE)&" "&VLOOKUP(A${lastRow},Master!A:D,4,FALSE),"— ไม่พบ —")`);
  hist.getRange(lastRow, 23).setFormula(`=IFERROR(U${lastRow}-T${lastRow},"")`);
  hist.getRange(lastRow, 24).setFormula(`=IFERROR(ROUND((U${lastRow}-T${lastRow})/T${lastRow}*100,2),"—")`);

  updateStatus(form.id, form.result);

  return { ok: true, msg: 'บันทึกประวัติปี ' + form.year + ' ของ '
    + check.title + check.fname + ' ' + check.lname + ' สำเร็จ' };
}

// ── อัปเดตสถานะใน Master ──
const STATUS_RANK = { 'Developing': 1, 'Existing': 2, 'ต้นแบบ': 3 };

function updateStatus(id, newStatus) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const ids    = master.getRange('A2:A').getValues().flat();
  const idx    = ids.findIndex(v => String(v).trim() === String(id).trim());
  if (idx < 0) return;

  const currentStatus = master.getRange(idx + 2, 25).getValue(); // column Y = 25
  const currentRank   = STATUS_RANK[currentStatus] || 0;
  const newRank       = STATUS_RANK[newStatus] || 0;

  // เขียนทับเฉพาะตอนสถานะใหม่สูงกว่าเดิม
  if (newRank > currentRank) {
    master.getRange(idx + 2, 25).setValue(newStatus);
  }
}

// ── ค้นหาเกษตรกร ──
function searchFarmer(query, province, status) {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const hist   = ss.getSheetByName(SHEET_HISTORY);
  const mData  = master.getDataRange().getValues();
  const hData  = hist.getDataRange().getValues();

  const q = (query || '').toLowerCase().trim();

  const results = mData.slice(1).filter(r => {
    const matchQ = !q ||
      String(r[0]).includes(q) ||
      (r[2] + r[3]).toLowerCase().includes(q);
    const matchP = !province || r[9] === province;
    const matchS = !status   || r[21] === status;
    return matchQ && matchP && matchS;
  }).map(r => {
    const id = String(r[0]).trim();
    const history = hData.slice(1)
      .filter(h => String(h[0]).trim() === id)
      .map(h => ({ year: h[2], result: h[36] }));
    return {
      id, title: r[1], fname: r[2], lname: r[3],
      province: r[9], status: r[24],
      years: history.length,
      history
    };
  });

  return results;
}
function getFarmerProfile(id) {
  try {
    Logger.log('id received: [' + id + '] type: ' + typeof id);

    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const master = ss.getSheetByName(SHEET_MASTER);
    const hist   = ss.getSheetByName(SHEET_HISTORY);
    const plot   = ss.getSheetByName(SHEET_PLOT);

    const sid = String(id).replace(/\.0+$/, '').trim();
    Logger.log('sid: [' + sid + ']');

    const mLast = master.getLastRow();
    if(mLast < 2) return { found: false };
    const mData = master.getRange(2, 1, mLast - 1, 29).getValues();

    mData.forEach((r,i) => {
      if(r[0]) Logger.log('row '+(i+2)+': ['+String(r[0])+'] match='+(String(r[0]).trim()===sid));
    });

    const mRow = mData.find(r => String(r[0]).trim() === sid);
    Logger.log('mRow found: ' + (mRow ? 'YES' : 'NO'));
    if(!mRow) return { found: false };

    const hLast = hist.getLastRow();
    const hData = hLast < 2 ? [] :
      hist.getRange(2, 1, hLast - 1, 46).getValues()
          .filter(r => String(r[0]).replace(/\.0+$/, '').trim() === sid);

    const pLast = plot.getLastRow();
    const pData = pLast < 2 ? [] :
      plot.getRange(2, 1, pLast - 1, 19).getValues()
          .filter(r => String(r[0]).replace(/\.0+$/, '').trim() === sid);

    const history = hData.map(r => ({
      year:      r[2],
      acts:      [r[3],r[4],r[5],r[6],r[7],r[8],r[9],r[10]].filter(Boolean).join(', '),
      areaRai:   r[11],
      yield:     r[12]&&r[13] ? r[12]+' '+r[13] : '',
      ch:        [r[14]?'พ่อค้า/ล้ง':'',r[15]?'ตลาดชุมชน':'',
                  r[16]?'ห้าง':'',r[17]?'ออนไลน์':''].filter(Boolean).join(', '),
      inc1:    String(r[19] || ''),
      inc2:    String(r[20] || ''),
      incSummary:r[21], 
      incDiff: String(r[22] || ''),
      incPct:  String(r[23] || ''),
      knowledge: [r[24]?'บริหารแปลง':'',r[25]?'พัฒนาผลิตภัณฑ์':'',
                  r[26]?'ช่องทางขาย':'',r[27]?'วางแผนอนาคต':'',
                  r[28]?'เผยแพร่':''].filter(Boolean).join(', '),
      kNot:      r[29],
      changes:   [r[30]?'รายได้เพิ่ม':'',r[31]?'รายจ่ายลด':'',
                  r[32]?'ต้นทุนลด':'',r[33]?'วางแผนผลิต':'',
                  r[34]?'วางแผนการตลาด':'',r[35]?'ผลผลิตเพิ่ม':'',
                  r[36]?'ผลผลิตมีคุณภาพ':'',r[37]||''].filter(Boolean).join(', '),
      problem:   r[38],
      result:    r[39], sfType: r[40],
      staff:     r[41],
      recordDate: (function(){
        const d = r[43];
        if(!d) return '—';
        if(d instanceof Date){
          const months = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                          'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
          return d.getDate()+' '+months[d.getMonth()]+' '+d.getFullYear().toString().slice(-2);
        }
        return String(d).substring(0,10);
      })()
    })).sort((a,b) => b.year - a.year);

      const uniquePlots = [];
      const seenPlots   = new Set();
      pData.forEach(r => {
        const key = String(r[8]) + '_' + String(r[9]);
        if(!seenPlots.has(key)){
          seenPlots.add(key);
          uniquePlots.push(r);
        }
      });

      const plots = uniquePlots.map(r => ({
        aTotal:   [r[2]?r[2]+' ไร่':'',r[3]?r[3]+' งาน':'',r[4]?r[4]+' ตร.ว.':''].filter(Boolean).join(' '),
        aJoint:   [r[5]?r[5]+' ไร่':'',r[6]?r[6]+' งาน':'',r[7]?r[7]+' ตร.ว.':''].filter(Boolean).join(' '),
        plotNo:   r[8], rawang: r[9],
        location: [r[10]?'หมู่ '+r[10]:'',r[11],r[12],r[13]].filter(Boolean).join(' '),
        water:r[14], power:r[15], signal:r[16], road:r[17], transport:r[18]
      }));

    return {
      found:true, id:String(mRow[0]),
      title:mRow[1], fname:mRow[2], lname:mRow[3],
      bday: (function(){
        const d = mRow[4];
        if(!d) return '—';
        if(d instanceof Date){
          const day = d.getDate();
          const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                          'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
          return day + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
        }
        return String(d);
      })(),
      address:[mRow[5]?'บ้านเลขที่ '+mRow[5]:'',
              mRow[6]?'หมู่ '+mRow[6]:'',
              mRow[7]?'ต.'+mRow[7]:'',
              mRow[8]?'อ.'+mRow[8]:'',
              mRow[9]].filter(Boolean).join(' '),
      province:mRow[9], mem:mRow[10], lab:mRow[11],
      edu:mRow[12], degree:mRow[13], major:mRow[14],
      own:mRow[15], ownerRel:mRow[16], ownerName:mRow[17], ownerID:mRow[18],
      regAgri:mRow[19], regRubber:mRow[20], coop:mRow[21], bigfarm:mRow[22], smce:mRow[23],
      status: getHighestStatus(history.map(h => h.result)),
      history, plots
    };

  } catch(err) {
    Logger.log('ERROR: ' + err.message);
    Logger.log('Stack: ' + err.stack);
    return { found: false, error: err.message };
}
}
function testGetProfile() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const lastRow = master.getLastRow();
  
  // ดึงเลขบัตรแถวแรกที่มีข้อมูลมาทดสอบ
  if(lastRow < 2) { Logger.log('Master ว่าง'); return; }
  const firstId = master.getRange(2, 1).getValue();
  Logger.log('ทดสอบกับเลขบัตร: ' + firstId);
  
  const result = getFarmerProfile(String(firstId).replace('.0','').trim());
  Logger.log('found: ' + result.found);
  if(result.found){
    Logger.log('ชื่อ: ' + result.fname + ' ' + result.lname);
    Logger.log('ประวัติ: ' + result.history.length + ' ปี');
    Logger.log('แปลง: ' + result.plots.length + ' แปลง');
  }
}
function testWebApp() {
  try {
    const html = HtmlService.createTemplateFromFile('WebApp');
    html.page = 'dashboard';
    const output = html.evaluate();
    Logger.log('WebApp OK — length: ' + output.getContent().length);
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}
function testDashboard() {
  try {
    const result = getDashboardData();
    Logger.log(JSON.stringify(result));
  } catch(e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
  }
}
function testSpreadsheetAccess() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('ชื่อไฟล์: ' + ss.getName());
    Logger.log('Master rows: ' + ss.getSheetByName(SHEET_MASTER).getLastRow());
    Logger.log('History rows: ' + ss.getSheetByName(SHEET_HISTORY).getLastRow());
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}
function testWebAppDashboard() {
  try {
    const result = getDashboardData();
    if(result) Logger.log('OK: total=' + result.total);
    else Logger.log('NULL returned');
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}
function debugFarmerID() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const data   = master.getRange(2, 1, 5, 1).getValues();
  data.forEach((r,i) => {
    Logger.log('row '+(i+2)+': value='+r[0]+' type='+typeof r[0]+' string='+String(r[0]));
  });
}
function debugMasterColumns() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = ss.getSheetByName(SHEET_MASTER);
  const row    = master.getRange(2, 1, 1, 29).getValues()[0];
  row.forEach((v, i) => {
    Logger.log('[' + i + '] ' + typeof v + ' = ' + v);
  });
}
function getHighestStatus(historyResults) {
  const STATUS_RANK = { 'Developing': 1, 'Existing': 2, 'ต้นแบบ': 3 };
  let highest = '';
  let highestRank = 0;
  historyResults.forEach(r => {
    const rank = STATUS_RANK[r] || 0;
    if (rank > highestRank) { highestRank = rank; highest = r; }
  });
  return highest || 'Developing';
}
function openImport() {
  const html = HtmlService.createHtmlOutputFromFile('Import')
    .setTitle('นำเข้าข้อมูลจาก CSV')
    .setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}
function importCSVData(csvText, yearOverride) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const master  = ss.getSheetByName(SHEET_MASTER);
    const hist    = ss.getSheetByName(SHEET_HISTORY);
    const plot    = ss.getSheetByName(SHEET_PLOT);

    // ── CSV Parser รองรับ quoted fields ──
    function parseFullCSV(text) {
      const results = [];
      let row = [], cur = '', inQ = false;
      for (let i = 0; i < text.length; i++) {
        const c = text[i], next = text[i+1];
        if (inQ) {
          if (c === '"' && next === '"') { cur += '"'; i++; }
          else if (c === '"') { inQ = false; }
          else { cur += c; }
        } else {
          if (c === '"') { inQ = true; }
          else if (c === ',') { row.push(cur.trim()); cur = ''; }
          else if (c === '\n') {
            row.push(cur.trim()); cur = '';
            if (row.some(v => v !== '')) results.push(row);
            row = [];
          } else if (c === '\r') { /* skip */ }
          else { cur += c; }
        }
      }
      if (cur || row.length) { row.push(cur.trim()); if (row.some(v=>v!=='')) results.push(row); }
      return results;
    }

    const allRows = parseFullCSV(csvText);
    Logger.log('Total parsed rows: ' + allRows.length);

    // ── หาแถวข้อมูลจริง (แถวที่ col6 เป็นตัวเลข 13 หลัก) ──
    let dataStartIdx = -1;
    for (let i = 0; i < allRows.length; i++) {
      const col6 = String(allRows[i][6] || '').replace(/\.0+$/,'').trim();
      if (/^\d{10,13}$/.test(col6.replace(/[^0-9]/g,''))) {
        // ตรวจเพิ่มว่า col0 เป็นตัวเลข (ลำดับที่)
        const col0 = String(allRows[i][0] || '').trim();
        if (/^\d+$/.test(col0)) { dataStartIdx = i; break; }
      }
    }

    if (dataStartIdx < 0) {
      // ลองหาแบบ relaxed — col6 มีตัวเลขยาวพอ
      for (let i = 0; i < allRows.length; i++) {
        const col6 = String(allRows[i][6] || '').replace(/[^0-9]/g,'');
        if (col6.length >= 10) { dataStartIdx = i; break; }
      }
    }

    if (dataStartIdx < 0) return { ok: false, msg: 'ไม่พบแถวข้อมูลในไฟล์ — กรุณาตรวจสอบว่าเป็นไฟล์ CSV แบบฟอร์ม Smart Farmer' };

    Logger.log('Data starts at row index: ' + dataStartIdx);
    const rows = allRows.slice(dataStartIdx);
    Logger.log('Data rows: ' + rows.length);
    Logger.log('Sample col6: ' + (rows[0] ? rows[0][6] : 'N/A'));

    if (!rows.length) return { ok: false, msg: 'ไม่พบข้อมูลในไฟล์' };

    // ── ดึงข้อมูลที่มีอยู่แล้ว ──
    const existingIDs = master.getRange('A2:A').getValues()
      .flat().map(v => String(v).replace(/\.0+$/,'').trim()).filter(Boolean);

    const histData = hist.getLastRow() > 1
      ? hist.getRange(2, 1, hist.getLastRow()-1, 3).getValues() : [];
    const existingHist = new Set(
      histData.map(r => String(r[0]).replace(/\.0+$/,'').trim() + '_' + String(r[2]))
    );

    const masterNew = [], plotNew = [], histNew = [];
    const dupMaster = [], dupHist = [];
    let skipped = 0;
    const year = yearOverride || '2568';

    rows.forEach((r, idx) => {
      const pid = fixID(r[6]);
      if (!pid || pid.length < 10) {
        Logger.log('Skip row ' + idx + ': pid=[' + pid + '] col6=[' + r[6] + ']');
        skipped++; return;
      }

      const result = getResultFromRow(r);
      const bday = safe(r[7]) && safe(r[8]) && safe(r[9])
        ? safe(r[7])+'-'+safe(r[8])+'-'+safe(r[9]) : '';

      if (existingIDs.includes(pid)) {
        dupMaster.push(pid);
      } else {
        masterNew.push([
          pid, safe(r[2]), safe(r[3]), safe(r[4]), bday,
          safe(r[10]), safe(r[11]), safe(r[12]), safe(r[13]), safe(r[14]),
          safe(r[15]), safe(r[16]),
          safe(r[23]), safe(r[26]), safe(r[27]),
          safe(r[17]), safe(r[18]), safe(r[20]), fixID(r[21]),
          safe(r[28])==='1'?'มี':'', safe(r[29])==='1'?'มี':'',
          safe(r[33]), safe(r[35]), safe(r[36]), result, '', ''
        ]);
        existingIDs.push(pid);
      }

      plotNew.push([
        pid, '',
        safe(r[38]), safe(r[39]), safe(r[40]),
        safe(r[41]), safe(r[42]), safe(r[43]),
        safe(r[44]), safe(r[45]),
        safe(r[46]), safe(r[47]), safe(r[48]), safe(r[49]),
        safe(r[50]), safe(r[51]), safe(r[52]), safe(r[53]), safe(r[54])
      ]);

      const histKey = pid + '_' + year;
      if (existingHist.has(histKey)) {
        dupHist.push(pid + ' ปี ' + year);
      } else {
        histNew.push([
          pid, '', year,
          safe(r[69]), safe(r[70]), safe(r[71]), safe(r[72]),
          safe(r[73]), safe(r[74]), safe(r[75]), safe(r[76]),
          safe(r[77]), safe(r[80]), safe(r[81]),
          toMeeImport(r[82]), toMeeImport(r[83]), safe(r[84]), safe(r[85]), safe(r[86]),
          safe(r[64]), safe(r[65]), safe(r[66]), safe(r[67]), safe(r[68]),
          toUseImport(r[88]), toUseImport(r[89]), toUseImport(r[90]),
          toUseImport(r[91]), toUseImport(r[92]), safe(r[93]),
          toMeeImport(r[94]), toMeeImport(r[95]), toMeeImport(r[96]),
          toMeeImport(r[97]), toMeeImport(r[98]), toMeeImport(r[99]),
          toMeeImport(r[100]), safe(r[101]),
          safe(r[102]), result, safe(r[110]),
          safe(r[107]), safe(r[108]), safe(r[109])
        ]);
        existingHist.add(histKey);
      }
    });

    // ── เขียนเข้า Sheet ──
    if (masterNew.length) {
      const lastRow = master.getLastRow() + 1;
      master.getRange(lastRow, 1, masterNew.length, masterNew[0].length).setValues(masterNew);
      masterNew.forEach((_, i) => {
        const row = lastRow + i;
        master.getRange(row, 26).setFormula(`=IFERROR(COUNTIF(History!A:A,A${row}),0)`);
        master.getRange(row, 27).setFormula(`=IFERROR(MAXIFS(History!C:C,History!A:A,A${row}),"—")`);
      });
    }

    if (plotNew.length) {
      const lastRow = plot.getLastRow() + 1;
      plot.getRange(lastRow, 1, plotNew.length, plotNew[0].length).setValues(plotNew);
      plotNew.forEach((_, i) => {
        const row = lastRow + i;
        plot.getRange(row, 2).setFormula(
          `=IFERROR(VLOOKUP(A${row},Master!A:D,2,FALSE)&VLOOKUP(A${row},Master!A:D,3,FALSE)&" "&VLOOKUP(A${row},Master!A:D,4,FALSE),"— ไม่พบ —")`);
      });
    }

    if (histNew.length) {
      const lastRow = hist.getLastRow() + 1;
      hist.getRange(lastRow, 1, histNew.length, histNew[0].length).setValues(histNew);
      histNew.forEach((_, i) => {
        const row = lastRow + i;
        hist.getRange(row, 2).setFormula(
          `=IFERROR(VLOOKUP(A${row},Master!A:D,2,FALSE)&VLOOKUP(A${row},Master!A:D,3,FALSE)&" "&VLOOKUP(A${row},Master!A:D,4,FALSE),"— ไม่พบ —")`);
        hist.getRange(row, 23).setFormula(`=IFERROR(U${row}-T${row},"")`);
        hist.getRange(row, 24).setFormula(`=IFERROR(ROUND((U${row}-T${row})/T${row}*100,2),"—")`);
      });
    }

    return {
      ok: true,
      total:     rows.length,
      master:    masterNew.length,
      plot:      plotNew.length,
      history:   histNew.length,
      dupMaster: dupMaster.length,
      dupHist:   dupHist.length,
      skipped
    };

  } catch(e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
    return { ok: false, msg: e.message };
  }
}

// ── Helper functions สำหรับ Import ──
function parseCSVLine(line) {
  const result = []; let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
    else cur += c;
  }
  result.push(cur.trim());
  return result;
}

function fixID(val) {
  if (!val) return '';
  try { return String(Math.round(parseFloat(val))); } catch(e) {}
  return String(val).replace(/\.0+$/,'').trim();
}

function safe(val) {
  if (val === undefined || val === null) return '';
  const s = String(val).trim();
  return (s==='nan'||s==='None'||s==='NaN') ? '' : s;
}

function toMeeImport(val) {
  try { if (parseFloat(safe(val)) >= 1) return 'มี'; } catch(e) {}
  return '';
}

function toUseImport(val) {
  try { if (parseFloat(safe(val)) >= 1) return 'นำไปใช้'; } catch(e) {}
  return '';
}

function getResultFromRow(r) {
  if (safe(r[105]) && parseFloat(safe(r[105])) >= 1) return 'ต้นแบบ';
  if (safe(r[104]) && parseFloat(safe(r[104])) >= 1) return 'Existing';
  if (safe(r[103]) && parseFloat(safe(r[103])) >= 1) return 'Developing';
  return 'Developing';
}
function debugImportCSV() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const files = DriveApp.getFilesByName('Headclean-SF68-sample.csv');
  if (!files.hasNext()) { Logger.log('ไม่พบไฟล์ใน Drive'); return; }
  const content = files.next().getBlob().getDataAsString('utf-8');
  const lines = content.split(/\r?\n/);
  Logger.log('จำนวนแถวทั้งหมด: ' + (lines.length - 1));

  for (let i = 1; i <= Math.min(10, lines.length-1); i++) {
    if (!lines[i].trim()) continue;
    const r = lines[i].split(',');
    const pid = r[6] ? String(r[6]).trim() : '';
    Logger.log('แถว ' + (i+1) + ': col6=[' + pid + '] fixID=[' + fixID(pid) + ']');
  }
}
function debugRow10() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const files = DriveApp.getFilesByName('Headclean-SF68-sample.csv');
  if (!files.hasNext()) { Logger.log('ไม่พบไฟล์'); return; }
  const lines = files.next().getBlob().getDataAsString('utf-8').split(/\r?\n/);
  const r = lines[9].split(',');
  Logger.log('แถว 10 col0=' + r[0]);
  Logger.log('แถว 10 col6=' + r[6]);
  Logger.log('แถว 10 col3=' + r[3]);
}
function debugFindDataStart() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const files = DriveApp.getFilesByName('Headclean-SF68-sample.csv');
  if (!files.hasNext()) { Logger.log('ไม่พบไฟล์'); return; }
  const lines = files.next().getBlob().getDataAsString('utf-8').split(/\r?\n/);
  Logger.log('จำนวนบรรทัดทั้งหมด: ' + lines.length);
  for (let i = 0; i < Math.min(20, lines.length); i++) {
    const cols = lines[i].split(',');
    Logger.log('บรรทัด ' + (i+1) + ': col0=[' + cols[0] + '] col6=[' + cols[6] + ']');
  }
}
function convertFromDrive(filename, year) {
  try {
    // หาไฟล์ใน Drive
    const files = DriveApp.getFilesByName(filename);
    if (!files.hasNext()) {
      return { ok: false, msg: `ไม่พบไฟล์ "${filename}" ใน Google Drive` };
    }
    const file = files.next();
    const ext  = filename.split('.').pop().toLowerCase();

    let rows = [];

    if (ext === 'xlsx') {
      // แปลง xlsx → Sheets ชั่วคราว แล้วอ่านข้อมูล
      const tempSheet = SpreadsheetApp.openById(
        Drive.Files.copy({ title: '__sf_temp__' }, file.getId()).id
      );
      const ws = tempSheet.getSheets()[0];
      const data = ws.getDataRange().getValues();
      // ลบไฟล์ temp
      DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
      // หาแถวข้อมูลจริง (col0 เป็นตัวเลข)
      let start = 1;
      for (let i = 1; i < data.length; i++) {
        if (/^\d+$/.test(String(data[i][0]).trim())) { start = i; break; }
      }
      rows = data.slice(start).filter(r => {
  const v = String(r[6] || '').trim();
  // รับทั้งเลขอาราบิก เลขไทย และเลขมีขีด/ช่องว่าง
  return v !== '' && v !== '0' && /[0-9๐-๙]/.test(v);
});

    } else if (ext === 'csv') {
      const content = file.getBlob().getDataAsString('utf-8');
      rows = parseFullCSVToRows(content);
    } else {
      return { ok: false, msg: 'รองรับเฉพาะ .xlsx และ .csv' };
    }

    if (!rows.length) return { ok: false, msg: 'ไม่พบข้อมูลในไฟล์' };

    // ใช้ importCSVData ที่มีอยู่แล้ว
    return importRowData(rows, year || '2568');

  } catch(e) {
    Logger.log('ERROR convertFromDrive: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// ── parse CSV ที่มี multiline header ──
function parseFullCSVToRows(text) {
  const allRows = [];
  let row = [], cur = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], next = text[i+1];
    if (inQ) {
      if (c==='"' && next==='"') { cur+='"'; i++; }
      else if (c==='"') { inQ=false; }
      else { cur+=c; }
    } else {
      if (c==='"') { inQ=true; }
      else if (c===',') { row.push(cur.trim()); cur=''; }
      else if (c==='\n') {
        row.push(cur.trim()); cur='';
        if (row.some(v=>v!=='')) allRows.push(row);
        row=[];
      } else if (c!=='\r') { cur+=c; }
    }
  }
  if (cur||row.length) { row.push(cur.trim()); if(row.some(v=>v!=='')) allRows.push(row); }

  // หาแถวข้อมูลจริง
  let start = -1;
  for (let i=0; i<allRows.length; i++) {
    const col6 = String(allRows[i][6]||'').replace(/[^0-9]/g,'');
    if (col6.length >= 10 && /^\d+$/.test(String(allRows[i][0]||'').trim())) {
      start=i; break;
    }
  }
  if (start < 0) return [];
  return allRows.slice(start).filter(r => r[6]);
}

// ── core import logic (แยกออกมาใช้ร่วมกัน) ──
function importRowData(rows, year) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master  = ss.getSheetByName(SHEET_MASTER);
  const hist    = ss.getSheetByName(SHEET_HISTORY);
  const plot    = ss.getSheetByName(SHEET_PLOT);

  const existingIDs = master.getRange('A2:A').getValues()
    .flat().map(v => String(v).replace(/\.0+$/,'').trim()).filter(Boolean);

  const histData = hist.getLastRow() > 1
    ? hist.getRange(2,1,hist.getLastRow()-1,3).getValues() : [];
  const existingHist = new Set(
    histData.map(r => String(r[0]).replace(/\.0+$/,'').trim()+'_'+String(r[2]))
  );

  const masterNew=[], plotNew=[], histNew=[];
  const dupMaster=[], dupHist=[];
  let skipped=0;

  rows.forEach((r,idx) => {
    // padding
    while(r.length < 113) r.push('');

    const pid = fixID(r[6]);
    if (!pid || pid.length < 10) { skipped++; return; }

    const result = getResultFromRow(r);
    const bday = safe(r[7])&&safe(r[8])&&safe(r[9])
      ? safe(r[7])+'-'+safe(r[8])+'-'+safe(r[9]) : '';

    if (existingIDs.includes(pid)) {
      dupMaster.push(pid);
    } else {
      masterNew.push([
        pid, safe(r[2]), safe(r[3]), safe(r[4]), bday,
        safe(r[10]), safe(r[11]), safe(r[12]), safe(r[13]), safe(r[14]),
        safe(r[15]), safe(r[16]),
        safe(r[23]), safe(r[26]), safe(r[27]),
        safe(r[17]), safe(r[18]), safe(r[20]), fixID(r[21]),
        safe(r[28])==='1'||safe(r[28])==='1.0'?'มี':'',
        safe(r[29])==='1'||safe(r[29])==='1.0'?'มี':'',
        safe(r[33]), safe(r[35]), safe(r[36]), result, '', ''
      ]);
      existingIDs.push(pid);
    }

    plotNew.push([
      pid,'',
      safe(r[38]),safe(r[39]),safe(r[40]),
      safe(r[41]),safe(r[42]),safe(r[43]),
      safe(r[44]),safe(r[45]),
      safe(r[46]),safe(r[47]),safe(r[48]),safe(r[49]),
      safe(r[50]),safe(r[51]),safe(r[52]),safe(r[53]),safe(r[54])
    ]);

    const histKey = pid+'_'+year;
    if (existingHist.has(histKey)) {
      dupHist.push(pid);
    } else {
      histNew.push([
        pid,'',year,
        safe(r[69]),safe(r[70]),safe(r[71]),safe(r[72]),
        safe(r[73]),safe(r[74]),safe(r[75]),safe(r[76]),
        safe(r[77]),safe(r[80]),safe(r[81]),
        toMeeImport(r[82]),toMeeImport(r[83]),safe(r[84]),safe(r[85]),safe(r[86]),
        safe(r[64]),safe(r[65]),safe(r[66]),safe(r[67]),safe(r[68]),
        toUseImport(r[88]),toUseImport(r[89]),toUseImport(r[90]),
        toUseImport(r[91]),toUseImport(r[92]),safe(r[93]),
        toMeeImport(r[94]),toMeeImport(r[95]),toMeeImport(r[96]),
        toMeeImport(r[97]),toMeeImport(r[98]),toMeeImport(r[99]),
        toMeeImport(r[100]),safe(r[101]),
        safe(r[102]),result,safe(r[110]),
        safe(r[107]),safe(r[108]),safe(r[109])
      ]);
      existingHist.add(histKey);
    }
  });

  // เขียนเข้า Sheet
  if (masterNew.length) {
    const lr = master.getLastRow()+1;
    master.getRange(lr,1,masterNew.length,masterNew[0].length).setValues(masterNew);
    masterNew.forEach((_,i)=>{
      const row=lr+i;
      master.getRange(row,26).setFormula(`=IFERROR(COUNTIF(History!A:A,A${row}),0)`);
      master.getRange(row,27).setFormula(`=IFERROR(MAXIFS(History!C:C,History!A:A,A${row}),"—")`);
    });
  }
  if (plotNew.length) {
    const lr = plot.getLastRow()+1;
    plot.getRange(lr,1,plotNew.length,plotNew[0].length).setValues(plotNew);
    plotNew.forEach((_,i)=>{
      const row=lr+i;
      plot.getRange(row,2).setFormula(
        `=IFERROR(VLOOKUP(A${row},Master!A:D,2,FALSE)&VLOOKUP(A${row},Master!A:D,3,FALSE)&" "&VLOOKUP(A${row},Master!A:D,4,FALSE),"— ไม่พบ —")`);
    });
  }
  if (histNew.length) {
    const lr = hist.getLastRow()+1;
    hist.getRange(lr,1,histNew.length,histNew[0].length).setValues(histNew);
    histNew.forEach((_,i)=>{
      const row=lr+i;
      hist.getRange(row,2).setFormula(
        `=IFERROR(VLOOKUP(A${row},Master!A:D,2,FALSE)&VLOOKUP(A${row},Master!A:D,3,FALSE)&" "&VLOOKUP(A${row},Master!A:D,4,FALSE),"— ไม่พบ —")`);
      hist.getRange(row,23).setFormula(`=IFERROR(U${row}-T${row},"")`);
      hist.getRange(row,24).setFormula(`=IFERROR(ROUND((U${row}-T${row})/T${row}*100,2),"—")`);
    });
  }

  return {
    ok:true, total:rows.length,
    master:masterNew.length, plot:plotNew.length, history:histNew.length,
    dupMaster:dupMaster.length, dupHist:dupHist.length, skipped
  };
}
function debugDriveFile() {
  const filename = 'test ระบบฐานข้อมูล.xlsx'; // แก้ตรงนี้
  const files = DriveApp.getFilesByName(filename);
  if (!files.hasNext()) { Logger.log('ไม่พบไฟล์'); return; }
  
  const file = files.next();
  const tempSheet = SpreadsheetApp.openById(
    Drive.Files.copy({ title: '__sf_debug__' }, file.getId()).id
  );
  const ws = tempSheet.getSheets()[0];
  const data = ws.getDataRange().getValues();
  
  Logger.log('แถวทั้งหมดใน Sheet: ' + data.length);
  
  // หาแถวที่มีเลขบัตร
  let withID = 0, empty = 0, other = 0;
  let firstDataRow = -1;
  
  for (let i = 0; i < data.length; i++) {
    const col0 = String(data[i][0]).trim();
    const col6 = String(data[i][6]).trim();
    
    if (/^\d+$/.test(col0) && col6) {
      withID++;
      if (firstDataRow < 0) firstDataRow = i;
    } else if (!col0 && !col6) {
      empty++;
    } else {
      other++;
      if (other <= 5) Logger.log(`แถวพิเศษ ${i+1}: col0=[${col0}] col6=[${col6.substring(0,20)}]`);
    }
  }
  
  Logger.log('แถวข้อมูลจริง (มีเลขบัตร): ' + withID);
  Logger.log('แถวว่าง: ' + empty);
  Logger.log('แถวอื่นๆ (header/sub-row): ' + other);
  Logger.log('แถวข้อมูลแรกอยู่ที่ index: ' + firstDataRow);
  
  // ตรวจสอบแถวรอบ 170
  if (firstDataRow >= 0 && data.length > firstDataRow + 170) {
    const row170 = data[firstDataRow + 169];
    const row171 = data[firstDataRow + 170];
    Logger.log('แถวที่ 170: col0=['+row170[0]+'] col6=['+String(row170[6]).substring(0,15)+']');
    Logger.log('แถวที่ 171: col0=['+row171[0]+'] col6=['+String(row171[6]).substring(0,15)+']');
  }
  
  DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
}
