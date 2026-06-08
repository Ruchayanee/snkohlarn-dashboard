const SPREADSHEET_ID = '1m8OQ3gfmJlRvRG1ikXJmz0KFsqH79HmUK3__SmMLKdA';
const FINANCE_SHEET = 'การตอบแบบฟอร์ม 4';
const DEFAULT_MODEL = 'gemini-2.5-flash-lite';
const DASHBOARD_ORIGIN = 'https://ruchayanee.github.io';

function doGet(e) {
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutput(bridgeHtml_())
      .setTitle('SN Koh Larn Management AI Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return json_({ status: 'ok', service: 'SN Koh Larn Management AI' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'askManagementAI') return json_(askManagementAI_(body));
    if (body.action === 'addExpense') return json_(addExpense_(body));
    return json_({ status: 'error', message: 'Unknown action' });
  } catch (error) {
    return json_({ status: 'error', message: error.message });
  }
}

function askManagementAI(body) {
  return askManagementAI_(body || {});
}

function addExpense_(body) {
  const props = PropertiesService.getScriptProperties();
  const adminPin = props.getProperty('ADMIN_PIN');
  if (!adminPin) return { status: 'setup_required', message: 'กรุณาตั้งค่า ADMIN_PIN ใน Script Properties' };
  if (String(body.pin || '') !== adminPin) return { status: 'unauthorized', message: 'PIN ไม่ถูกต้อง' };

  const data = body.data || {};
  const dateText = formatSheetDate_(data.date || data.sheetDate);
  const note = String(data.note || '').trim();
  if (!dateText) return { status: 'error', message: 'กรุณาระบุวันที่' };
  if (!note) return { status: 'error', message: 'กรุณาระบุหมายเหตุค่าใช้จ่าย' };

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FINANCE_SHEET);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const row = headers.map(function(header) {
    if (header === 'วันที่เข้าพัก') return dateText;
    if (header === 'ราคาขายจริง') return 0;
    if (header === 'หมายเหตุ') return note;
    return '';
  });
  sheet.appendRow(row);

  return {
    status: 'success',
    message: 'บันทึกค่าใช้จ่ายแล้ว',
    date: dateText,
    note: note
  };
}

function askManagementAI_(body) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const adminPin = props.getProperty('ADMIN_PIN');
  if (!apiKey || !adminPin) return { status: 'setup_required', message: 'กรุณาตั้งค่า GEMINI_API_KEY และ ADMIN_PIN ใน Script Properties' };
  if (String(body.pin || '') !== adminPin) return { status: 'unauthorized', message: 'PIN ไม่ถูกต้อง' };
  const question = String(body.question || '').trim().slice(0, 1000);
  if (!question) return { status: 'error', message: 'กรุณาพิมพ์คำถาม' };

  const context = buildBusinessContext_();
  const prompt = [
    'คุณคือผู้ช่วยบริหารของรีสอร์ต SN Koh Larn ตอบภาษาไทย กระชับ และใช้ตัวเลขจากข้อมูลที่ได้รับเท่านั้น',
    'ห้ามแต่งข้อมูล ห้ามเปิดเผยข้อมูลส่วนตัวลูกค้า หากข้อมูลไม่พอให้บอกตรงๆ',
    'เมื่อวิเคราะห์ค่าไฟที่บิลยังไม่มา ให้ใช้ electricityEstimate และระบุว่าเป็นประมาณการจากค่าไฟต่อคืนห้องพักของเดือนก่อน ไม่ใช่ยอดบิลจริง',
    'คำแนะนำราคาต้องระบุว่าเป็นข้อเสนอเพื่อให้ผู้บริหารอนุมัติ ไม่ใช่การเปลี่ยนราคาอัตโนมัติ',
    'วันนี้: ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    'ข้อมูลสรุปธุรกิจ JSON: ' + JSON.stringify(context),
    'คำถามผู้บริหาร: ' + question
  ].join('\n\n');

  const model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 900 }
    }),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const data = JSON.parse(response.getContentText() || '{}');
  if (code < 200 || code >= 300) throw new Error((data.error && data.error.message) || 'Gemini API error ' + code);
  return { status: 'success', answer: extractGeminiText_(data), updatedAt: context.updatedAt };
}

function buildBusinessContext_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FINANCE_SHEET);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach(function(name, index) { idx[name] = index; });
  const now = new Date();
  const from = new Date(now.getTime() - 365 * 86400000);
  const monthly = {};
  const recentRooms = {};

  values.forEach(function(row) {
    const date = parseDate_(row[idx['วันที่เข้าพัก']]);
    if (!date || date < from) return;
    const key = Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM');
    const item = monthly[key] || (monthly[key] = {
      income: 0, electricity: 0, laborAndTips: 0, roomNights: 0, occupiedDates: {}
    });
    item.income += number_(row[idx['ราคาขายจริง']]);
    item.electricity += extractAmount_(row[idx['หมายเหตุ']], ['ค่าไฟ']);
    item.laborAndTips += extractAmount_(row[idx['หมายเหตุ']], ['คนงาน', 'ทิป']);
    const room = String(row[idx['เลขห้อง']] || '').trim();
    const day = Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM-dd');
    if (room) {
      item.roomNights++;
      item.occupiedDates[day] = true;
      recentRooms[room] = (recentRooms[room] || 0) + 1;
    }
  });

  Object.keys(monthly).forEach(function(key) {
    const item = monthly[key];
    item.netBeforeShare = item.income - item.electricity - item.laborAndTips;
    item.share30Percent = Math.round(item.netBeforeShare * 0.3 * 100) / 100;
    item.occupiedDayCount = Object.keys(item.occupiedDates).length;
    delete item.occupiedDates;
  });
  return {
    updatedAt: Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
    roomCount: 10,
    roomBaseRates: { Standard: 1900, Family: 2900, DeluxeTwinBed: 2900, Jacuzzi: 3500 },
    monthly: monthly,
    electricityEstimate: buildElectricityEstimate_(monthly, now),
    roomNightPopularityLast365Days: recentRooms,
    nextMonthPlanning: nextMonthPlanning_(now),
    note: 'ข้อมูลนี้ไม่มีชื่อ เบอร์โทร LINE หรือข้อมูลส่วนตัวลูกค้า'
  };
}

function buildElectricityEstimate_(monthly, now) {
  const currentMonth = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
  const previousDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const previousMonth = Utilities.formatDate(previousDate, 'Asia/Bangkok', 'yyyy-MM');
  const current = monthly[currentMonth] || { roomNights: 0, electricity: 0 };
  const previous = monthly[previousMonth] || { roomNights: 0, electricity: 0 };
  const canEstimate = previous.electricity > 0 && previous.roomNights > 0;
  const electricityPerRoomNight = canEstimate ? previous.electricity / previous.roomNights : 0;
  return {
    status: canEstimate ? 'estimated' : 'insufficient_data',
    currentMonth: currentMonth,
    previousMonth: previousMonth,
    previousMonthElectricityBill: previous.electricity,
    previousMonthRoomNights: previous.roomNights,
    currentMonthLatestBookedRoomNights: current.roomNights,
    electricityPerRoomNightFromPreviousMonth: Math.round(electricityPerRoomNight * 100) / 100,
    estimatedCurrentMonthElectricity: Math.round(electricityPerRoomNight * current.roomNights * 100) / 100,
    actualCurrentMonthElectricityRecorded: current.electricity,
    formula: 'previousMonthElectricityBill / previousMonthRoomNights * currentMonthLatestBookedRoomNights',
    note: canEstimate
      ? 'เป็นประมาณการจากคืนห้องพักที่ถูกจองล่าสุด ไม่ใช่ยอดบิลค่าไฟจริง'
      : 'ยังประมาณการไม่ได้ ต้องมีบิลค่าไฟและจำนวนคืนห้องพักของเดือนก่อน'
  };
}

function nextMonthPlanning_(now) {
  const holidays = {
    '2026-06-03': 'วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี',
    '2026-07-28': 'วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว',
    '2026-07-29': 'วันอาสาฬหบูชา',
    '2026-07-30': 'วันเข้าพรรษา',
    '2026-08-12': 'วันแม่แห่งชาติ',
    '2026-10-13': 'วันนวมินทรมหาราช',
    '2026-10-23': 'วันปิยมหาราช',
    '2026-12-05': 'วันพ่อแห่งชาติ',
    '2026-12-10': 'วันรัฐธรรมนูญ',
    '2026-12-31': 'วันสิ้นปี'
  };
  const target = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const month = Utilities.formatDate(target, 'Asia/Bangkok', 'yyyy-MM');
  const result = {};
  Object.keys(holidays).forEach(function(date) {
    if (date.indexOf(month + '-') === 0) result[date] = holidays[date];
  });
  return { month: month, thaiHolidays: result };
}

function parseDate_(value) {
  const text = String(value || '').trim();
  if (!text) return null;
  const parts = text.split('/');
  if (parts.length === 3) {
    const year = Number(parts[2]) > 2400 ? Number(parts[2]) - 543 : Number(parts[2]);
    return new Date(year, Number(parts[1]) - 1, Number(parts[0]));
  }
  const date = new Date(text);
  return isNaN(date.getTime()) ? null : date;
}

function number_(value) {
  return Number(String(value || '').replace(/,/g, '')) || 0;
}

function formatSheetDate_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return Number(iso[3]) + '/' + Number(iso[2]) + '/' + iso[1];
  return text;
}

function extractAmount_(text, words) {
  const value = String(text || '');
  return words.reduce(function(total, word) {
    const escaped = word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped + '\\s*[:=]?\\s*(\\d[\\d,.]*)', 'g');
    let match;
    while ((match = regex.exec(value)) !== null) {
      total += number_(match[1]);
    }
    return total;
  }, 0);
}

function extractGeminiText_(data) {
  const parts = [];
  (data.candidates || []).forEach(function(candidate) {
    ((candidate.content && candidate.content.parts) || []).forEach(function(part) {
      if (part.text) parts.push(part.text);
    });
  });
  return parts.join('\n').trim() || 'AI ไม่ได้ส่งคำตอบกลับมา';
}

function bridgeHtml_() {
  return '<!DOCTYPE html><html><body><script>' +
    'var allowedOrigin=' + JSON.stringify(DASHBOARD_ORIGIN) + ';' +
    'function send(type,requestId,data){window.top.postMessage({type:type,requestId:requestId,data:data},allowedOrigin);}' +
    'window.addEventListener("message",function(event){' +
      'if(event.origin!==allowedOrigin)return;' +
      'var message=event.data||{};' +
      'if(message.type==="SN_MANAGEMENT_AI_PING"){send("SN_MANAGEMENT_AI_READY");return;}' +
      'if(message.type!=="SN_MANAGEMENT_AI_REQUEST")return;' +
      'google.script.run.withSuccessHandler(function(data){send("SN_MANAGEMENT_AI_RESPONSE",message.requestId,data);})' +
      '.withFailureHandler(function(error){send("SN_MANAGEMENT_AI_RESPONSE",message.requestId,{status:"error",message:(error&&error.message)||"Apps Script error"});})' +
      '.askManagementAI(message.payload);' +
    '});' +
    'send("SN_MANAGEMENT_AI_READY");' +
  '<\/script></body></html>';
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
