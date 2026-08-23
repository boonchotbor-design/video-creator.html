// ============================================================
// webhook.gs — เวอร์ชั่นรองรับคำสั่ง "เริ่ม" และ "เริ่มคลิป"
// ============================================================

function doPost(e) {
  const res = ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  if (!e || !e.postData) return res;
  
  try {
    const body = JSON.parse(e.postData.contents);
    // --- Telegram webhook ---
    if (body.update_id !== undefined) { handleTelegramUpdate(body); return res; }
    const events = body.events || [];
    events.forEach(ev => {
      if (ev.type === 'postback') {
        handlePostback(ev);
      } else if (ev.type === 'message' && ev.message.text) {
        const txt = ev.message.text.trim().toLowerCase();
        // คำสั่งอนุมัติ/ยกเลิกจากหน้าเว็บ: "approve 12" / "อนุมัติ 12" / "reject 12"
        const cmdMatch = txt.match(/^(approve|อนุมัติ|reject|ปฏิเสธ|ยกเลิก)\s+(\d+)$/);
        if (cmdMatch) {
          handleApproveCommand(cmdMatch[1], Number(cmdMatch[2]));
        } else if (txt === 'เริ่ม' || txt === 'start') {
          startDemo();
        } else if (txt === 'เริ่มคลิป' || txt === 'reel') {
          startReelDemo();
        } else if (txt === 'status' || txt === 'สถานะ') {
          sendStatusSummary();
        }
      }
    });
  } catch (err) { console.error('doPost error:', err.message); }
  
  return res;
}

function startDemo() {
  const sheet = getSheet();
  const data = [ 'P-'+Date.now().toString().slice(-4), 'หูฟัง ProMax', 'เบสแน่น|กันน้ำ', 'https://picsum.photos/800/600.jpg', 'IMAGE', 'https://www.google.com', 'FACEBOOK' ];
  const row = sheet.getLastRow()+1;
  sheet.getRange(row, 1, 1, 7).setValues([data]);
  sheet.getRange(row, 12).setValue('PENDING');
  sendLineText('🚀 กำลังสร้างโพสต์รูปภาพตัวอย่าง...');
  processAllPending();
}

function startReelDemo() {
  const sheet = getSheet();
  const videoUrl = 'https://www.w3schools.com/html/mov_bbb.mp4'; 
  const data = [ 'R-'+Date.now().toString().slice(-4), 'คลิปทดสอบ Reels', 'วิดีโอตัวอย่างความชัดสูง', videoUrl, 'REEL', 'https://www.google.com', 'FACEBOOK' ];
  const row = sheet.getLastRow()+1;
  sheet.getRange(row, 1, 1, 7).setValues([data]);
  sheet.getRange(row, 12).setValue('PENDING');
  sendLineText('🚀 กำลังสร้างคลิป Reels ทดสอบ...');
  processAllPending();
}

// ── Telegram polling: ดึงคำสั่ง/การกดปุ่มเองทุกนาที (แก้ปัญหา GAS 302 กับ webhook) ──
// ติดตั้งครั้งเดียว: รันฟังก์ชัน setupTelegramPolling ใน Editor
function setupTelegramPolling() {
  // ยกเลิก webhook ก่อน (กันชนกับ getUpdates)
  const tgToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (tgToken) {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/deleteWebhook?drop_pending_updates=false', { muteHttpExceptions: true });
  }
  // ล้าง trigger เก่า แล้วสร้างใหม่: ทุก 1 นาที
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkTelegramUpdates') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkTelegramUpdates').timeBased().everyMinutes(1).create();
  notifyAdmin('✅ ระบบรับคำสั่ง Telegram (polling) เริ่มทำงานแล้ว — กดปุ่มอนุมัติได้เลย');
}

function checkTelegramUpdates() {
  const tgToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!tgToken) return;
  const props = PropertiesService.getScriptProperties();
  const offset = Number(props.getProperty('TG_OFFSET') || 0);
  try {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/getUpdates?timeout=0&offset=' + (offset + 1), { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (!json.ok) return;
    json.result.forEach(up => {
      props.setProperty('TG_OFFSET', String(up.update_id));
      handleTelegramUpdate(up);
    });
  } catch (e) { console.error('poll error: ' + e.message); }
}

// ── Telegram: รับการกดปุ่มอนุมัติ/ยกเลิก + คำสั่งพิมพ์ ──
function handleTelegramUpdate(body) {
  const tgToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!tgToken) return;

  // กดปุ่ม inline (callback_query)
  if (body.callback_query) {
    const cb = body.callback_query;
    const [action, rowNumStr] = String(cb.data || '').split(':');
    const rowNum = Number(rowNumStr);
    // ตอบกลับ Telegram ให้ปุ่มหมดสถานะ loading ก่อน (จะได้ไม่ค้าง)
    UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/answerCallbackQuery', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ callback_query_id: cb.id }), muteHttpExceptions: true
    });
    if ((action === 'approve' || action === 'reject') && rowNum) {
      handleApproveCommand(action, rowNum);
      // แก้ข้อความเดิมให้เห็นผลการกด
      editTelegramMessage(tgToken, cb.message,
        (action === 'approve' ? '✅ กดอนุมัติแล้ว — กำลังโพสต์ (แถว ' : '🚫 ยกเลิกแล้ว (แถว ') + rowNum + ')');
    }
    return;
  }

  // ข้อความพิมพ์: approve 12 / reject 12 / สถานะ / เริ่ม
  const txt = (body.message && body.message.text || '').trim().toLowerCase();
  const cmdMatch = txt.match(/^(approve|อนุมัติ|reject|ปฏิเสธ|ยกเลิก)\s+(\d+)$/);
  if (cmdMatch) {
    handleApproveCommand(cmdMatch[1], Number(cmdMatch[2]));
  } else if (txt === 'status' || txt === 'สถานะ') {
    sendStatusSummary();
  } else if (txt === 'เริ่ม' || txt === 'start') {
    startDemo();
  }
}

function editTelegramMessage(tgToken, msg, newText) {
  if (!msg) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/editMessageText', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ chat_id: msg.chat.id, message_id: msg.message_id, text: newText }),
    muteHttpExceptions: true
  });
}

// ── อนุมัติ/ยกเลิกตามเลขแถวใน Sheet (ใช้กับงานจากหน้าเว็บ) ──
function handleApproveCommand(action, rowNum) {
  const sheet = getSheet();
  if (rowNum < 2 || rowNum > sheet.getLastRow()) {
    sendLineText('❌ ไม่พบแถวที่ ' + rowNum + ' ใน Sheet');
    return;
  }
  const status = sheet.getRange(rowNum, 12).getValue();
  const productName = sheet.getRange(rowNum, 2).getValue();
  if (action === 'approve' || action === 'อนุมัติ') {
    if (status !== 'PENDING_REVIEW') {
      sendLineText('⚠️ แถวที่ ' + rowNum + ' สถานะเป็น ' + status + ' (ไม่ใช่ PENDING_REVIEW) อนุมัติไม่ได้');
      return;
    }
    sheet.getRange(rowNum, 12).setValue('APPROVED');
    sendLineText('✅ อนุมัติแล้ว: ' + productName + ' — เริ่มโพสต์เลย!');
    postApprovedProduct(rowNum);
  } else {
    sheet.getRange(rowNum, 12).setValue('REJECTED');
    sendLineText('🚫 ยกเลิกแล้ว: ' + productName);
  }
}

// ── ฟังก์ชันอื่นๆ ยังคงเดิม ──
function handlePostback(event) {
  const data = parseQueryString(event.postback.data);
  const rowNum = findRowByProductId(data.product_id);
  if (!rowNum) return;
  if (data.action === 'approve') {
    getSheet().getRange(rowNum, 12).setValue('APPROVED');
    postApprovedProduct(rowNum);
  } else {
    getSheet().getRange(rowNum, 12).setValue('REJECTED');
    sendLineText('🚫 Rejected แล้ว');
  }
}

function findRowByProductId(productId) {
  const ids = getSheet().getRange(2, 1, getSheet().getLastRow()-1, 1).getValues();
  for (let i=0; i<ids.length; i++) { if (String(ids[i][0]) === String(productId)) return i+2; }
  return null;
}

function parseQueryString(str) {
  const res = {};
  str.split('&').forEach(p => { const [k,v] = p.split('='); res[k] = v; });
  return res;
}

function sendStatusSummary() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { sendLineText('ไม่มีข้อมูลครับ'); return; }
  const stats = sheet.getRange(2, 12, lastRow-1, 1).getValues();
  const counts = { PENDING:0, AI_DONE:0, PENDING_REVIEW:0, APPROVED:0, REJECTED:0, POSTED:0, ERROR:0 };
  stats.forEach(s => { if (counts[s[0]] !== undefined) counts[s[0]]++; });
  const msg = `📊 สรุปสถานะ\n⏳ PENDING: ${counts.PENDING}\n🚀 Posted: ${counts.POSTED}\n❌ Rejected: ${counts.REJECTED}\n🔴 Error: ${counts.ERROR}`;
  sendLineText(msg);
}
