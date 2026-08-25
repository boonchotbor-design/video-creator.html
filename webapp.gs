// ============================================================
// webapp.gs — Video Creator Web App (UI สั่งสร้างวิดีโอขายสินค้า)
// ไฟล์นี้ทำงานได้แบบสแตนด์อโลน: ถ้าไม่มี config.gs ก็จะสร้างค่าเอง
// ============================================================

// ---- กันไฟล์ขาด: ถ้า config.gs ไม่มีในโปรเจกต์ ให้ใช้ค่าเริ่มต้นตรงนี้ ----
if (typeof CONFIG === 'undefined') {
  globalThis.CONFIG = {
    SHEET_ID: '15tjRtRJRx7owdZf6yEwRUiOIQWUnCkt0vd7BHjWR2i4',
    SHEET_NAME: 'Control',
    GEMINI_API_KEY: (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''),
    COL: {
      PRODUCT_ID: 1, PRODUCT_NAME: 2, KEY_FEATURES: 3, MEDIA_URL: 4, MEDIA_TYPE: 5,
      AFFILIATE_LINK: 6, TARGET_PLATFORM: 7, SCHEDULED_DATE: 8, EXTRA_NOTES: 9,
      AI_CAPTION: 10, AI_GENERATED_AT: 11, STATUS: 12, LINE_MSG_ID: 13,
      FB_POST_ID: 14, POSTED_AT: 16, ERROR_LOG: 17
    },
    STATUS: {
      PENDING: 'PENDING', AI_DONE: 'AI_DONE', PENDING_REVIEW: 'PENDING_REVIEW',
      APPROVED: 'APPROVED', REJECTED: 'REJECTED', POSTED: 'POSTED', ERROR: 'ERROR'
    }
  };
}
if (typeof getSheet === 'undefined') {
  globalThis.getSheet = function () {
    return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  };
}

// ---- Facebook posting (สแตนด์อโลน + รองรับหลายเพจ) ----
// เพจหลัก: FB_PAGE_ID / FB_PAGE_TOKEN, เพจเพิ่มเติม: FB_PAGE_ID_2 / FB_PAGE_TOKEN_2, _3, _4...
function _fbPages() {
  const props = PropertiesService.getScriptProperties();
  const pages = [];
  const primaryId = props.getProperty('FB_PAGE_ID') || CONFIG.FB_PAGE_ID;
  const primaryToken = String(props.getProperty('FB_PAGE_TOKEN') || CONFIG.FB_PAGE_TOKEN || '').replace(/\s+/g, '');
  if (primaryId && primaryToken) pages.push({ id: primaryId, token: primaryToken });
  for (let n = 2; n <= 5; n++) {
    const id = props.getProperty('FB_PAGE_ID_' + n);
    const token = String(props.getProperty('FB_PAGE_TOKEN_' + n) || '').replace(/\s+/g, '');
    if (id && token) pages.push({ id: id, token: token });
  }
  return pages;
}

function _fbToken() {
  const pages = _fbPages();
  return pages.length ? pages[0].token : '';
}

if (typeof callFacebookAPI === 'undefined') {
  globalThis.callFacebookAPI = function (url, payload) {
    const token = _fbToken();
    if (!token) throw new Error('ยังไม่ได้ใส่ FB_PAGE_TOKEN — ใส่ใน Script Properties หรือผ่านเมนู 🚀 ใน Google Sheet');
    const options = { method: 'post', payload: payload, muteHttpExceptions: true };
    const res = UrlFetchApp.fetch(`${url}?access_token=${token}`, options);
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) throw new Error(json.error && json.error.message || res.getContentText());
    return json.id || json.post_id || json.video_id;
  };
}

if (typeof postFacebookReel === 'undefined') {
  globalThis.postFacebookReel = function (caption, videoUrl) {
    if (!videoUrl) throw new Error('ไม่มีลิงก์วีดีโอในคอลัมน์ D (Media URL) — ใส่ลิงก์ไฟล์จริงก่อนโพสต์');
    const token = _fbToken();
    const pageId = PropertiesService.getScriptProperties().getProperty('FB_PAGE_ID') || CONFIG.FB_PAGE_ID;
    const ver = CONFIG.FB_API_VERSION || 'v21.0';
    const baseUrl = `https://graph.facebook.com/${ver}/${pageId}/video_reels`;
    try {
      // 1) ลองโพสต์เป็น Reel ก่อน
      const initRes = UrlFetchApp.fetch(baseUrl, { method: 'post', payload: { upload_phase: 'start', access_token: token }, muteHttpExceptions: true });
      const videoId = JSON.parse(initRes.getContentText()).video_id;
      if (!videoId) throw new Error('เริ่ม Reel ไม่ได้');
      const pubRes = UrlFetchApp.fetch(baseUrl, {
        method: 'post', muteHttpExceptions: true,
        payload: {
          upload_phase: 'finish', video_id: videoId, video_state: 'PUBLISHED',
          description: caption, file_url: videoUrl, video_url: videoUrl, access_token: token
        }
      });
      const pub = JSON.parse(pubRes.getContentText());
      if (pub.error) throw new Error(pub.error.message);
      return pub.id || videoId;
    } catch (e) {
      // 2) Reel ไม่ผ่าน → สลับไปโพสต์วีดีโอปกติ (ผ่านแน่นอนกว่า)
      console.warn('Reel ไม่สำเร็จ (' + e.message + ') — สลับไปโพสต์วีดีโอปกติ');
      return globalThis.postFacebookVideo(caption, videoUrl, '');
    }
  };
}

if (typeof postFacebookImage === 'undefined') {
  globalThis.postFacebookImage = function (caption, imageUrl, affiliateLink) {
    const pageId = PropertiesService.getScriptProperties().getProperty('FB_PAGE_ID') || CONFIG.FB_PAGE_ID;
    return globalThis.callFacebookAPI(
      `https://graph.facebook.com/${CONFIG.FB_API_VERSION || 'v21.0'}/${pageId}/photos`,
      { url: imageUrl, caption: `${caption}\n\n🛒 สั่งซื้อ: ${affiliateLink || ''}` });
  };
}

if (typeof postFacebookVideo === 'undefined') {
  globalThis.postFacebookVideo = function (caption, videoUrl, affiliateLink) {
    const pageId = PropertiesService.getScriptProperties().getProperty('FB_PAGE_ID') || CONFIG.FB_PAGE_ID;
    return globalThis.callFacebookAPI(
      `https://graph.facebook.com/${CONFIG.FB_API_VERSION || 'v21.0'}/${pageId}/videos`,
      { file_url: videoUrl, description: `${caption}\n\n🛒 สั่งซื้อ: ${affiliateLink || ''}` });
  };
}

// โพสต์วีดีโอไปยัง 1 เพจ: ลอง Reel ก่อน ถ้าพลาดสลับไปโพสต์วีดีโอปกติ
function _postVideoToPage(page, caption, videoUrl, affiliateLink) {
  const ver = CONFIG.FB_API_VERSION || 'v21.0';
  // 1) ลองโพสต์วีดีโอปกติ (เสถียรสุด ทดสอบแล้วผ่านกับไฟล์ URL ภายนอก)
  try {
    const res = UrlFetchApp.fetch(`https://graph.facebook.com/${ver}/${page.id}/videos?access_token=${page.token}`, {
      method: 'post', muteHttpExceptions: true,
      payload: { file_url: videoUrl, description: `${caption}\n\n🛒 สั่งซื้อ: ${affiliateLink || ''}` }
    });
    const json = JSON.parse(res.getContentText());
    if (json.id) return json.id;
  } catch (e) { /* ไปลอง Reel */ }
  // 2) ลอง Reel
  const baseUrl = `https://graph.facebook.com/${ver}/${page.id}/video_reels`;
  const initRes = UrlFetchApp.fetch(baseUrl, { method: 'post', payload: { upload_phase: 'start', access_token: page.token }, muteHttpExceptions: true });
  const videoId = JSON.parse(initRes.getContentText()).video_id;
  if (videoId) {
    const pubRes = UrlFetchApp.fetch(baseUrl, {
      method: 'post', muteHttpExceptions: true,
      payload: {
        upload_phase: 'finish', video_id: videoId, video_state: 'PUBLISHED',
        description: caption, file_url: videoUrl, video_url: videoUrl, access_token: page.token
      }
    });
    const pub = JSON.parse(pubRes.getContentText());
    if (pub.id) return pub.id;
    throw new Error(pub.error ? pub.error.message : 'Reel และวีดีโอปกติไม่สำเร็จ');
  }
  throw new Error('เริ่ม Reel ไม่ได้ — ตรวจสอบ token/สิทธิ์ของเพจ ' + page.id);
}

if (typeof postApprovedProduct === 'undefined') {
  globalThis.postApprovedProduct = function (rowNum) {
    const sheet = getSheet();
    const row = sheet.getRange(rowNum, 1, 1, 17).getValues()[0];
    const caption = row[9];
    const mediaUrl = row[3];
    const affiliateLink = row[5];
    const productName = row[1];

    try {
      if (!mediaUrl) throw new Error('ไม่มีลิงก์วีดีโอในคอลัมน์ D (Media URL) — ใส่ลิงก์ไฟล์จริงก่อนโพสต์');
      const pages = _fbPages();
      if (!pages.length) throw new Error('ยังไม่ได้ใส่ FB_PAGE_TOKEN — ใส่ใน Script Properties');

      // โพสต์ทุกเพจที่ตั้งค่าไว้ (หลัก + เพจ 2, 3, 4, 5)
      const results = [];
      const errors = [];
      pages.forEach(pg => {
        try {
          const pid = _postVideoToPage(pg, caption, mediaUrl, affiliateLink);
          results.push({ page: pg.id, postId: pid });
        } catch (err) { errors.push('เพจ ' + pg.id + ': ' + err.message); }
      });
      if (!results.length) throw new Error(errors.join(' | ') || 'โพสต์ไม่สำเร็จทุกเพจ');

      sheet.getRange(rowNum, 14).setValue(results.map(r => r.postId).join(', '));
      sheet.getRange(rowNum, 16).setValue(new Date().toISOString());
      sheet.getRange(rowNum, 12).setValue('POSTED');
      let msg = `✅ โพสต์สำเร็จ!\n📦 ${productName}\n📄 แถว ${rowNum}\n`;
      results.forEach(r => { msg += `📘 https://facebook.com/${r.postId}\n`; });
      if (errors.length) msg += '⚠️ บางเพจไม่สำเร็จ: ' + errors.join(' | ');
      notifyAdmin(msg);
    } catch (e) {
      sheet.getRange(rowNum, 12).setValue('ERROR');
      sheet.getRange(rowNum, 17).setValue(e.message);
      notifyAdmin(`❌ โพสต์ไม่สำเร็จ: ${e.message}`);
    }
  };
}

/**
 * จุดเข้าหน้าเว็บ — คัดลอกไฟล์ ui/video-creator.html ไปเป็นไฟล์ HTML
 * ชื่อ "VideoCreator" ในโปรเจกต์ Apps Script ก่อน Deploy
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('VideoCreator')
    .setTitle('AZ Creator — สั่งสร้างวิดีโอขายสินค้า')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 1) ดึงรายการสินค้าจาก Sheet มาแสดงใน dropdown ────────────
function getProducts() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, products: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL.AFFILIATE_LINK).getValues();
  const products = data
    .filter(r => r[CONFIG.COL.PRODUCT_ID - 1] || r[CONFIG.COL.PRODUCT_NAME - 1])
    .map(r => ({
      productId: String(r[CONFIG.COL.PRODUCT_ID - 1]),
      productName: String(r[CONFIG.COL.PRODUCT_NAME - 1]),
      features: String(r[CONFIG.COL.KEY_FEATURES - 1] || ''),
      affiliateLink: String(r[CONFIG.COL.AFFILIATE_LINK - 1] || '')
    }));
  return { ok: true, products: products };
}

// ── 2) AI สร้างเรื่อง + บท 3 ซีน + แคปชัน (เรียก Gemini) ────
function generateStory(payload) {
  const p = payload || {};
  const prompt = `คุณคือนักเขียนบทวิดีโอสั้นขายสินค้าสำหรับ TikTok/Facebook ชาวไทย

จงสร้างบทวิดีโอสั้นโดยใช้ข้อมูลนี้:
- สไตล์วิดีโอ: ${p.style || 'Pixar Story'}
- แนวเรื่อง: ${p.theme || 'สิ่งของปากแจ๋ว (ของใช้จริง)'}
- ความยาว: ${p.duration || '24 วินาที'}
- ชื่อสินค้า: ${p.productName || ''}
- จุดขายสินค้า: ${p.features || ''}
- คีย์เวิร์ดปัญหาที่คนเจอ: ${p.keyword || ''}
- เสียงบรรยาย: ${p.voiceGender || 'ผู้ชาย'} โทน${p.voiceTone || 'หยาบดิบ'} ความยาวบท${p.scriptLength || 'กลาง'}

ตอบกลับเป็น JSON เท่านั้น ห้ามมีอย่างอื่น รูปแบบ:
{
  "topic": "หัวข้อเรื่อง 1 ประโยค",
  "characters": "รายละเอียดตัวละคร",
  "setting": "รายละเอียดฉาก",
  "scenes": [
    {"title": "Hook (0-6 วิ)", "image": "คำอธิบายภาพที่จะให้ AI วาด", "script": "บทพูดซีน 1"},
    {"title": "Solution (6-18 วิ)", "image": "...", "script": "..."},
    {"title": "CTA (18-24 วิ)", "image": "...", "script": "..."}
  ],
  "caption": "แคปชันโพสต์พร้อม hashtag",
  "hashtags": "#tag1 #tag2 #tag3"
}`;

  const raw = callGemini(prompt);
  let result = extractJson(raw);
  if (!result) {
    // ครั้งแรกตอบมาไม่ใช่ JSON — สั่งซ้ำครั้งเดียวโดยบอกให้ตอบ JSON เท่านั้น
    const raw2 = callGemini(prompt + '\n\nข้อความก่อนหน้าของคุณไม่ใช่ JSON — ตอบใหม่เป็น JSON ล้วน ห้ามมีข้อความนำ ห้ามมี ``` เริ่มต้นที่ { และปิดที่ } เท่านั้น');
    result = extractJson(raw2);
  }
  if (!result) {
    throw new Error('AI ตอบไม่ใช่ JSON (ลองกดใหม่อีกครั้ง): ' + raw.substring(0, 150));
  }
  return { ok: true, story: result };
}

// ── ดึง JSON จากคำตอบ AI แม่นขึ้น: ตัด ``` และหาวงเล็บปีกกา { ... } ──
function extractJson(text) {
  if (!text) return null;
  let t = String(text).replace(/```json|```/g, '').trim();
  try { return JSON.parse(t); } catch (e) {}
  const start = t.indexOf('{');
  const end = t.lastIndexOf('}');
  if (start !== -1 && end > start) {
    try { return JSON.parse(t.substring(start, end + 1)); } catch (e) {}
  }
  return null;
}

// ── 3) บันทึกงานสร้างวิดีโอลง Sheet (รอการสร้าง/อนุมัติ) ─────
function createVideoJob(payload) {
  const p = payload || {};
  const sheet = getSheet();
  const rowNum = sheet.getLastRow() + 1;
  const now = new Date().toISOString();

  sheet.getRange(rowNum, CONFIG.COL.PRODUCT_ID).setValue(p.productId || ('JOB-' + Date.now()));
  sheet.getRange(rowNum, CONFIG.COL.PRODUCT_NAME).setValue(p.productName || '');
  sheet.getRange(rowNum, CONFIG.COL.KEY_FEATURES).setValue(p.features || '');
  sheet.getRange(rowNum, CONFIG.COL.MEDIA_TYPE).setValue('REEL');
  sheet.getRange(rowNum, CONFIG.COL.AFFILIATE_LINK).setValue(p.affiliateLink || '');
  sheet.getRange(rowNum, CONFIG.COL.TARGET_PLATFORM).setValue(p.platforms || 'TikTok,Facebook');
  sheet.getRange(rowNum, CONFIG.COL.EXTRA_NOTES).setValue(
    'STYLE:' + (p.style || '') +
    '|KEYWORD:' + (p.keyword || '') +
    '|VOICE:' + (p.voiceGender || '') + '/' + (p.voiceTone || '') +
    '|LOOP:' + (p.loopCount || 1) +
    '|TOPIC:' + (p.topic || '')
  );
  sheet.getRange(rowNum, CONFIG.COL.AI_CAPTION).setValue(p.caption || '');
  sheet.getRange(rowNum, CONFIG.COL.AI_GENERATED_AT).setValue(now);
  sheet.getRange(rowNum, CONFIG.COL.STATUS).setValue(CONFIG.STATUS.PENDING_REVIEW);

  // แจ้งเตือนทุกงานที่สร้างใหม่เสมอ (ถ้ารออนุมัติจะมีปุ่มกดใน Telegram)
  try {
    notifyNewJob(p, rowNum);
  } catch (e) { console.error('แจ้งเตือนไม่สำเร็จ: ' + e.message); }

  return { ok: true, row: rowNum };
}

// ── ตัวช่วยเรียก Gemini (ใช้ Smart Quota เหมือน ai_handler) ──
// ── แจ้งเตือนแอดมิน: Telegram (แนบวีดีโอ/รูป + ปุ่มอนุมัติ) + LINE ──
function notifyAdmin(msg, approveRow, mediaUrl) {
  // --- Telegram ---
  const tgToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  const tgChat = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  if (tgToken && tgChat) {
    try {
      const buttons = approveRow ? { inline_keyboard: [[
        { text: '✅ อนุมัติโพสต์', callback_data: 'approve:' + approveRow },
        { text: '❌ ยกเลิก', callback_data: 'reject:' + approveRow }
      ]] } : null;

      const url = (mediaUrl || '').trim();
      const isPhoto = /\.(jpe?g|png|webp)(\?|$)/i.test(url);
      if (url) {
        // แนบไฟล์จริง: รูป → sendPhoto, อื่นๆ → sendVideo พร้อม caption + ปุ่ม
        const method = isPhoto ? 'sendPhoto' : 'sendVideo';
        const payload = { chat_id: tgChat, caption: msg.substring(0, 1000) };
        payload[isPhoto ? 'photo' : 'video'] = url;
        if (buttons) payload.reply_markup = buttons;
        UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/' + method, {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify(payload), muteHttpExceptions: true
        });
      } else {
        // ไม่มีไฟล์จริง → ส่งข้อความ + ปุ่ม
        const payload = { chat_id: tgChat, text: msg };
        if (buttons) payload.reply_markup = buttons;
        UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/sendMessage', {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify(payload), muteHttpExceptions: true
        });
      }
    } catch (e) { console.error('Telegram send failed: ' + e.message); }
  } else {
    console.warn('ยังไม่ได้ตั้งค่า TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID ใน Script Properties');
  }

  // --- LINE ---
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN');
  const userId = PropertiesService.getScriptProperties().getProperty('LINE_USER_ID');
  if (token && userId) {
    try {
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: msg }] }),
        muteHttpExceptions: true
      });
    } catch (e) { console.error('LINE push failed: ' + e.message); }
  }
}

// ส่งตัวเลขแถวที่สร้าง พร้อมแจ้งเตือนทุกครั้งที่มีงานใหม่ (แนบวีดีโอ + ปุ่มอนุมัติเสมอ)
function notifyNewJob(p, rowNum) {
  const msg = '🎬 มีคลิปใหม่รออนุมัติ' +
    '\n📦 ' + (p.productName || '') +
    '\n🔑 คีย์เวิร์ด: ' + (p.keyword || '') +
    '\n📊 สถานะ: PENDING_REVIEW (รออนุมัติ)' +
    '\n📄 แถวที่: ' + rowNum +
    (p.mediaUrl ? '' : '\n🎥 (แนบวีดีโอตัวอย่าง — ยังไม่มีไฟล์จริง)');
  // ถ้าไม่มีไฟล์จริง ใช้วีดีโอตัวอย่างแนบไปก่อน / ปุ่มอนุมัติส่งเสมอ
  notifyAdmin(msg, rowNum, p.mediaUrl || 'https://www.w3schools.com/html/mov_bbb.mp4');
}

function callGemini(prompt) {
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.9, maxOutputTokens: 2048 }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // โมเดลที่ทดสอบกับคีย์นี้แล้วยังใช้ได้จริง (โมเดล 1.5/2.0/2.5 ถูกปิดหมดแล้ว)
  const preferred = [
    'models/gemini-3.6-flash',
    'models/gemini-3-flash-preview',
    'models/gemini-3.1-flash-lite',
    'models/gemini-flash-latest',
    'models/gemini-flash-lite-latest'
  ];

  // เศษวลีจากพรอมป์ ใช้เช็คว่า AI ตอบพรอมป์เรากลับมา (echo) หรือไม่
  const echoMark = prompt.replace(/\s+/g, ' ').trim().substring(20, 60);

  let lastError = '';
  for (let i = 0; i < preferred.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/${preferred[i]}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        const parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
        const text = (parts || []).map(p => p.text || '').join('').trim();
        if (text && text.length > 40 && text.indexOf(echoMark) === -1) {
          return text;
        }
        console.warn(`โมเดล ${preferred[i]} ตอบผิดปกติ (echo/ว่าง) ข้ามไปตัวถัดไป`);
        lastError = 'AI ตอบแบบ echo/ว่างเปล่า';
      } else {
        lastError = res.getContentText();
      }
      Utilities.sleep(400);
    } catch (e) { lastError = e.message; }
  }
  throw new Error('Gemini ตอบไม่สำเร็จทุกโมเดล: ' + String(lastError).substring(0, 150));
}

// ============================================================
// โหมดครบวงจรในหน้าเว็บ: สร้างภาพ AI + อัปโหลดไฟล์ + โพสต์ทันที
// ============================================================

// สร้างภาพประกอบซีนด้วย AI → บันทึก Drive → คืนลิงก์สาธารณะ
function generateSceneImage(imagePrompt) {
  const payload = {
    contents: [{ parts: [{ text: 'Create a vertical 9:16 cinematic illustration for a short product video. Style: colorful 3D animation like Pixar. Scene: ' + imagePrompt + '. No text in image.' }] }],
    generationConfig: { responseModalities: ['TEXT', 'IMAGE'] }
  };
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-pro-image:generateContent?key=' + CONFIG.GEMINI_API_KEY,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('สร้างภาพไม่สำเร็จ: ' + res.getContentText().substring(0, 150));
  const parts = JSON.parse(res.getContentText()).candidates[0].content.parts;
  const imgPart = (parts || []).find(function (p) { return p.inlineData || p.inline_data; });
  if (!imgPart) throw new Error('AI ไม่ได้ส่งภาพกลับมา ลองกดใหม่');
  const inline = imgPart.inlineData || imgPart.inline_data;
  const blob = Utilities.newBlob(Utilities.base64Decode(inline.data), inline.mimeType || 'image/png');
  return saveToDrive(blob, 'scene_' + Date.now() + '.png');
}

// อัปโหลดไฟล์ (วีดีโอ/รูป) จากหน้าเว็บ → Drive → ลิงก์สาธารณะ
function uploadMediaFile(base64Data, mimeType, fileName) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName || ('upload_' + Date.now()));
  return saveToDrive(blob, fileName || ('upload_' + Date.now() + '.mp4'));
}

function saveToDrive(blob, name) {
  let folder;
  const it = DriveApp.getFoldersByName('VideoCreator Media');
  folder = it.hasNext() ? it.next() : DriveApp.createFolder('VideoCreator Media');
  const file = folder.createFile(blob);
  file.setName(name);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/uc?export=download&id=' + file.getId(), fileId: file.getId() };
}

// โพสต์ทันทีจากหน้าเว็บ (ระบบบันทึก Sheet เอง ผู้ใช้ไม่ต้องแตะ)
function postNow(payload) {
  const p = payload || {};
  if (!p.mediaUrl) return { ok: false, error: 'ยังไม่มีไฟล์วีดีโอ — กด "อัปโหลดคลิป" หรือ "สร้างภาพทุกซีน" ก่อน' };

  const sheet = getSheet();
  const rowNum = sheet.getLastRow() + 1;
  sheet.getRange(rowNum, CONFIG.COL.PRODUCT_ID).setValue(p.productId || ('DIRECT-' + Date.now()));
  sheet.getRange(rowNum, CONFIG.COL.PRODUCT_NAME).setValue(p.productName || '');
  sheet.getRange(rowNum, CONFIG.COL.MEDIA_URL).setValue(p.mediaUrl);
  sheet.getRange(rowNum, CONFIG.COL.MEDIA_TYPE).setValue(p.mimeType && p.mimeType.indexOf('image') === 0 ? 'IMAGE' : 'REEL');
  sheet.getRange(rowNum, CONFIG.COL.AFFILIATE_LINK).setValue(p.affiliateLink || '');
  sheet.getRange(rowNum, CONFIG.COL.AI_CAPTION).setValue(p.caption || '');
  sheet.getRange(rowNum, CONFIG.COL.STATUS).setValue('POSTING');

  try {
    const pages = _fbPages();
    if (!pages.length) throw new Error('ยังไม่ได้ใส่ FB_PAGE_TOKEN ใน Script Properties');
    const results = [];
    const errors = [];
    pages.forEach(function (pg) {
      try {
        const pid = _postVideoToPage(pg, p.caption || '', p.mediaUrl, p.affiliateLink || '');
        results.push('📘 https://facebook.com/' + pid);
      } catch (err) { errors.push('เพจ ' + pg.id + ': ' + err.message); }
    });
    if (!results.length) throw new Error(errors.join(' | ') || 'โพสต์ไม่สำเร็จ');

    sheet.getRange(rowNum, CONFIG.COL.FB_POST_ID).setValue(results.join(', '));
    sheet.getRange(rowNum, CONFIG.COL.POSTED_AT).setValue(new Date().toISOString());
    sheet.getRange(rowNum, CONFIG.COL.STATUS).setValue('POSTED');
    try { notifyAdmin('🚀 โพสต์ทันทีสำเร็จ\n📦 ' + (p.productName || '') + '\n' + results.join('\n')); } catch (e) {}
    return { ok: true, row: rowNum, links: results, errors: errors };
  } catch (e) {
    sheet.getRange(rowNum, CONFIG.COL.STATUS).setValue('ERROR');
    sheet.getRange(rowNum, CONFIG.COL.ERROR_LOG).setValue(e.message);
    return { ok: false, error: e.message };
  }
}
