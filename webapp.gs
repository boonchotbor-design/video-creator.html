// ============================================================
// webapp.gs — Video Creator Web App (UI สั่งสร้างวิดีโอขายสินค้า)
// Deploy: Extensions > Apps Script > Deploy > Web app
// ============================================================

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

  // ตัด ```json ที่ Gemini ชอบใส่ออก
  const jsonText = raw.replace(/```json|```/g, '').trim();
  let result;
  try {
    result = JSON.parse(jsonText);
  } catch (e) {
    throw new Error('AI ตอบไม่ใช่ JSON: ' + raw.substring(0, 200));
  }
  return { ok: true, story: result };
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
  sheet.getRange(rowNum, CONFIG.COL.STATUS).setValue(
    p.requireApproval ? CONFIG.STATUS.PENDING_REVIEW : CONFIG.STATUS.PENDING
  );

  // ถ้าเปิดอนุมัติผ่าน LINE ให้ส่งแจ้งเตือนทันที
  if (p.requireApproval) {
    try {
      sendLineText('🎬 มีคลิปใหม่รออนุมัติ\n📦 ' + (p.productName || '') +
        '\n🔑 คีย์เวิร์ด: ' + (p.keyword || '') +
        '\nพิมพ์ APPROVE ' + rowNum + ' เพื่ออนุมัติโพสต์');
    } catch (e) { /* ไม่ให้ LINE fail กระทบการบันทึก */ }
  }

  return { ok: true, row: rowNum };
}

// ── ตัวช่วยเรียก Gemini (ใช้ Smart Quota เหมือน ai_handler) ──
function callGemini(prompt) {
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.9, maxOutputTokens: 2048, responseMimeType: 'application/json' }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const urlList = `https://generativelanguage.googleapis.com/v1beta/models?key=${CONFIG.GEMINI_API_KEY}`;
  const resList = UrlFetchApp.fetch(urlList, { muteHttpExceptions: true });
  if (resList.getResponseCode() !== 200) {
    throw new Error('ดึงรายชื่อโมเดลไม่ได้: ' + resList.getContentText());
  }
  const models = JSON.parse(resList.getContentText()).models
    .filter(m => m.supportedGenerationMethods.includes('generateContent'))
    .map(m => m.name)
    .sort((a, b) => (a.includes('1.5-flash') ? -1 : b.includes('1.5-flash') ? 1 : 0));

  let lastError = '';
  for (let i = 0; i < Math.min(models.length, 10); i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/${models[i]}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text.trim();
      }
      lastError = res.getContentText();
      Utilities.sleep(500);
    } catch (e) { lastError = e.message; }
  }
  throw new Error('Gemini ทุกโมเดล Quota เต็ม: ' + String(lastError).substring(0, 150));
}
