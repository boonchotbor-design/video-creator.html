// ============================================================
// fb_handler.gs — เวอร์ชั่นสมบูรณ์ (Reels 2-Step + Auto-Clean)
// ============================================================

function postFacebookImage(caption, imageUrl, affiliateLink) {
  const payload = { url: imageUrl, caption: `${caption}\n\n🛒 สั่งซื้อ: ${affiliateLink}` };
  return callFacebookAPI(`https://graph.facebook.com/${CONFIG.FB_API_VERSION}/${CONFIG.FB_PAGE_ID}/photos`, payload);
}

function postFacebookVideo(caption, videoUrl, affiliateLink) {
  const payload = { file_url: videoUrl, description: `${caption}\n\n🛒 สั่งซื้อ: ${affiliateLink}` };
  return callFacebookAPI(`https://graph.facebook.com/${CONFIG.FB_API_VERSION}/${CONFIG.FB_PAGE_ID}/videos`, payload);
}

function postFacebookReel(caption, videoUrl) {
  // 1) โพสต์วีดีโอปกติก่อน (เสถียรกว่า — ทดสอบกับไฟล์ URL ภายนอกแล้วผ่าน)
  try {
    const pageId = PropertiesService.getScriptProperties().getProperty('FB_PAGE_ID') || CONFIG.FB_PAGE_ID;
    const vToken = String(_fbToken ? _fbToken() : CONFIG.FB_PAGE_TOKEN).replace(/\s+/g, '');
    const vRes = UrlFetchApp.fetch(`https://graph.facebook.com/${CONFIG.FB_API_VERSION}/${pageId}/videos?access_token=${vToken}`, {
      method: 'post', muteHttpExceptions: true,
      payload: { file_url: videoUrl, description: caption }
    });
    const vJson = JSON.parse(vRes.getContentText());
    if (vJson.id) return vJson.id;
  } catch (e) { /* ไปลอง Reel */ }

  // 2) ถ้าวีดีโอปกติไม่ผ่าน ลอง Reel 2-step
  const token = String(CONFIG.FB_PAGE_TOKEN).replace(/\s+/g, '');
  const baseUrl = `https://graph.facebook.com/${CONFIG.FB_API_VERSION}/${CONFIG.FB_PAGE_ID}/video_reels`;

  try {
    const initPayload = { upload_phase: 'start', access_token: token };
    const initRes = UrlFetchApp.fetch(baseUrl, { method: 'post', payload: initPayload, muteHttpExceptions: true });
    const videoId = JSON.parse(initRes.getContentText()).video_id;

    const publishPayload = {
      upload_phase: 'finish',
      video_id: videoId,
      video_state: 'PUBLISHED',
      description: caption,
      file_url: videoUrl,
      video_url: videoUrl,
      access_token: token
    };
    const pubRes = UrlFetchApp.fetch(baseUrl, { method: 'post', payload: publishPayload, muteHttpExceptions: true });
    const pub = JSON.parse(pubRes.getContentText());
    if (pub.error) throw new Error(pub.error.message);
    return pub.id || videoId;
  } catch (e) { throw new Error(`Reels API Error: ${e.message}`); }
}

function callFacebookAPI(url, payload) {
  const token = String(CONFIG.FB_PAGE_TOKEN).replace(/\s+/g, '');
  const finalUrl = `${url}?access_token=${token}`;
  const options = { method: 'post', payload: payload, muteHttpExceptions: true };
  const res = UrlFetchApp.fetch(finalUrl, options);
  const json = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error(json.error?.message || res.getContentText());
  return json.id || json.post_id || json.video_id;
}
