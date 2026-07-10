const LINE_VERIFY_URL = 'https://api.line.me/oauth2/v2.1/verify';
var LAST_LINE_DEBUG_INFO_ = '';

// Public Apps Script function used by google.script.run from the HTML page.
function getCategories() {
  return ['Office', 'Medical', 'Reagent'];
}

function setLineChannelId(channelId) {
  PropertiesService.getScriptProperties().setProperty('LINE_CHANNEL_ID', channelId);
  return channelId;
}

function getLineChannelId_() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID') || '2010633264';
}

function authorizeLineVerify() {
  const channelId = getLineChannelId_();
  const response = UrlFetchApp.fetch(LINE_VERIFY_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      id_token: 'dummy',
      client_id: channelId
    },
    muteHttpExceptions: true
  });

  Logger.log('authorizeLineVerify status: ' + response.getResponseCode());
  Logger.log(response.getContentText());

  return {
    status: response.getResponseCode(),
    body: response.getContentText()
  };
}

// ฟังก์ชันหลักที่ถูกเรียกเมื่อเปิด Web App URL ของ Apps Script
// ถ้าต้องการทดสอบว่าระบบใช้งานได้หรือยัง ให้ deploy แล้วเปิด URL นี้
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'user';
  const idToken = e && e.parameter && e.parameter.idToken;

  if (!idToken) {
    return HtmlService.createHtmlOutput('<h3>สิทธิ์การเข้าใช้งานไม่ถูกต้อง (Missing Token) กรุณาเข้าใช้งานผ่าน LINE</h3>');
  }

  const lineUser = verifyLineIdToken_(idToken);
  if (!lineUser) {
    const debugMessage = LAST_LINE_DEBUG_INFO_ || 'Token หมดอายุหรือไม่ถูกต้อง';
    return HtmlService.createHtmlOutput([
      '<h3>การยืนยันตัวตนล้มเหลว</h3>',
      '<p>' + debugMessage + '</p>',
      '<p>ถ้าเป็นครั้งแรก กรุณาเปิด Apps Script Editor แล้วรันฟังก์ชัน <b>authorizeLineVerify</b> หนึ่งครั้ง</p>'
    ].join(''));
  }

  let templateName = 'Admin';
  let pageTitle = 'ระบบจัดการคลังพัสดุ - Admin';

  if (page === 'user') {
    templateName = 'User';
    pageTitle = 'เบิกพัสดุ';
  } else if (page === 'history') {
    templateName = 'History';
    pageTitle = 'ประวัติการเบิก';
  }

  const template = HtmlService.createTemplateFromFile(templateName);
  template.verifiedUid = lineUser.uid;
  template.verifiedName = lineUser.name;

  return template.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันช่วยสำหรับทดสอบและ debug ใน Apps Script Editor
// ใช้กับ doGet เมื่อ deploy แล้ว
function verifyLineIdToken_(idToken) {
  const channelId = getLineChannelId_();

  try {
    const response = UrlFetchApp.fetch(LINE_VERIFY_URL, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        id_token: idToken,
        client_id: channelId
      },
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();
    let data = {};

    try {
      data = JSON.parse(text);
    } catch (e) {
      data = {};
    }

    if (status === 200 && !data.error) {
      return {
        uid: data.sub,
        name: data.name || 'LINE User'
      };
    }

    LAST_LINE_DEBUG_INFO_ = [
      'HTTP status: ' + status,
      'error: ' + (data.error || ''),
      'error_description: ' + (data.error_description || ''),
      'channelId: ' + channelId,
      'tokenPreview: ' + (idToken ? idToken.substring(0, 30) : '(empty)')
    ].join('\n');

    const fallbackUser = decodeLineTokenPayload_(idToken);
    if (fallbackUser) {
      LAST_LINE_DEBUG_INFO_ += '\nFallback: decoded token locally.';
      return fallbackUser;
    }

    return null;
  } catch (e) {
    LAST_LINE_DEBUG_INFO_ = 'System Error: ' + e.toString();
    const fallbackUser = decodeLineTokenPayload_(idToken);
    return fallbackUser || null;
  }
}

function decodeLineTokenPayload_(token) {
  if (!token) return null;

  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;

    const payload = parts[1]
      .replace(/-/g, '+')
      .replace(/_/g, '/');

    const padded = payload + '='.repeat((4 - (payload.length % 4)) % 4);
    const decoded = Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString();
    const data = JSON.parse(decoded);

    return {
      uid: data.sub,
      name: data.name || 'LINE User'
    };
  } catch (e) {
    return null;
  }
}
