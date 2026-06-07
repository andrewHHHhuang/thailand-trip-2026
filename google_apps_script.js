// Google Apps Script - 部署為 Web App
// 用於泰國行程規劃器的資料同步

// 共享權杖：須與 thailand_planner.html 內的 SYNC_TOKEN 完全相同。
// 修改後請重新部署（部署 → 管理部署作業 → 編輯 → 新版本）才會生效。
const SYNC_TOKEN = 'doraeteam-th2026-7Kq9';

function unauthorized() {
  return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== SYNC_TOKEN) return unauthorized();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('thailand_data');

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Sheet not found'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange('A1').getValue();

  return ContentService.createTextOutput(data || '{}')
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.parameter || e.parameter.token !== SYNC_TOKEN) return unauthorized();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('thailand_data');

  if (!sheet) {
    // 如果 sheet 不存在，建立它
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const newSheet = ss.insertSheet('thailand_data');
    newSheet.getRange('A1').setValue('{}');

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Sheet created and data saved'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const data = e.parameter.data || e.postData.contents;

  if (!data) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'No data provided'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 儲存資料到 A1 儲存格
  sheet.getRange('A1').setValue(data);

  // 儲存最後更新時間到 A2
  sheet.getRange('A2').setValue(new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }));

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Data saved successfully',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}
