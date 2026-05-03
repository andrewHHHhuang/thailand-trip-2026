// Google Apps Script - 部署為 Web App
// 用於泰國行程規劃器的資料同步

function doGet(e) {
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
