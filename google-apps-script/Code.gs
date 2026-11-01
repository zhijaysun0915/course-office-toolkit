/**
 * 辦公室 Excel 萬能工具箱 — Google Drive 雲端同步與網頁代管腳本 (Google Apps Script)
 *
 * 【用途 1：將整個工具箱網頁直接架設在您的 Google Drive 上】
 * 【用途 2：讓您在家 / 辦公室 / 差旅一鍵將比對與盤點進度直接讀寫至私人 Google Drive】
 *
 * 部署步驟（只需 1 分鐘）：
 * 1. 前往 https://script.google.com/ 建立新專案
 * 2. 將本檔案 (Code.gs) 內容貼上取代預設程式碼
 * 3. 點擊右上角「部署」->「新增部署作業」-> 類型選「網頁應用程式 (Web App)」
 * 4. 執行身分選「我」，誰可以存取選「所有人」，點擊部署並複製網址
 * 5. 將該網址貼回工具箱右上角「☁️ Google Drive 移動辦公設定」即可一鍵雲端同步！
 */

const SNAPSHOT_FILE_NAME = '辦公室萬能工具箱_雲端同步資料庫.json';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load';
  if (action === 'load') {
    const data = loadSnapshotFromDrive();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Course Office Toolkit Cloud API Ready' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = e.postData.contents;
    saveSnapshotToDrive(body);
    return ContentService.createTextOutput(JSON.stringify({ status: 'saved', updatedAt: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSnapshotToDrive(jsonString) {
  const files = DriveApp.getFilesByName(SNAPSHOT_FILE_NAME);
  if (files.hasNext()) {
    const file = files.next();
    file.setContent(jsonString);
  } else {
    DriveApp.createFile(SNAPSHOT_FILE_NAME, jsonString, MimeType.PLAIN_TEXT);
  }
}

function loadSnapshotFromDrive() {
  const files = DriveApp.getFilesByName(SNAPSHOT_FILE_NAME);
  if (files.hasNext()) {
    const file = files.next();
    return JSON.parse(file.getBlob().getDataAsString());
  }
  return { status: 'empty', message: '尚無雲端快照紀錄' };
}
