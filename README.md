# test

function mergeCsvFilesToSheet() {
  var folderName = 'Gmail CSV Attachments'; // Google Drive 中的文件夹名称
  var sheetName = 'Merged CSV'; // 要合并数据的 Google Sheets 工作表名称

  var folder = DriveApp.getFoldersByName(folderName).next(); // 获取文件夹
  var files = folder.getFilesByType(MimeType.CSV); // 获取文件夹中的 CSV 文件

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 获取当前活动的 Google Sheets 文件
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName); // 如果工作表不存在，则创建新的工作表
  }

  var rowData = [];
  while (files.hasNext()) {
    var file = files.next();
    var csvData = file.getBlob().getDataAsString();
    var parsedData = Utilities.Csv.parse(csvData); // 解析 CSV 数据为二维数组
    rowData = rowData.concat(parsedData); // 合并数据
  }

  var range = sheet.getRange(1, 1, rowData.length, rowData[0].length);
  range.setValues(rowData); // 将合并的数据写入工作表
}
