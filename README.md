# test

function saveCsvAttachmentsToDriveAndSheets() {
  var searchQuery = 'has:attachment filename:csv'; // 附件筛选条件，只提取 CSV 文件
  var folderName = 'Gmail CSV Attachments'; // 存储附件的文件夹名称
  var sheetName = 'CSV Data'; // Google Sheets 中的工作表名称

  var threads = GmailApp.search(searchQuery);
  var folder, isFolderExist;
  var attachmentsCount = 0;

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        if (attachment.getContentType() === 'text/csv') { // 只处理 CSV 文件类型
          if (!isFolderExist) {
            folder = DriveApp.createFolder(folderName);
            isFolderExist = true;
          }
          var csvData = attachment.getDataAsString(); // 获取 CSV 文件内容
          var sheet = createSheetWithData(sheetName, csvData); // 创建 Google Sheets 并写入数据
          var sheetFile = DriveApp.getFileById(sheet.getId()); // 获取 Sheets 文件
          var file = folder.createFile(sheetFile.getBlob()); // 将 Sheets 文件保存到 Drive
          attachmentsCount++;
        }
      }
    }
  }

  Logger.log(attachmentsCount + ' CSV attachments saved to Google Drive and Sheets');
}

function createSheetWithData(sheetName, csvData) {
  var spreadsheet = SpreadsheetApp.create(sheetName); // 创建新的 Google Sheets
  var sheet = spreadsheet.getActiveSheet();
  var data = Utilities.Csv.parse(csvData); // 解析 CSV 数据为二维数组
  var rows = data.length;
  var columns = data[0].length;
  sheet.getRange(1, 1, rows, columns).setValues(data); // 将数据写入 Sheets
  return spreadsheet;
}
