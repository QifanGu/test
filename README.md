# test

function saveAttachmentsToDrive() {
  var searchQuery = 'has:attachment'; // 附件筛选条件
  var folderName = 'Gmail Attachments'; // 存储附件的文件夹名称

  var threads = GmailApp.search(searchQuery);
  var folder, isFolderExist;
  var attachmentsCount = 0;

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        if (!isFolderExist) {
          folder = DriveApp.createFolder(folderName);
          isFolderExist = true;
        }
        var attachment = attachments[k];
        var file = folder.createFile(attachment);
        attachmentsCount++;
      }
    }
  }

  Logger.log(attachmentsCount + ' attachments saved to Google Drive');
}
