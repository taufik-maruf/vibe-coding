function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SIP')
    .addItem('Generate', 'generateSIP')
    .addToUi();
}

function generateSIP() {
  var templateId = "XXXXX"; 
  var folderId = "XXXXX-zPkJl";   
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate SIP");
  
  // Ambil data sebagai string sesuai tampilan di sheet
  var data = sheet.getDataRange().getDisplayValues();
  
  var headers = data[0];
  var folder = DriveApp.getFolderById(folderId);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[14]; // kolom O
    
    if (status === "Proses") {
      var nama = row[3]; // kolom D
      
      // Buat salinan langsung di folder target
      var copy = DriveApp.getFileById(templateId).makeCopy("SIP 2026 - " + nama, folder);
      var doc = DocumentApp.openById(copy.getId());
      var body = doc.getBody();
      
      // Ganti placeholder
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j].trim();
        var value = row[j] || ""; // langsung string dari getDisplayValues()
        
        if (key) {
          body.replaceText('{{' + key + '}}', value);
        }
      }
      
      doc.saveAndClose();
      
      // Konversi ke PDF
      var pdfBlob = copy.getAs("application/pdf");
      var pdfFile = folder.createFile(pdfBlob);
      pdfFile.setName("SIP 2026 - " + nama + ".pdf");
      
      // Update status & link
      sheet.getRange(i+1, 15).setValue("Selesai"); 
      sheet.getRange(i+1, 16).setValue("https://drive.google.com/file/d/" + pdfFile.getId());
      
      // Delete
      DriveApp.getFileById(copy.getId()).setTrashed(true);
    }
  }
}
