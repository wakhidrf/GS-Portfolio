/**
 * onOpen() adalah fungsi khusus di Google Apps Script yang berjalan secara otomatis
 * saat Google Sheet dibuka. Digunakan di sini untuk membuat menu kustom.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Downloads')
      .addItem('Open Download Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * showSidebar() membuat dan menampilkan bilah sisi kustom di Google Sheet.
 * Bilah sisi ini akan berisi tombol untuk memicu unduhan data.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Mover UI')
      .setTitle('Download Filtered Data to Drive')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * downloadFilteredDataToDrive(filterValue) menyaring data dari lembar 'Form Responses 1',
 * membuat Google Sheet baru dengan hasilnya, dan menyimpannya ke folder Drive 'BANK PENGOLAHAN EXCEL'.
 *
 * @param {string} filterValue Nilai yang tepat untuk disaring di kolom C.
 * @returns {string} String JSON yang berisi URL file Drive yang baru atau pesan kesalahan.
 */
function downloadFilteredDataToDrive(filterValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Form Responses 1');

  if (!sourceSheet) {
    return 'Error: Sheet "Form Responses 1" not found. Please ensure it exists.';
  }

  var range = sourceSheet.getDataRange();
  var values = range.getValues();

  if (values.length === 0) {
    return 'No data to filter in "Form Responses 1" sheet.';
  }

  var header = values[0];
  var dataRows = values.slice(1);

  // **** MODIFIKASI: Kolom E adalah indeks 4 (0-indexed). ****
  var columnIndexToFilter = 4; // E adalah kolom ke-5, jadi indeksnya 4.

  var filteredDataRows = dataRows.filter(function(row) {
    return row.length > columnIndexToFilter && String(row[columnIndexToFilter]).trim() === String(filterValue).trim();
  });

  var finalData = [header].concat(filteredDataRows);

  if (finalData.length <= 1) {
    return 'No data found for filter: "' + filterValue + '".';
  }

  // **** MODIFIKASI: Penanganan Folder Drive dan Pembuatan Google Sheet Baru ****
  var folderName = 'BANK PENGOLAHAN EXCEL';
  var folder = DriveApp.getFoldersByName(folderName);
  var targetFolder = null;

  if (folder.hasNext()) {
    targetFolder = folder.next();
  } else {
    return 'Error: Folder "' + folderName + '" not found in your Google Drive. Please create it or change the folder name in the script.';
  }

  // --- START MODIFIKASI NAMA FILE ---
  var academicPeriod = "2024/2025 Genap";
  var surveyTitle = "Kepuasan Mahasiswa Terhadap Layanan Akademik";
  // Membersihkan filterValue agar sesuai untuk nama file (menghapus karakter khusus dan mengganti spasi dengan _)
  var cleanedFilterValue = filterValue.replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_');
  
  var newSpreadsheetName = `${academicPeriod}_${cleanedFilterValue}_${surveyTitle}`;
  // Pastikan nama file tidak terlalu panjang (Google Drive memiliki batasan panjang nama file)
  newSpreadsheetName = newSpreadsheetName.substring(0, 90); 
  // --- END MODIFIKASI NAMA FILE ---

  try {
    // Buat spreadsheet baru di root Drive terlebih dahulu
    var newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);
    var newSheet = newSpreadsheet.getSheets()[0]; // Ambil lembar pertama

    newSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);

    // Pindahkan spreadsheet baru ke folder yang ditentukan
    var file = DriveApp.getFileById(newSpreadsheet.getId());
    targetFolder.addFile(file); // Tambahkan ke folder target
    DriveApp.getRootFolder().removeFile(file); // Hapus dari root Drive

    var fileUrl = newSpreadsheet.getUrl();
    var fileId = newSpreadsheet.getId();

    Logger.log('Generated Google Sheet URL: ' + fileUrl + ', File ID: ' + fileId);

    return JSON.stringify({ url: fileUrl, fileId: fileId });

  } catch (e) {
    return 'Error during data processing or file creation: ' + e.message;
  }
}